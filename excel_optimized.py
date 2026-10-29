"""Sparse, cached and resumable Microsoft Excel COM translation pipeline."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path

from excel_translate import (
    _load_excel_com, _iter_text_cells, _extract_cell_runs, _extract_shape_runs,
    _write_cell_runs, _write_shape_runs, _fit_shape_text, _safe_get,
)
from translation_batch import translatable, translate_context_batch, literal_terms_from_glossary
from technical_literals import split_styled_runs
from shape_review import write_shape_review

CHECKPOINT_INTERVAL = 75
CHECKPOINT_VERSION = 1
MSO_GROUP = 6


def digest(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for part in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(part)
    return h.hexdigest()


def checkpoint_paths(output):
    return (
        output.with_name(output.stem + ".checkpoint" + output.suffix),
        output.with_name(output.stem + ".checkpoint.json"),
    )


def load_checkpoint(source_hash, output, signature):
    book, manifest = checkpoint_paths(output)
    if not (book.exists() and manifest.exists()):
        return None
    try:
        state = json.loads(manifest.read_text(encoding="utf-8"))
        valid = (
            state.get("version") == CHECKPOINT_VERSION
            and state.get("source_hash") == source_hash
            and state.get("translation_signature") == signature
            and state.get("checkpoint_hash") == digest(book)
        )
        if not valid:
            print("Excel checkpoint hash/version mismatch: starting over.")
            return None
        print(f"Excel resume: {len(state['done'])} completed items")
        return state
    except (OSError, KeyError, ValueError, TypeError) as exc:
        print(f"Excel checkpoint ignored: {exc}")
        return None


def save_checkpoint(workbook, output, source_hash, done, cache, stats, signature, review):
    book, manifest = checkpoint_paths(output)
    candidate = book.with_name(book.stem + ".writing" + book.suffix)
    manifest_tmp = manifest.with_name(manifest.name + ".writing")

    # SaveCopyAs leaves the active workbook, extension and formulas intact.
    workbook.SaveCopyAs(str(candidate))
    state = {
        "version": CHECKPOINT_VERSION,
        "source_hash": source_hash,
        "translation_signature": signature,
        "review": review,
        "checkpoint_hash": digest(candidate),
        "done": sorted(done),
        "cache": cache,
        "stats": stats,
    }
    manifest_tmp.write_text(
        json.dumps(state, ensure_ascii=False), encoding="utf-8"
    )
    # Order matters: the new manifest references exactly the new copy's hash.
    os.replace(candidate, book)
    os.replace(manifest_tmp, manifest)
    print(f"Excel checkpoint: {len(done)} completed, {len(cache)} cached")


def iter_shapes(shapes, sheet_index, prefix=()):
    for i in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(i)
        address = prefix + (i,)
        if int(_safe_get(shape, "Type", -1)) == MSO_GROUP:
            yield from iter_shapes(shape.GroupItems, sheet_index, address)
            continue
        try:
            frame = shape.TextFrame
            text = frame.Characters().Text
            if isinstance(text, str) and text.strip():
                yield {
                    "sheet": str(_safe_get(shape, "Parent", "") or ""),
                    "id": "S/{}/{}".format(
                        sheet_index, ".".join(map(str, address))
                    ),
                    "kind": "shape",
                    "target": shape,
                    "text": text,
                }
        except Exception:
            # Shapes without text (images/charts/OLE) are not translated.
            continue


def iter_items(workbook):
    for si in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(si)
        print(f"Excel sheet: {sheet.Name}")
        for cell in _iter_text_cells(sheet):
            address = str(cell.Address)
            yield {
                "sheet": str(sheet.Name),
                "id": f"C/{si}/{address}",
                "kind": "cell",
                "target": cell,
                "text": str(cell.Value2),
            }
        yield from iter_shapes(sheet.Shapes, si)


def runs_for(item):
    if item["kind"] == "cell":
        return _extract_cell_runs(item["target"], item["text"])
    return _extract_shape_runs(item["target"], item["text"])


def write_item(item, runs, stats, review):
    target_text = "".join(run["text"] for run in runs)
    if item["kind"] == "cell":
        if target_text != item["text"]:
            _write_cell_runs(item["target"], runs)
        stats["cells"] += 1
    else:
        if target_text != item["text"]:
            shape = item["target"]
            old_width, old_height = float(shape.Width), float(shape.Height)
            old_font = _safe_get(shape.TextFrame.Characters().Font, "Size")
            _write_shape_runs(shape, runs)
            shrunk, resized = _fit_shape_text(shape, runs)
            stats["shrunk"] += int(shrunk)
            stats["resized"] += int(resized)
            if shrunk or resized:
                review.append({
                    "backend": "excel", "sheet": item.get("sheet", ""),
                    "object_id": item["id"],
                    "object_name": str(_safe_get(shape, "Name", "") or ""),
                    "actions": "+".join(x for x, flag in
                                        (("shrunk", shrunk), ("resized", resized)) if flag),
                    "original_font_pt": old_font,
                    "result_font_pt": _safe_get(shape.TextFrame.Characters().Font, "Size"),
                    "original_width": old_width, "result_width": float(shape.Width),
                    "original_height": old_height, "result_height": float(shape.Height),
                    "note": "Review resized/shrunk shape",
                })
        stats["objects"] += 1


def flush(batch, cache, done, stats, translate_func, review, glossary, source_lang, target_lang):
    if not batch:
        return
    unique = []
    seen = set()
    for item in batch:
        item["segments"] = split_styled_runs(item["runs"])
        for run in item["segments"]:
            text = run["text"]
            if run["protected"]:
                continue
            if not translatable(text, literal_terms_from_glossary(glossary), source_lang, target_lang):
                cache[text] = text
            elif text not in cache and text not in seen:
                seen.add(text)
                unique.append(text)

    if unique:
        translated = translate_context_batch(
            unique, translate_func, literal_terms_from_glossary(glossary),
            source_lang, target_lang
        )
        if len(translated) != len(unique):
            raise RuntimeError("COM batch produced wrong number of results")
        cache.update(zip(unique, translated))

    for item in batch:
        translated_runs = [
            {"text": run["text"] if run["protected"] else cache[run["text"]], "style": run["style"]}
            for run in item["segments"]
        ]
        try:
            write_item(item, translated_runs, stats, review)
        except Exception as exc:
            raise RuntimeError(
                f"Excel COM failed writing {item['kind']} "
                f"{item['id']}: {type(exc).__name__}: {exc}"
            ) from exc
        done.add(item["id"])


def translate_excel_optimized(input_path, output_path, translate_func,
                              max_chars=1200, signature="", glossary=None,
                              source_lang="vi", target_lang="ja"):
    """COM equivalent of UNO sparse scan, dedup, contextual batches and resume."""
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    if input_path == output_path:
        raise ValueError("Source and output must be different")
    if not input_path.is_file():
        raise FileNotFoundError(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    glossary = glossary or {}
    source_hash = digest(input_path)
    state = load_checkpoint(source_hash, output_path, signature)
    book_checkpoint, manifest = checkpoint_paths(output_path)
    done = set(state["done"]) if state else set()
    cache = dict(state["cache"]) if state else {}
    review = list(state.get("review", [])) if state else []
    stats = dict(state["stats"]) if state else {
        "cells": 0, "objects": 0, "shrunk": 0, "resized": 0
    }

    pythoncom, win32 = _load_excel_com()
    app = workbook = None
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="jpvn_excel_") as temp:
            # Work only on staged bytes; keep source untouched.
            staging = Path(temp) / ("working" + input_path.suffix)
            shutil.copyfile(book_checkpoint if state else input_path, staging)

            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.ScreenUpdating = False
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass
            workbook = app.Workbooks.Open(
                str(staging), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, AddToMru=False,
            )
            batch = []
            pending_since_save = 0
            resumed = 0

            def save_progress():
                nonlocal batch, pending_since_save
                if not batch:
                    return
                flush(batch, cache, done, stats, translate_func, review, glossary,
                      source_lang, target_lang)
                pending_since_save += len(batch)
                batch = []
                if pending_since_save >= CHECKPOINT_INTERVAL:
                    save_checkpoint(
                        workbook, output_path, source_hash,
                        done, cache, stats, signature, review,
                    )
                    pending_since_save = 0

            for item in iter_items(workbook):
                if item["id"] in done:
                    resumed += 1
                    continue
                item["runs"] = runs_for(item)
                if not item["runs"]:
                    continue
                batch.append(item)
                if len(batch) >= CHECKPOINT_INTERVAL:
                    save_progress()
            save_progress()

            if pending_since_save:
                save_checkpoint(
                    workbook, output_path, source_hash,
                    done, cache, stats, signature, review,
                )

            # Publish only a full workbook, never partial state.
            candidate = output_path.with_name(
                output_path.stem + ".complete" + output_path.suffix
            )
            workbook.SaveCopyAs(str(candidate))
            os.replace(candidate, output_path)
            report = write_shape_review(output_path, review)
            print(f"Shape review report: {report}")
            print(
                f"Excel complete: {stats}; cache={len(cache)}; "
                f"resumed_skipped={resumed}"
            )
            for path in (book_checkpoint, manifest):
                try:
                    path.unlink()
                except FileNotFoundError:
                    pass
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
