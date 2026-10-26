"""LibreOffice UNO worker: sparse scan, contextual translation, checkpoints.

Run with the Python interpreter bundled with LibreOffice. Ollama calls belong
to the parent process; all IPC messages use ASCII-safe JSON lines.
"""
import hashlib
import json
import os
import sys
import time
import traceback
import unicodedata
from translation_batch import translatable
from rich_text_alignment import align_rich_text_runs, RichTextAlignmentError
import uuid
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

CELL_TEXT = uno.Enum("com.sun.star.table.CellContentType", "TEXT")
FONT_PROPERTIES = (
    "CharFontName", "CharHeight", "CharWeight", "CharPosture",
    "CharColor", "CharUnderline", "CharStrikeout", "CharEscapement",
)
MIN_FONT_PT = 8.0
MAX_RICH_RUNS = 300
DEFAULT_CHECKPOINT_INTERVAL = 75
PROTOCOL_VERSION = 2


def send(kind, **values):
    sys.stdout.write(json.dumps({"type": kind, **values}, ensure_ascii=True) + "\n")
    sys.stdout.flush()


def receive():
    line = sys.stdin.readline()
    if not line:
        raise RuntimeError("Translation host closed the UNO IPC channel")
    return json.loads(line)


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def get_prop(obj, name):
    try:
        return getattr(obj, name)
    except Exception:
        return None


def checksum(path):
    hasher = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(block)
    return hasher.hexdigest()


def is_bridge_lost(exc):
    text = (type(exc).__name__ + ": " + str(exc)).lower()
    return any(key in text for key in (
        "disposedexception", "bridge disposed", "binary urp",
        "connection lost", "connection reset", "no connection",
    ))


def _encode_style(value):
    if isinstance(value, (bool, int, float, str)):
        return value
    if hasattr(value, "typeName") and hasattr(value, "value"):
        return {"uno_enum": value.typeName, "value": value.value}
    return None


def _decode_style(value):
    if isinstance(value, dict) and "uno_enum" in value:
        return uno.Enum(value["uno_enum"], value["value"])
    return value


def get_runs(text_obj):
    full = text_obj.getString()
    if not full:
        return []
    runs = []
    try:
        paragraphs = text_obj.createEnumeration()
        while paragraphs.hasMoreElements():
            paragraph = paragraphs.nextElement()
            parts = paragraph.createEnumeration()
            while parts.hasMoreElements():
                part = parts.nextElement()
                fragment = part.getString()
                if not fragment:
                    continue
                style = {}
                for name in FONT_PROPERTIES:
                    value = get_prop(part, name)
                    encoded = _encode_style(value)
                    if encoded is not None:
                        style[name] = encoded
                if runs and runs[-1]["style"] == style:
                    runs[-1]["text"] += fragment
                else:
                    runs.append({"text": fragment, "style": style})
                if len(runs) > MAX_RICH_RUNS:
                    break
            if len(runs) > MAX_RICH_RUNS:
                break
    except Exception:
        runs = []

    # UNO paragraph enumeration may omit breaks or normalize newline
    # separators. Align style portions conservatively, never flatten.
    if not runs:
        cursor = text_obj.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        style = {}
        for name in FONT_PROPERTIES:
            encoded = _encode_style(get_prop(cursor, name))
            if encoded is not None:
                style[name] = encoded
        return [{"text": full, "style": style}]

    return align_rich_text_runs(full, runs)


def set_style(cursor, style):
    for key, value in style.items():
        if key in FONT_PROPERTIES:
            try:
                setattr(cursor, key, _decode_style(value))
            except Exception:
                pass


def replace_text(target, runs):
    text = "".join(run["text"] for run in runs)
    target.setString(text)
    cursor = target.createTextCursor()
    cursor.gotoStart(False)
    for run in runs:
        if not run["text"]:
            continue
        cursor.goRight(len(run["text"]), True)
        set_style(cursor, run["style"])
        cursor.collapseToEnd()


def fit_shape(shape, translated_runs=None):
    try:
        text = shape.getString()
        if not text.strip():
            return "none"
        size = shape.getSize()
        original_width = size.Width
        original_height = size.Height
        cursor = shape.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        base_size = float(get_prop(cursor, "CharHeight") or 10)
        font = max(MIN_FONT_PT, base_size)
        lines = text.split("\n")
        width_units = max((
            sum(1 if ord(char) > 255 else 0.6 for char in line)
            for line in lines
        ), default=0)
        unit_to_mm100 = 35.28
        required_width = width_units * font * unit_to_mm100 * 1.10
        required_height = len(lines) * font * unit_to_mm100 * 1.5
        scale = min(
            1.0,
            original_width / max(required_width, 1),
            original_height / max(required_height, 1),
        )
        new_font = max(MIN_FONT_PT, font * scale)
        if new_font < font:
            scale_factor = new_font / font
            if translated_runs:
                # Keep per-run relative font sizes, bold, strikeout and color.
                start = shape.createTextCursor()
                start.gotoStart(False)
                for run in translated_runs:
                    n = len(run["text"])
                    if not n:
                        continue
                    start.goRight(n, True)
                    original_pt = run["style"].get("CharHeight")
                    if isinstance(original_pt, (int, float)) and original_pt > 0:
                        try:
                            start.CharHeight = max(
                                MIN_FONT_PT, float(original_pt) * scale_factor
                            )
                        except Exception:
                            pass
                    start.collapseToEnd()
            else:
                cursor.CharHeight = float(new_font)
            font = new_font
        new_width = int(width_units * font * unit_to_mm100 * 1.10)
        new_height = int(len(lines) * font * unit_to_mm100 * 1.5)
        resized = False
        if new_width > original_width:
            size.Width = new_width  # Width first.
            resized = True
        if new_height > original_height:
            size.Height = new_height
            resized = True
        did_shrink = new_font < base_size
        if resized:
            shape.setSize(size)
            return "shrunk+resized" if did_shrink else "resized"
        return "shrunk" if did_shrink else "unchanged"
    except Exception as exc:
        if is_bridge_lost(exc):
            raise
        send("log", text=f"Shape fit estimate unavailable: {exc}")
        return "unavailable"


def iter_text_cells(sheet, sheet_index):
    """Query only text-filled ranges; no bounding rectangle or area cap."""
    text_flag = uno.getConstantByName("com.sun.star.sheet.CellFlags.STRING")
    ranges = sheet.queryContentCells(int(text_flag))
    try:
        addresses = ranges.getRangeAddresses()
    except Exception:
        # An empty sheet can return a range collection with no addresses.
        addresses = ()
    for address in addresses:
        for row in range(address.StartRow, address.EndRow + 1):
            for col in range(address.StartColumn, address.EndColumn + 1):
                cell = sheet.getCellByPosition(col, row)
                kind = cell.getType()
                if kind != CELL_TEXT and getattr(kind, "value", "") != "TEXT":
                    continue
                text = cell.getString()
                if text.strip():
                    yield {
                        "id": f"C/{sheet_index}/{col}/{row}",
                        "target": cell,
                        "kind": "cell",
                        "text": text,
                    }


def iter_text_shapes(sheet, sheet_index):
    page = sheet.getDrawPage()

    def walk(shape, chain):
        try:
            if shape.supportsService("com.sun.star.drawing.GroupShape"):
                for index in range(shape.getCount()):
                    yield from walk(shape.getByIndex(index), chain + [index])
                return
            if not hasattr(shape, "getString") or not hasattr(shape, "createTextCursor"):
                return
            text = shape.getString()
            if text.strip():
                yield {
                    "id": f"S/{sheet_index}/" + ".".join(map(str, chain)),
                    "target": shape,
                    "kind": "shape",
                    "text": text,
                }
        except Exception as exc:
            if is_bridge_lost(exc):
                raise
            send("log", text=f"Skipped unsupported shape {chain}: {exc}")

    for index in range(page.getCount()):
        yield from walk(page.getByIndex(index), [index])


def checkpoint_paths(output):
    """Legacy checkpoint path plus atomic manifest path.

    The fixed .checkpoint.xlsx name is retained ONLY for reading checkpoints
    created before versioned storage. New checkpoint workbooks always receive
    a unique filename, never overwritten while LibreOffice has one open.
    """
    return (
        output.with_name(output.stem + ".checkpoint" + output.suffix),
        output.with_name(output.stem + ".checkpoint.json"),
    )


def _checkpoint_file_for_state(output, state):
    legacy_book, _ = checkpoint_paths(output)
    filename = state.get("book_file")
    if filename is None:
        return legacy_book

    # The manifest is local state, but validate it before opening anything:
    # never allow reading outside the workbook's output directory.
    if not isinstance(filename, str):
        raise ValueError("Invalid checkpoint workbook filename")
    prefix = output.stem + ".checkpoint."
    if (
        Path(filename).name != filename
        or not filename.startswith(prefix)
        or not filename.endswith(output.suffix)
        or len(filename) <= len(prefix) + len(output.suffix)
    ):
        raise ValueError("Unsafe checkpoint workbook filename")
    return output.with_name(filename)


def load_checkpoint(source_hash, output, signature):
    _, manifest = checkpoint_paths(output)
    if not manifest.is_file():
        return None
    try:
        with manifest.open("r", encoding="utf-8") as stream:
            state = json.load(stream)
        book = _checkpoint_file_for_state(output, state)
        if (
            state.get("version") != PROTOCOL_VERSION
            or state.get("input_sha256") != source_hash
            or state.get("translation_signature") != signature
            or not book.is_file()
            or state.get("book_sha256") != checksum(book)
        ):
            send("log", text="Checkpoint mismatch; restarting from original.")
            return None

        # Transient key for selecting the exact file to open on resume.
        # Old fixed-name checkpoints remain readable without modification.
        state["_checkpoint_path"] = str(book)
        send(
            "log",
            text=f"Resuming checkpoint: {len(state['processed'])} objects done "
                 f"from {book.name}",
        )
        return state
    except (OSError, ValueError, KeyError, TypeError) as exc:
        send("log", text=f"Checkpoint unusable, restarting: {exc}")
        return None


def _clean_old_checkpoint_books(output, keep=None):
    """Best-effort garbage collection, after a committed manifest update.

    A workbook currently open by soffice may remain locked on Windows. Its
    deletion is safely deferred until LibreOffice closes the document.
    """
    legacy_book, _ = checkpoint_paths(output)
    candidates = [legacy_book]
    candidates.extend(
        output.parent.glob(output.stem + ".checkpoint.*" + output.suffix)
    )
    for candidate in candidates:
        if keep is not None and candidate == keep:
            continue
        try:
            candidate.unlink()
        except FileNotFoundError:
            pass
        except PermissionError:
            # An old checkpoint is still open in the UNO document.
            # Never fail translation because garbage collection was blocked.
            pass


def checkpoint_save(document, output, source_hash, processed, cache, stats, filter_name, signature, review, rich_text_warnings):
    # Never overwrite a checkpoint workbook. On Windows an old checkpoint
    # can be locked for the ENTIRE duration of resume because it is the
    # source file of the open UNO document.
    _, manifest = checkpoint_paths(output)
    token = uuid.uuid4().hex
    book = output.with_name(
        output.stem + ".checkpoint." + token + output.suffix
    )
    tmp_manifest = output.with_name(
        manifest.name + "." + token + ".writing"
    )

    # storeToURL keeps the active document associated with its original
    # loaded path, which must not be replaced while still open.
    document.storeToURL(
        uno.systemPathToFileUrl(str(book)),
        (prop("FilterName", filter_name), prop("Overwrite", False)),
    )
    new_hash = checksum(book)
    state = {
        "version": PROTOCOL_VERSION,
        "input_sha256": source_hash,
        "translation_signature": signature,
        "review": review,
        "rich_text_warnings": rich_text_warnings,
        "book_file": book.name,
        "book_sha256": new_hash,
        "processed": sorted(processed),
        "cache": cache,
        "stats": stats,
    }
    with tmp_manifest.open("w", encoding="utf-8") as stream:
        json.dump(state, stream, ensure_ascii=False)
        stream.flush()
        os.fsync(stream.fileno())

    # Publish the pointer LAST. If the worker is killed before this step,
    # the previous checkpoint manifest and workbook remain consistent.
    os.replace(tmp_manifest, manifest)

    # Only clean after manifest points to the new, fully-written workbook.
    # Locked source checkpoint files are ignored until process shutdown.
    _clean_old_checkpoint_books(output, keep=book)

    send(
        "log",
        text=f"Checkpoint saved: {len(processed)} objects, "
             f"{len(cache)} unique strings cached ({book.name})",
    )


def flush_batch(batch, cache, processed, stats, review, source_lang, target_lang):
    if not batch:
        return
    unique = []
    seen = set()
    for item in batch:
        for run in item["runs"]:
            source = run["text"]
            if not translatable(source, source_lang=source_lang, target_lang=target_lang):
                cache[source] = source
            elif source not in cache and source not in seen:
                unique.append(source)
                seen.add(source)

    if unique:
        send("translate_batch", texts=unique)
        reply = receive()
        if not reply.get("ok"):
            raise RuntimeError("Batch translation failed: " + str(reply.get("error")))
        translations = reply.get("translations")
        if not isinstance(translations, list) or len(translations) != len(unique):
            raise RuntimeError("Batch returned an unexpected number of translations.")
        for original, target in zip(unique, translations):
            if not isinstance(target, str):
                raise RuntimeError("Batch returned a non-string translation.")
            cache[original] = target

    for item in batch:
        runs = [
            {"text": cache[run["text"]], "style": run["style"]}
            for run in item["runs"]
        ]
        # Fail fast rather than publish an incomplete workbook.
        if "".join(x["text"] for x in runs) != item["text"]:
            replace_text(item["target"], runs)
        if item["kind"] == "shape" and "".join(x["text"] for x in runs) != item["text"]:
            shape = item["target"]
            before = shape.getSize()
            before_width, before_height = int(before.Width), int(before.Height)
            font_before = get_prop(shape.createTextCursor(), "CharHeight")
            outcome = fit_shape(shape, runs)
            if "shrunk" in outcome or "resized" in outcome:
                after = shape.getSize()
                review.append({
                    "backend": "libreoffice", "sheet": item.get("sheet", ""),
                    "object_id": item["id"], "object_name": str(get_prop(shape, "Name") or ""),
                    "actions": outcome,
                    "original_font_pt": font_before,
                    "result_font_pt": get_prop(shape.createTextCursor(), "CharHeight"),
                    "original_width": before_width, "result_width": int(after.Width),
                    "original_height": before_height, "result_height": int(after.Height),
                    "note": "UNO shape fitting is approximate; visually review",
                })
            stats["objects"] += 1
            if "shrunk" in outcome:
                stats["shrunk"] += 1
            if "resized" in outcome:
                stats["resized"] += 1
        elif item["kind"] == "cell":
            stats["cells"] += 1
        else:
            stats["objects"] += 1
        processed.add(item["id"])


def connect(port):
    context = uno.getComponentContext()
    resolver = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", context
    )
    address = (
        f"uno:socket,host=127.0.0.1,port={port};urp;"
        "StarOffice.ComponentContext"
    )
    for _ in range(80):
        try:
            remote = resolver.resolve(address)
            return remote.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", remote
            )
        except Exception:
            time.sleep(0.25)
    raise RuntimeError("Could not connect to LibreOffice UNO listener")


def main():
    port = int(sys.argv[1])
    staging = Path(sys.argv[2]).resolve()
    output = Path(sys.argv[3]).resolve()
    original = Path(sys.argv[4]).resolve()
    interval = int(sys.argv[5])
    signature = sys.argv[6]
    source_lang = sys.argv[7]
    target_lang = sys.argv[8]
    if interval < 1:
        raise RuntimeError("Checkpoint interval must be positive.")

    original_hash = checksum(original)
    state = load_checkpoint(original_hash, output, signature)
    checkpoint_book, manifest = checkpoint_paths(output)
    workbook_path = Path(state["_checkpoint_path"]) if state else staging
    processed = set(state["processed"]) if state else set()
    cache = dict(state["cache"]) if state else {}
    review = list(state.get("review", [])) if state else []
    rich_text_warnings = list(state.get("rich_text_warnings", [])) if state else []
    stats = dict(state["stats"]) if state else {
        "cells": 0, "objects": 0, "shrunk": 0, "resized": 0,
    }
    filters = {".xlsx": "Calc MS Excel 2007 XML", ".xls": "MS Excel 97"}
    filter_name = filters[original.suffix.lower()]
    desktop = connect(port)
    document = None

    try:
        document = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(workbook_path)),
            "_blank", 0,
            (
                prop("Hidden", True),
                prop("ReadOnly", False),
                prop("UpdateDocMode", 0),
                prop("MacroExecutionMode", 0),
            ),
        )
        if document is None:
            raise RuntimeError(f"LibreOffice could not open {workbook_path}")

        batch = []
        items_since_checkpoint = 0
        discovered = 0
        skipped = 0

        def flush():
            nonlocal batch, items_since_checkpoint
            if not batch:
                return
            flush_batch(batch, cache, processed, stats, review, source_lang, target_lang)
            items_since_checkpoint += len(batch)
            batch = []
            if items_since_checkpoint >= interval:
                checkpoint_save(
                    document, output, original_hash,
                    processed, cache, stats, filter_name, signature, review,
                    rich_text_warnings
                )
                items_since_checkpoint = 0

        for sheet_index in range(document.Sheets.getCount()):
            sheet = document.Sheets.getByIndex(sheet_index)
            sheet_name = str(sheet.Name)
            send("log", text=f"Scanning sheet: {sheet_name} (sparse text cells)")
            for iterator in (
                iter_text_cells(sheet, sheet_index),
                iter_text_shapes(sheet, sheet_index),
            ):
                for item in iterator:
                    discovered += 1
                    if item["id"] in processed:
                        skipped += 1
                        continue
                    try:
                        item["runs"] = get_runs(item["target"])
                    except RichTextAlignmentError as exc:
                        # Preserve the source cell unchanged. Persist its ID
                        # so a later restart does not repeatedly fail here.
                        warning = {
                            "sheet": sheet_name, "object_id": item["id"],
                            "kind": item["kind"], "source_text": item["text"],
                            "reason": str(exc),
                        }
                        rich_text_warnings.append(warning)
                        processed.add(item["id"])
                        items_since_checkpoint += 1
                        send("log", text=(
                            f"Needs manual review: {sheet_name} {item['id']}: {exc}"
                        ))
                        if items_since_checkpoint >= interval:
                            checkpoint_save(
                                document, output, original_hash,
                                processed, cache, stats, filter_name,
                                signature, review, rich_text_warnings
                            )
                            items_since_checkpoint = 0
                        continue
                    if not item["runs"]:
                        continue
                    item["sheet"] = sheet_name
                    batch.append(item)
                    if len(batch) >= interval:
                        flush()
            flush()

        if items_since_checkpoint:
            checkpoint_save(
                document, output, original_hash,
                processed, cache, stats, filter_name, signature, review,
                rich_text_warnings
            )

        # Publish final workbook only after all cells and shapes complete.
        final_tmp = output.with_name(output.stem + ".complete" + output.suffix)
        document.storeToURL(
            uno.systemPathToFileUrl(str(final_tmp)),
            (prop("FilterName", filter_name), prop("Overwrite", True)),
        )
        os.replace(final_tmp, output)

        # Release the input checkpoint before attempting cleanup; Windows
        # denies deleting or replacing an open LibreOffice workbook.
        document.close(True)
        document = None

        # The final workbook is safely published. Remove manifest first,
        # then stale generation files; cleanup is best-effort.
        try:
            manifest.unlink()
        except FileNotFoundError:
            pass
        except PermissionError as exc:
            send("log", text=f"Checkpoint manifest cleanup deferred: {exc}")
        _clean_old_checkpoint_books(output)

        send("done", review=review, rich_text_warnings=rich_text_warnings, stats={
            **stats, "candidates_scanned": discovered,
            "resumed_skipped": skipped, "unique_strings": len(cache),
        })

        # This worker runs against a dedicated headless LibreOffice profile.
        # Ask the desktop to shut down gracefully so Windows releases profile
        # files such as extensions.pmap before the host removes that folder.
        # The output is already complete; shutdown issues must not change the
        # reported translation result.
        try:
            if not desktop.terminate():
                send("log", text=(
                    "LibreOffice did not terminate immediately; "
                    "temporary profile cleanup may be deferred."
                ))
        except Exception as exc:
            send("log", text=(
                "LibreOffice shutdown warning (output already saved): "
                f"{type(exc).__name__}: {exc}"
            ))

    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                try:
                    document.dispose()
                except Exception:
                    pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        send("error", text=traceback.format_exc())
        sys.exit(1)
