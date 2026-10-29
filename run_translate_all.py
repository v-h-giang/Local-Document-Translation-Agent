import argparse
import json
import re
import sys
import time
from pathlib import Path
from collections import Counter

import ollama

from excel_translate import EXCEL_EXTENSIONS
from spreadsheet_backend import translate_spreadsheet
from technical_literals import mask_markdown_literals
from language_pairs import (validate_pair, language_name, pair_glossary_path,
                            DEFAULT_SOURCE_LANG, DEFAULT_TARGET_LANG)

DEFAULT_MODEL = "translategemma:4b"
MODEL = DEFAULT_MODEL
SOURCE_LANG = DEFAULT_SOURCE_LANG
TARGET_LANG = DEFAULT_TARGET_LANG
NUM_CTX = 4096
MAX_CHARS = 1200
MAX_RETRIES = 2
NUM_PREDICT = 1536
MODEL_KEEP_ALIVE = "30m"
UNLOAD_MODEL_AT_EXIT = True
MODEL_WAS_USED = False

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
DEFAULT_GLOSSARY = BASE_DIR / "glossary.json"
MARKDOWN_EXTENSIONS = {".md", ".markdown"}
SUPPORTED_EXTENSIONS = MARKDOWN_EXTENSIONS | EXCEL_EXTENSIONS


def add_protection(mapping, value):
    token = f"__PROTECTED_{len(mapping):06d}__"
    mapping[token] = value
    return token


def protect_markdown(text):
    mapping = {}

    # Protect YAML front matter.
    if text.startswith("---\n") or text.startswith("---\r\n"):
        m = re.match(
            r"\A---\s*\r?\n.*?\r?\n---\s*(?:\r?\n|$)",
            text,
            flags=re.S,
        )
        if m:
            token = add_protection(mapping, m.group(0))
            text = token + text[m.end():]

    patterns = [
        # Fenced code blocks.
        r"(?ms)^(?P<fence>```|~~~)[^\n]*\n.*?^(?P=fence)[ \t]*(?:\n|$)",
        # HTML comments.
        r"(?s)<!--.*?-->",
        # Markdown table separator rows.
        r"(?m)^[ \t]*\|?[ \t]*:?-{3,}:?[ \t]*(?:\|[ \t]*:?-{3,}:?[ \t]*)+\|?[ \t]*$",
        # Inline code.
        r"`+[^`\n]+`+",
        # Markdown link/image destination inside ](...).
        r"(?<=\]\()[^)]+(?=\))",
        # Reference link definition URL: [id]: https://...
        r"(?m)(?<=\]:[ \t])\S+",
        # Raw URLs.
        r"https?://[^\s<>)\]]+",
        # HTML/XML tags.
        r"</?[A-Za-z][^>\n]*>",
        # Common template/variable expressions.
        r"\{\{.*?\}\}|\$\{.*?\}|%\{.*?\}",
    ]

    for pattern in patterns:
        rx = re.compile(pattern)
        text = rx.sub(
            lambda m: add_protection(mapping, m.group(0)),
            text,
        )

    # URLs, paths and filenames must remain byte-for-byte intact, including
    # names embedded inside otherwise translatable Markdown prose.
    text = mask_markdown_literals(text, lambda value: add_protection(mapping, value))
    return text, mapping


def restore_markdown(text, mapping):
    for token, original in mapping.items():
        text = text.replace(token, original)
    return text


def structure_signature(text):
    return {
        "headings": len(re.findall(r"(?m)^[ \t]*#{1,6}[ \t]+", text)),
        "ul_items": len(re.findall(r"(?m)^[ \t]*[-+*][ \t]+", text)),
        "ol_items": len(re.findall(r"(?m)^[ \t]*\d+[.)][ \t]+", text)),
        "quotes": len(re.findall(r"(?m)^[ \t]*>[ \t]?", text)),
        "pipes": text.count("|"),
    }


def placeholder_ok(source, target):
    src = Counter(re.findall(r"__PROTECTED_\d{6}__", source))
    dst = Counter(re.findall(r"__PROTECTED_\d{6}__", target))
    return src == dst


def split_long_piece(piece, max_chars):
    if len(piece) <= max_chars:
        return [piece]

    # Prefer sentence and newline boundaries.
    units = re.split(r"(?<=[.!?。！？])([ \t]+)|(?<=\n)", piece)
    units = [u for u in units if u is not None and u != ""]

    result = []
    buf = ""

    for unit in units:
        if len(buf) + len(unit) <= max_chars:
            buf += unit
            continue

        if buf:
            result.append(buf)
            buf = ""

        # Rare fallback for one extremely long unit.
        if len(unit) > max_chars:
            words = re.split(r"(\s+)", unit)
            tmp = ""
            for word in words:
                if len(tmp) + len(word) > max_chars and tmp:
                    result.append(tmp)
                    tmp = word
                else:
                    tmp += word
            if tmp:
                buf = tmp
        else:
            buf = unit

    if buf:
        result.append(buf)

    return result


def chunk_text(text, max_chars=MAX_CHARS):
    # Preserve blank-line separators.
    pieces = re.split(r"(\r?\n(?:[ \t]*\r?\n)+)", text)

    normalized = []
    for piece in pieces:
        if len(piece) > max_chars:
            normalized.extend(split_long_piece(piece, max_chars))
        else:
            normalized.append(piece)

    chunks = []
    buf = ""

    for piece in normalized:
        if len(buf) + len(piece) <= max_chars:
            buf += piece
        else:
            if buf:
                chunks.append(buf)
            buf = piece

    if buf:
        chunks.append(buf)

    return chunks


def load_glossary(path):
    if not path:
        return {}

    p = Path(path)
    if not p.exists():
        return {}

    with p.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError(
            "glossary.json must be a JSON object: "
            '{"Vietnamese term": "Japanese term"}'
        )

    return {str(k): str(v) for k, v in data.items()}


def build_prompt(text, glossary, strict=False, mode="markdown"):
    """Build a mode-specific prompt; spreadsheet text is not Markdown."""
    if mode not in {"markdown", "spreadsheet"}:
        raise ValueError(f"Unsupported translation mode: {mode}")

    prompt = (
        f"You are a professional translator from {language_name(SOURCE_LANG)} "
        f"({SOURCE_LANG}) into {language_name(TARGET_LANG)} ({TARGET_LANG}). "
        f"Produce only the {language_name(TARGET_LANG)} translation, "
        "with no commentary.\n"
    )

    if mode == "markdown":
        prompt += (
            "Preserve Markdown structure, headings, list markers, table pipes, "
            "links, line breaks and protected placeholders exactly.\n"
            "Do not remove, duplicate, translate or alter any "
            "__PROTECTED_000000__ token.\n"
        )
    else:
        prompt += (
            "This is spreadsheet cell or drawing text, NOT Markdown. "
            "Translate its content naturally. Do not add explanations, "
            "headings, numbering or bullet markers that were not in the source. "
            "Preserve existing identifiers, numbers and line breaks when relevant. "
            "If the input contains [[SEG_NNN]] identifiers, copy those IDs "
            "verbatim and keep exactly one corresponding line per identifier.\n"
        )

    if glossary:
        prompt += "Use this terminology consistently:\n"
        for source_term, target_term in glossary.items():
            prompt += f'- "{source_term}" -> "{target_term}"\n'

    if strict:
        if mode == "markdown":
            prompt += (
                "Retry: the previous output changed Markdown structure or "
                "protected tokens. Preserve those elements exactly.\n"
            )
        else:
            prompt += (
                "Retry: return a nonempty translation with no added "
                "explanations or formatting.\n"
            )

    # TranslateGemma expects blank lines between task and source text.
    prompt += (
        f"Please translate the following {language_name(SOURCE_LANG)} text "
        f"into {language_name(TARGET_LANG)}:\n"
        "\n\n" + text
    )
    return prompt


def translate_chunk(text, glossary, mode="markdown"):
    """Translate with structural validation ONLY for Markdown.

    Excel/UNO workbook text may legitimately resemble Markdown numbering
    after translation, so Markdown list/heading counters must never reject
    a spreadsheet translation.
    """
    global MODEL_WAS_USED

    if mode not in {"markdown", "spreadsheet"}:
        raise ValueError(f"Unsupported translation mode: {mode}")

    source_sig = structure_signature(text) if mode == "markdown" else None

    for attempt in range(1, MAX_RETRIES + 1):
        MODEL_WAS_USED = True
        started = time.perf_counter()

        response = ollama.chat(
            model=MODEL,
            messages=[
                {
                    "role": "user",
                    "content": build_prompt(
                        text,
                        glossary,
                        strict=(attempt > 1),
                        mode=mode,
                    ),
                }
            ],
            options={
                "num_ctx": NUM_CTX,
                "temperature": 0.1,
                "top_k": 64,
                "top_p": 0.95,
                "num_predict": NUM_PREDICT,
            },
            keep_alive=MODEL_KEEP_ALIVE,
        )

        elapsed = time.perf_counter() - started
        eval_count = getattr(response, "eval_count", None)
        eval_duration = getattr(response, "eval_duration", None)
        load_duration = getattr(response, "load_duration", None)

        perf = f"elapsed={elapsed:.1f}s"
        if eval_count and eval_duration:
            tokens_per_second = eval_count / (eval_duration / 1_000_000_000)
            perf += (
                f", output_tokens={eval_count}, "
                f"speed={tokens_per_second:.2f} tok/s"
            )
        if load_duration:
            perf += f", load={load_duration / 1_000_000_000:.1f}s"
        print(f"  attempt {attempt}/{MAX_RETRIES}: {perf}")

        output = (response.message.content or "").strip()

        # Only Markdown requires its original surrounding whitespace intact.
        if mode == "markdown":
            leading = re.match(r"^\s*", text).group(0)
            trailing = re.search(r"\s*$", text).group(0)
            if leading and not output.startswith(leading):
                output = leading + output
            if trailing and not output.endswith(trailing):
                output += trailing

            tokens_ok = placeholder_ok(text, output)
            target_sig = structure_signature(output)
            structure_ok = target_sig == source_sig

            if output.strip() and tokens_ok and structure_ok:
                return output

            print(
                f"  Retry {attempt}/{MAX_RETRIES}: "
                f"placeholder_ok={tokens_ok}, "
                f"structure_ok={structure_ok}, "
                f"source_structure={source_sig}, "
                f"target_structure={target_sig}",
                file=sys.stderr,
            )
        else:
            # Spreadsheet segments are ordinary text: do not apply
            # Markdown-specific heading/list/table constraints to them.
            if output:
                return output
            print(
                f"  Retry {attempt}/{MAX_RETRIES}: "
                "spreadsheet translation returned empty text",
                file=sys.stderr,
            )

        if attempt < MAX_RETRIES:
            time.sleep(0.5)

    raise RuntimeError(
        "TranslateGemma returned invalid Markdown markup"
        if mode == "markdown"
        else "TranslateGemma returned an empty spreadsheet translation"
    )


def translate_file(
    input_path,
    output_path,
    glossary_path=None,
    max_chars=MAX_CHARS,
):
    source = Path(input_path).read_text(encoding="utf-8-sig")

    protected, mapping = protect_markdown(source)
    chunks = chunk_text(protected, max_chars=max_chars)
    glossary = load_glossary(glossary_path)

    print(f"Model              : {MODEL}")
    print(f"Input              : {input_path}")
    print(f"Output             : {output_path}")
    print(f"Chunks             : {len(chunks)}")
    print(f"Protected elements : {len(mapping)}")

    translated_chunks = []

    for index, chunk in enumerate(chunks, start=1):
        if not chunk.strip():
            translated = chunk
        else:
            print(
                f"[{index}/{len(chunks)}] "
                f"translating {len(chunk)} chars..."
            )
            translated = translate_chunk(chunk, glossary)

        translated_chunks.append(translated)

    translated = "".join(translated_chunks)

    if not placeholder_ok(protected, translated):
        raise RuntimeError(
            "Final protected-token integrity check failed."
        )

    restored = restore_markdown(translated, mapping)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_text(
        restored,
        encoding="utf-8",
        newline="",
    )

    print("")
    print("DONE")
    print(f"Saved: {Path(output_path).resolve()}")


def unload_model():
    """Unload TranslateGemma after the batch so other local agents can use RAM."""
    if not UNLOAD_MODEL_AT_EXIT or not MODEL_WAS_USED:
        return

    try:
        print("")
        print(f"Unloading Ollama model: {MODEL} ...")
        ollama.generate(
            model=MODEL,
            prompt="",
            keep_alive=0,
        )
        print(f"Unloaded: {MODEL}")
    except Exception as exc:
        # Do not turn an unload failure into a failed translation job.
        print(
            f"Warning: could not unload Ollama model {MODEL}: "
            f"{type(exc).__name__}: {exc}",
            file=sys.stderr,
        )


def find_supported_files(input_dir):
    return sorted(
        path
        for path in input_dir.rglob("*")
        if path.is_file()
        and path.suffix.lower() in SUPPORTED_EXTENSIONS
        and not path.name.startswith("~$")
    )


def translate_directory(
    input_dir,
    output_dir,
    glossary_path=None,
    max_chars=MAX_CHARS,
):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    files = find_supported_files(input_dir)

    print("=" * 70)
    print(f"Local Translation Agent: {SOURCE_LANG} -> {TARGET_LANG}")
    print("=" * 70)
    print(f"Input directory  : {input_dir.resolve()}")
    print(f"Output directory : {output_dir.resolve()}")
    print(f"Files found      : {len(files)}")
    print("=" * 70)

    if not files:
        print("")
        print(
            "No supported files found. "
            "Supported: .md, .markdown, .xlsx, .xlsm, .xls, .xlsb"
        )
        return

    success = 0
    failed = 0

    for file_index, input_path in enumerate(files, start=1):
        relative_path = input_path.relative_to(input_dir)
        output_path = output_dir / relative_path

        print("")
        print("=" * 70)
        print(f"[FILE {file_index}/{len(files)}] {relative_path}")
        print("=" * 70)

        try:
            suffix = input_path.suffix.lower()

            if suffix in MARKDOWN_EXTENSIONS:
                translate_file(
                    input_path=input_path,
                    output_path=output_path,
                    glossary_path=glossary_path,
                    max_chars=max_chars,
                )
            elif suffix in EXCEL_EXTENSIONS:
                glossary = load_glossary(glossary_path)

                def excel_translate_func(text):
                    return translate_chunk(
                        text,
                        glossary,
                        mode="spreadsheet",
                    )

                translate_spreadsheet(
                    input_path=input_path,
                    output_path=output_path,
                    translate_func=excel_translate_func,
                    max_chars=max_chars,
                    model=MODEL,
                    glossary=glossary,
                    source_lang=SOURCE_LANG,
                    target_lang=TARGET_LANG,
                )
            else:
                raise RuntimeError(
                    f"Unsupported file type: {suffix}"
                )

            success += 1
        except Exception as exc:
            failed += 1
            print(
                f"FAILED: {relative_path}: "
                f"{type(exc).__name__}: {exc}",
                file=sys.stderr,
            )

    print("")
    print("=" * 70)
    print("BATCH FINISHED")
    print("=" * 70)
    print(f"Success : {success}")
    print(f"Failed  : {failed}")
    print(f"Total   : {len(files)}")
    print(f"Output  : {output_dir.resolve()}")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Translate Markdown and Excel files between configured languages "
            "from input directory to output directory with "
            "TranslateGemma through Ollama."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=str(INPUT_DIR),
        help="Input directory. Default: ./input",
    )
    parser.add_argument(
        "--output-dir",
        default=str(OUTPUT_DIR),
        help="Output directory. Default: ./output",
    )
    parser.add_argument(
        "--from", dest="source_lang", default=DEFAULT_SOURCE_LANG,
        help="Source language (default: vi). Examples: ja, en, zh.",
    )
    parser.add_argument(
        "--to", dest="target_lang", default=DEFAULT_TARGET_LANG,
        help="Target language (default: ja). Examples: vi, en, ja.",
    )
    parser.add_argument(
        "--glossary",
        default=None,
        help="Override directional glossary JSON; otherwise use "
             "./glossaries/<from>-<to>.json if present.",
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=(
            "Ollama model. Default: translategemma:4b. "
            "Use translategemma:12b-it-q4_K_M for slower quality mode."
        ),
    )
    parser.add_argument(
        "--max-chars",
        type=int,
        default=MAX_CHARS,
        help="Maximum source characters per request. Default: 1200.",
    )

    args = parser.parse_args()

    global MODEL, SOURCE_LANG, TARGET_LANG
    MODEL = args.model
    try:
        SOURCE_LANG, TARGET_LANG = validate_pair(
            args.source_lang, args.target_lang
        )
    except ValueError as exc:
        parser.error(str(exc))

    if args.max_chars < 300:
        parser.error("--max-chars must be at least 300")

    try:
        glossary_path = pair_glossary_path(
            BASE_DIR, SOURCE_LANG, TARGET_LANG, args.glossary
        )
    except (ValueError, FileNotFoundError) as exc:
        parser.error(str(exc))
    print(f"Language pair      : {SOURCE_LANG} -> {TARGET_LANG}")
    print(f"Glossary           : {glossary_path or '(none)'}")

    try:
        translate_directory(
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            glossary_path=glossary_path,
            max_chars=args.max_chars,
        )
    finally:
        unload_model()


if __name__ == "__main__":
    main()
 