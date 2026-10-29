"""Protect URLs, filesystem paths and extension-bearing file names.

Protection is structural: protected bytes are NEVER submitted to the model.
This is safer than asking a small translation model to reproduce placeholders.
Uses only Python's standard library; shared by UNO, Excel COM and Markdown.
"""
from __future__ import annotations

import re
from urllib.parse import urlsplit

# More specific forms win when they overlap. Filenames are the last fallback.
_URL = re.compile(
    r"(?i)(?<![\w@])(?:https?|ftps?|file|smb|ssh|git)://[^\s<>'\"\x60]+"
    r"|(?i:(?<![\w@])www\.[^\s<>'\"\x60]+)"
)
_EMAIL = re.compile(
    r"(?i)(?<![\w@])[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}(?![\w@])"
)
_WINDOWS_PATH = re.compile(
    r"(?i)(?<!\w)(?:[A-Z]:[\\/]|\\\\[^\s\\/<>'\"\x60]+[\\/])"
    r"[^\s<>'\"\x60|?*]+"
)
_EXPLICIT_PATH = re.compile(
    r"(?<![\w@])(?:~[\\/]|\.{1,2}[\\/]|/(?!/))"
    r"[^\s<>'\"\x60|?*]+"
)
_RELATIVE_PATH = re.compile(
    r"(?<![\w@])(?:[\w\-.]+[\\/])+(?:[\w\-.]+)"
)
_FILENAME = re.compile(
    r"(?<![\w@./\\])[\w@+\-\u0080-\uffff]+"
    r"\.(?:[A-Za-z][A-Za-z0-9]{0,11})(?![\w])"
)
# Quoted filenames with spaces; protect the filename together with its quotes.
_QUOTED_FILENAME = re.compile(
    r"(?P<q>[\x22\x27\x60])"
    r"(?P<name>[^\r\n\x22\x27\x60]{1,220}"
    r"\.[A-Za-z][A-Za-z0-9]{0,11})"
    r"(?P=q)"
)
_TRAILING_PUNCTUATION = ".,;:!?。，、；：！？"
_TRAILING_CLOSERS = ")]}"


def _trim_span(text, start, end):
    """Drop sentence punctuation, but retain balanced URL parentheses."""
    while end > start and text[end - 1] in _TRAILING_PUNCTUATION:
        end -= 1
    pairs = {")": "(", "]": "[", "}": "{"}
    while end > start and text[end - 1] in _TRAILING_CLOSERS:
        closing = text[end - 1]
        snippet = text[start:end]
        if snippet.count(closing) > snippet.count(pairs[closing]):
            end -= 1
        else:
            break
    return start, end


def protected_spans(text):
    """Sorted, non-overlapping (start, end) offsets into the original text."""
    if not isinstance(text, str):
        raise TypeError("text must be str")
    proposals = []
    patterns = (
        (_URL, 0), (_EMAIL, 1), (_QUOTED_FILENAME, 2),
        (_WINDOWS_PATH, 3), (_EXPLICIT_PATH, 4),
        (_RELATIVE_PATH, 5), (_FILENAME, 6),
    )
    for pattern, priority in patterns:
        for match in pattern.finditer(text):
            begin, end = match.span()
            if priority not in (2,):
                begin, end = _trim_span(text, begin, end)
            if end <= begin:
                continue
            proposals.append((begin, -(end - begin), priority, end))

    # Prefer earliest longest match; a substring of a protected URL/path
    # never becomes a separate filename token.
    proposals.sort()
    selected = []
    cursor = 0
    for begin, _, _, end in proposals:
        if begin < cursor:
            continue
        selected.append((begin, end))
        cursor = end
    return selected


def split_styled_runs(runs):
    """Split arbitrary rich-text runs without losing original style/characters.

    Every returned piece has text, style and protected. A URL/path spanning
    multiple differently styled runs stays protected across ALL its pieces.
    """
    full = "".join(run["text"] for run in runs)
    if not full:
        return []
    spans = protected_spans(full)
    boundaries = sorted(
        {0, len(full), *(x for span in spans for x in span)}
    )
    output = []
    absolute = 0
    for run in runs:
        text = run["text"]
        if not text:
            continue
        segment_end = absolute + len(text)
        cuts = [absolute] + [
            b for b in boundaries if absolute < b < segment_end
        ] + [segment_end]
        for left, right in zip(cuts, cuts[1:]):
            piece = text[left - absolute:right - absolute]
            protected = any(a <= left and right <= b for a, b in spans)
            output.append({
                "text": piece,
                "style": dict(run["style"]),
                "protected": protected,
            })
        absolute = segment_end

    # Edge spaces belong to the original Excel/UNO text, not the model.
    # Translate only the substantive interior of a non-protected piece.
    stable = []
    for item in output:
        if item["protected"] or not item["text"].strip():
            stable.append({**item, "protected": True if not item["text"].strip() else item["protected"]})
            continue
        source = item["text"]
        left = len(source) - len(source.lstrip())
        right = len(source.rstrip())
        if left:
            stable.append({**item, "text": source[:left], "protected": True})
        stable.append({**item, "text": source[left:right]})
        if right < len(source):
            stable.append({**item, "text": source[right:], "protected": True})
    if "".join(p["text"] for p in stable) != full:
        raise RuntimeError("Technical literal segmentation lost source text")
    return stable


def mask_markdown_literals(text, make_token):
    """Replace only newly detected technical spans using caller's token map."""
    pieces = []
    offset = 0
    for begin, end in protected_spans(text):
        pieces.append(text[offset:begin])
        pieces.append(make_token(text[begin:end]))
        offset = end
    pieces.append(text[offset:])
    return "".join(pieces)
