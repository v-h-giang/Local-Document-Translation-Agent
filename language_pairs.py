"""Language-pair configuration for TranslateGemma's local translation agent.

Pair-specific glossaries are directional: vi-ja.json is not used for ja-vi.
"""
import re
from pathlib import Path

DEFAULT_SOURCE_LANG = "vi"
DEFAULT_TARGET_LANG = "ja"

LANGUAGE_NAMES = {
    "vi": "Vietnamese",
    "ja": "Japanese",
    "en": "English",
    "zh": "Chinese",
    "ko": "Korean",
    "th": "Thai",
    "id": "Indonesian",
    "ms": "Malay",
    "fr": "French",
    "de": "German",
    "es": "Spanish",
    "pt": "Portuguese",
    "it": "Italian",
    "nl": "Dutch",
    "pl": "Polish",
    "ru": "Russian",
    "uk": "Ukrainian",
    "hi": "Hindi",
    "ar": "Arabic",
    "tr": "Turkish",
    "cs": "Czech",
    "bn": "Bengali",
    "ta": "Tamil",
    "ur": "Urdu",
    "fa": "Persian",
    "sw": "Swahili",
}


def normalize_language(code):
    if not isinstance(code, str):
        raise ValueError("Language code must be a string")
    code = code.strip().replace("_", "-").lower()
    if not re.fullmatch(r"[a-z]{2,3}(?:-[a-z0-9]{2,8})*", code):
        raise ValueError(
            f"Invalid language code {code!r}. "
            "Examples: vi, ja, en, zh, ko, zh-CN."
        )
    return code


def language_name(code):
    code = normalize_language(code)
    return LANGUAGE_NAMES.get(code, LANGUAGE_NAMES.get(code.split("-")[0], code))


def validate_pair(source_lang, target_lang):
    source = normalize_language(source_lang)
    target = normalize_language(target_lang)
    if source == target:
        raise ValueError("Source and target language must be different")
    return source, target


def pair_glossary_path(base_dir, source_lang, target_lang, override=None):
    """Return the optional directional glossary file; do not reuse vi-ja.

    A user-supplied --glossary is mandatory and must exist. A missing automatic
    pair-specific glossary is treated as an empty glossary.
    """
    source, target = validate_pair(source_lang, target_lang)
    if override:
        path = Path(override).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"Glossary not found: {path}")
        return path.resolve()
    path = Path(base_dir) / "glossaries" / f"{source}-{target}.json"
    return path.resolve() if path.is_file() else None
