"""Lightweight regression tests for language-pair configuration."""
import json
import tempfile
import unittest
from pathlib import Path

from language_pairs import (
    normalize_language, pair_glossary_path, validate_pair, language_name,
)
from translation_batch import translatable, translate_context_batch


class LanguagePairTests(unittest.TestCase):
    def test_pairs_are_directional(self):
        self.assertEqual(validate_pair("JA", "vi"), ("ja", "vi"))
        with self.assertRaises(ValueError):
            validate_pair("vi", "vi")
        with self.assertRaises(ValueError):
            validate_pair("../vi", "ja")

    def test_names(self):
        self.assertEqual(language_name("ja"), "Japanese")
        self.assertEqual(language_name("zh-CN"), "Chinese")

    def test_glossary_selection(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "glossaries"
            path.mkdir()
            (path / "vi-ja.json").write_text(
                json.dumps({"HTC": "現行システム"}), encoding="utf-8"
            )
            selected = pair_glossary_path(folder, "vi", "ja")
            self.assertEqual(selected.name, "vi-ja.json")
            self.assertIsNone(pair_glossary_path(folder, "ja", "vi"))
            with self.assertRaises(FileNotFoundError):
                pair_glossary_path(folder, "ja", "en", path / "missing.json")

    def test_source_language_filter(self):
        self.assertTrue(translatable("次期システムの変更", source_lang="ja", target_lang="vi"))
        self.assertFalse(translatable("APIの仕様", source_lang="vi", target_lang="ja"))
        self.assertTrue(translatable("cập nhật hệ thống", source_lang="vi", target_lang="ja"))
        self.assertFalse(translatable("API", ("API",), source_lang="en", target_lang="ja"))
        self.assertTrue(translatable("Review API", ("API",), source_lang="en", target_lang="ja"))

    def test_context_prompt_uses_selected_pair(self):
        def fake(text):
            self.assertIn("from ja to vi", text)
            return '[[SEG_000]] "Xác nhận"\n[[SEG_001]] "Cập nhật"'
        actual = translate_context_batch(
            ["確認", "更新"], fake, source_lang="ja", target_lang="vi"
        )
        self.assertEqual(actual, ["Xác nhận", "Cập nhật"])


if __name__ == "__main__":
    unittest.main()
