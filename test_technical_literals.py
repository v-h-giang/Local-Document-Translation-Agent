"""Regression tests: URL, filesystem path and filenames are never translated."""
import unittest
from technical_literals import protected_spans, split_styled_runs, mask_markdown_literals


class TechnicalLiteralsTests(unittest.TestCase):
    def assertProtected(self, text, expected):
        protected = [text[a:b] for a, b in protected_spans(text)]
        self.assertIn(expected, protected, protected)

    def test_urls(self):
        self.assertProtected("Mở https://example.com/api/v1?q=test&v=2 để xem", "https://example.com/api/v1?q=test&v=2")
        self.assertProtected("Xem (https://example.org/docs/file.xlsx).", "https://example.org/docs/file.xlsx")
        self.assertProtected("URL www.example.jp/help", "www.example.jp/help")

    def test_windows_paths(self):
        self.assertProtected(r"Dùng C:\Users\Public\project\report.xlsx để kiểm tra",
                             r"C:\Users\Public\project\report.xlsx")
        self.assertProtected(r"Mở \\server\share\report.xlsx để xem",
                             r"\\server\share\report.xlsx")

    def test_unix_relative_paths(self):
        self.assertProtected("Đọc /opt/app/config.json trước khi chạy", "/opt/app/config.json")
        self.assertProtected("Sửa ./reports/daily.xlsx sau đó", "./reports/daily.xlsx")
        self.assertProtected("Đọc reports/daily.xlsx trước", "reports/daily.xlsx")

    def test_file_names(self):
        self.assertProtected("Đọc report.xlsx trước khi chạy", "report.xlsx")
        self.assertProtected("Kiểm tra 審査システム移行_報告書_20261024.xlsx trước", "審査システム移行_報告書_20261024.xlsx")
        self.assertProtected('Mở "weekly report 2026.xlsx" rồi dịch',
                             '"weekly report 2026.xlsx"')

    def test_protected_multistyle_spans(self):
        source = [
            {"text": "Mở C:\\new", "style": {"bold": True}},
            {"text": "\\test.xlsx rồi dịch", "style": {"strike": True}},
        ]
        parts = split_styled_runs(source)
        self.assertEqual("".join(x["text"] for x in parts), "".join(x["text"] for x in source))
        self.assertEqual(
            "".join(x["text"] for x in parts if x["protected"]),
            r"C:\new\test.xlsx",
        )
        self.assertTrue(any(x["style"].get("strike") for x in parts))

    def test_markdown_mask(self):
        protected = []
        def token(value):
            protected.append(value)
            return "__PROTECTED_000001__"
        text = mask_markdown_literals("Dịch report.xlsx và mở https://example.com/doc", token)
        self.assertNotIn("report.xlsx", text)
        self.assertNotIn("https://example.com/doc", text)
        self.assertIn("report.xlsx", protected)
        self.assertIn("https://example.com/doc", protected)

    def test_prose_unaffected(self):
        text = "Kiểm tra hệ thống trước khi triển khai."
        self.assertEqual(protected_spans(text), [])


if __name__ == "__main__":
    unittest.main()
