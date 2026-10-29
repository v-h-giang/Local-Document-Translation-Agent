"""20 synthetic XLSX integration cases: content, styling, paths and restart."""


def cell(text=None, *, runs=None, formula=None, number=None, bold=False,
         italic=False, strike=False, color=None, height=None):
    item = {}
    if text is not None:
        item["text"] = text
    if runs is not None:
        item["runs"] = runs
    if formula is not None:
        item["formula"] = formula
    if number is not None:
        item["number"] = number
    item["style"] = {"bold": bold, "italic": italic, "strike": strike,
                     "color": color, "height": height}
    return item


def sheet(name="WBS", cells=None, shapes=None, merged=None):
    return {"name": name, "cells": cells or {}, "shapes": shapes or [],
            "merged": merged or []}


def shape(name, text, *, width=14000, height=4500, runs=None):
    item = {"name": name, "text": text, "width": width, "height": height}
    if runs is not None:
        item["runs"] = runs
    return item


CASES = [
    {"id": "01_basic", "sheets": [sheet(cells={"A1": cell("cập nhật")})],
     "translated": ["WBS!A1"]},
    {"id": "02_batch_dedup", "sheets": [sheet(cells={
        "A1": cell("xác nhận"), "A2": cell("xác nhận"), "B1": cell("màn hình"),
        "B2": cell("đặc tả")})],
     "translated": ["WBS!A1", "WBS!A2", "WBS!B1", "WBS!B2"],
     "require_batch": True},
    {"id": "03_japanese_unchanged", "sheets": [sheet(cells={
        "A1": cell("現行システムの仕様"), "B1": cell("APIの設定"),
        "C1": cell(runs=[{"text": "API", "bold": True},
                           {"text": "の設定", "italic": True}]),
        "D1": cell("WBS更新 / ID管理")})],
     "unchanged": ["WBS!A1", "WBS!B1", "WBS!C1", "WBS!D1"],
     "no_model_calls": True},
    {"id": "04_windows_path", "sheets": [sheet(cells={
        "A1": cell(r"Mở C:\new\test.xlsx trước khi chạy")})],
     "literal": {"WBS!A1": [r"C:\new\test.xlsx"]}},
    {"id": "05_url_query", "sheets": [sheet(cells={
        "A1": cell("Kiểm tra https://example.com/a?q=1&lang=vi trước khi chạy")})],
     "literal": {"WBS!A1": ["https://example.com/a?q=1&lang=vi"]}},
    {"id": "06_unicode_filename", "sheets": [sheet(cells={
        "A1": cell("Mở 審査システム移行_報告書_20261024.xlsx trước")})],
     "literal": {"WBS!A1": ["審査システム移行_報告書_20261024.xlsx"]}},
    {"id": "07_unix_relative_paths", "sheets": [sheet(cells={
        "A1": cell("Sửa /opt/app/config.json trước khi chạy"),
        "A2": cell("Đọc ./reports/daily.xlsx trước")})],
     "literal": {"WBS!A1": ["/opt/app/config.json"],
                 "WBS!A2": ["./reports/daily.xlsx"]}},
    {"id": "08_unc_path", "sheets": [sheet(cells={
        "A1": cell(r"Mở \\server\share\report.xlsx để kiểm tra")})],
     "literal": {"WBS!A1": [r"\\server\share\report.xlsx"]}},
    {"id": "09_mixed_rich", "sheets": [sheet(cells={
        "A1": cell(runs=[
            {"text": "cập nhật ", "bold": True, "height": 14},
            {"text": "xác nhận", "color": 255, "height": 10}])})],
     "translated": ["WBS!A1"], "check_rich": {"WBS!A1": ["bold", "color"]}},
    {"id": "10_strikethrough", "sheets": [sheet(cells={
        "A1": cell(runs=[
            {"text": "cũ ", "strike": True, "color": 16711680},
            {"text": "mới", "bold": True}])})],
     "translated": ["WBS!A1"], "check_rich": {"WBS!A1": ["strike", "bold"]}},
    {"id": "11_multiline_italic", "sheets": [sheet(cells={
        "A1": cell(runs=[
            {"text": "cập nhật\n", "italic": True, "height": 14},
            {"text": "xác nhận", "strike": True, "height": 10}])})],
     "translated": ["WBS!A1"], "contains_newline": ["WBS!A1"],
     "check_rich": {"WBS!A1": ["italic", "strike"]}},
    {"id": "12_formulas_numbers", "sheets": [sheet(cells={
        "A1": cell("kiểm tra"), "B1": cell(formula="=SUM(C1:C2)"),
        "C1": cell(number=5), "C2": cell(number=7)})],
     "translated": ["WBS!A1"], "formulas": {"WBS!B1": "=SUM(C1:C2)"},
     "numbers": {"WBS!C1": 5, "WBS!C2": 7}},
    {"id": "13_merged_format", "sheets": [sheet(cells={
        "A1": cell("cập nhật", bold=True, color=255)}, merged=["A1:B1"])],
     "translated": ["WBS!A1"], "merged": {"WBS": ["A1:B1"]}},
    {"id": "14_shape_text", "sheets": [sheet(shapes=[
        shape("DescriptionBox", "xác nhận", width=16000, height=8000)])],
     "shape_translated": ["WBS!DescriptionBox"]},
    {"id": "15_shape_strike", "sheets": [sheet(shapes=[
        shape("HistoryBox", "cũ mới", width=16000, height=8000,
              runs=[{"text": "cũ ", "strike": True},
                    {"text": "mới", "bold": True}])])],
     "shape_translated": ["WBS!HistoryBox"],
     "shape_rich": {"WBS!HistoryBox": ["strike", "bold"]}},
    {"id": "16_shape_fit_review", "sheets": [sheet(shapes=[
        shape("NarrowBox", "cập nhật hệ thống cần xác nhận ngay",
              width=1300, height=700)])],
     "shape_translated": ["WBS!NarrowBox"], "review_shape": True},
    {"id": "17_multi_sheet_sparse", "sheets": [
        sheet("WBS", {"Z5000": cell("cập nhật")}),
        sheet("審査ソースコード", {"A1": cell("xác nhận"),
                               "B1": cell("次期システム")})],
     "translated": ["WBS!Z5000", "審査ソースコード!A1"],
     "unchanged": ["審査ソースコード!B1"]},
    {"id": "18_quote_filename", "sheets": [sheet(cells={
        "A1": cell('Mở "weekly report 2026.xlsx" trước khi kiểm tra')})],
     "literal": {"WBS!A1": ['"weekly report 2026.xlsx"']}},
    {"id": "19_ja_to_vi", "source": "ja", "target": "vi",
     "sheets": [sheet(cells={"A1": cell("現行システム")})],
     "translated": ["WBS!A1"]},
    {"id": "20_checkpoint_resume", "checkpoint": True,
     "sheets": [sheet(cells={
         **{f"A{i}": cell("cập nhật") for i in range(1, 76)},
         **{f"A{i}": cell("xác nhận") for i in range(76, 82)},
     })], "translated": ["WBS!A1", "WBS!A81"]},
]

SMOKE_IDS = {"01_basic", "05_url_query", "10_strikethrough",
             "14_shape_text", "20_checkpoint_resume"}
