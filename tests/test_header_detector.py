from openpyxl import Workbook

from excel_matcher.excel.header_detector import detect_header
from excel_matcher.excel.parser import parse_workbook


def test_detect_header_after_title_and_blank_rows(tmp_path):
    path = tmp_path / "headers.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单"
    sheet.merge_cells("A1:D1")
    sheet["A1"] = "销售订单导入模板"
    sheet["A3"] = "订单号"
    sheet["B3"] = "客户名称"
    sheet["C3"] = "金额"
    sheet["D3"] = "下单日期"
    sheet["A4"] = "SO001"
    sheet["B4"] = "腾讯"
    sheet["C4"] = 1000
    sheet["D4"] = "2026-05-31"
    workbook.save(path)
    parsed = parse_workbook(path)

    result = detect_header(parsed.sheets[0])

    assert result.title_row == 1
    assert result.header_row == 3
    assert result.data_start_row == 4
    assert result.confidence >= 0.7
    assert result.evidence
