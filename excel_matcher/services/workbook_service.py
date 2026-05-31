from pathlib import Path
from typing import Any

from excel_matcher.excel.header_detector import detect_header
from excel_matcher.excel.parser import parse_workbook
from excel_matcher.excel.profiler import profile_columns


def analyze_workbook(path: str | Path) -> dict[str, Any]:
    workbook = parse_workbook(path)
    headers = {}
    profiles = {}
    for sheet in workbook.sheets:
        header = detect_header(sheet)
        headers[sheet.name] = header
        profiles[sheet.name] = profile_columns(sheet, header)
    return {
        "workbook": workbook,
        "headers": headers,
        "profiles": profiles,
    }
