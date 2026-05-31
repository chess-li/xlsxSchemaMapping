from datetime import date

import pytest

from excel_matcher.excel.profiler import profile_columns
from excel_matcher.models import CellData, DataType, HeaderDetectionResult, SheetData


def _sheet_from_rows(rows):
    cells = []
    for row_index, row in enumerate(rows, start=1):
        cells.append(
            [
                CellData(
                    row=row_index,
                    column=column_index,
                    coordinate=f"{column_index}:{row_index}",
                    raw_value=value,
                    display_value=value,
                )
                for column_index, value in enumerate(row, start=1)
            ]
        )
    return SheetData(
        name="订单",
        max_row=len(rows),
        max_column=max(len(row) for row in rows),
        cells=cells,
    )


def test_profile_columns_infers_types_statistics_samples_and_blank_headers():
    sheet = _sheet_from_rows(
        [
            ["客户名称", "金额", "下单日期", "是否付款", None],
            ["腾讯", 100.5, date(2026, 5, 1), "是", "A"],
            ["阿里", 200, date(2026, 5, 2), "否", "B"],
            [None, None, None, None, None],
            ["腾讯", 300, date(2026, 5, 3), "是", "C"],
        ]
    )
    header = HeaderDetectionResult(
        sheet_name="订单",
        header_row=1,
        data_start_row=2,
        confidence=1.0,
    )

    profiles = profile_columns(sheet, header)

    assert [profile.column_name for profile in profiles] == [
        "客户名称",
        "金额",
        "下单日期",
        "是否付款",
        "column_5",
    ]
    assert [profile.data_type for profile in profiles] == [
        DataType.TEXT,
        DataType.NUMBER,
        DataType.DATE,
        DataType.BOOLEAN,
        DataType.TEXT,
    ]
    assert profiles[0].samples == ["腾讯", "阿里", "腾讯"]
    assert profiles[1].null_rate == pytest.approx(0.25)
    assert profiles[0].unique_rate == pytest.approx(2 / 3)
    assert profiles[4].samples == ["A", "B", "C"]
