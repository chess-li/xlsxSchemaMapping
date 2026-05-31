from datetime import date, datetime
from typing import Any

from excel_matcher.models import ColumnProfile, DataType, HeaderDetectionResult, SheetData


BOOLEAN_STRINGS = {"是", "否", "y", "n", "yes", "no", "true", "false"}


def profile_columns(sheet: SheetData, header: HeaderDetectionResult) -> list[ColumnProfile]:
    profiles = []
    header_values = _row_values(sheet, header.header_row)
    data_rows = [
        _row_values(sheet, row_number)
        for row_number in range(header.data_start_row, sheet.max_row + 1)
    ]
    for column_index in range(1, sheet.max_column + 1):
        header_value = _value_at(header_values, column_index)
        column_name = _column_name(header_value, column_index)
        values = [_value_at(row, column_index) for row in data_rows]
        non_null_values = [value for value in values if not _is_blank(value)]
        profiles.append(
            ColumnProfile(
                column_name=column_name,
                column_index=column_index,
                data_type=_infer_data_type(non_null_values),
                null_rate=_null_rate(values),
                unique_rate=_unique_rate(non_null_values),
                samples=non_null_values[:5],
                evidence=[f"profiled {len(values)} data rows"],
            )
        )
    return profiles


def _row_values(sheet: SheetData, row_number: int) -> list[Any]:
    if row_number < 1 or row_number > len(sheet.cells):
        return []
    return [cell.display_value for cell in sheet.cells[row_number - 1]]


def _value_at(values: list[Any], column_index: int) -> Any:
    index = column_index - 1
    if index >= len(values):
        return None
    return values[index]


def _column_name(value: Any, column_index: int) -> str:
    if _is_blank(value):
        return f"column_{column_index}"
    return str(value).strip()


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _null_rate(values: list[Any]) -> float:
    if not values:
        return 0.0
    return sum(1 for value in values if _is_blank(value)) / len(values)


def _unique_rate(values: list[Any]) -> float:
    if not values:
        return 0.0
    return len({str(value) for value in values}) / len(values)


def _infer_data_type(values: list[Any]) -> DataType:
    if not values:
        return DataType.TEXT
    if all(_is_boolean(value) for value in values):
        return DataType.BOOLEAN
    if all(_is_number(value) for value in values):
        return DataType.NUMBER
    if all(_is_date(value) for value in values):
        return DataType.DATE
    return DataType.TEXT


def _is_boolean(value: Any) -> bool:
    if isinstance(value, bool):
        return True
    if isinstance(value, str):
        return value.strip().lower() in BOOLEAN_STRINGS
    return False


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_date(value: Any) -> bool:
    if isinstance(value, (date, datetime)) and not isinstance(value, bool):
        return True
    if not isinstance(value, str):
        return False
    text = value.strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            pass
    return False
