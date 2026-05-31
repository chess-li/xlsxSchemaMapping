from dataclasses import dataclass
from typing import Any

from excel_matcher.models import HeaderDetectionResult, SheetData


BUSINESS_TERMS = {
    "订单",
    "单号",
    "客户",
    "购买方",
    "企业",
    "金额",
    "日期",
    "数量",
    "商品",
    "产品",
    "物料",
    "库存",
    "账户",
    "部门",
    "联系人",
    "电话",
    "发票",
    "order",
    "customer",
    "amount",
    "date",
    "quantity",
    "product",
    "inventory",
    "account",
}


@dataclass(frozen=True)
class RowScore:
    row_number: int
    score: float
    non_empty_ratio: float
    text_ratio: float
    business_hits: int
    next_row_data_shape: bool


def detect_header(sheet: SheetData) -> HeaderDetectionResult:
    scored_rows = [_score_row(sheet, row_number) for row_number in range(1, sheet.max_row + 1)]
    candidates = [row for row in scored_rows if row.non_empty_ratio > 0]
    if not candidates:
        return HeaderDetectionResult(
            sheet_name=sheet.name,
            header_row=1,
            data_start_row=2,
            confidence=0.0,
            evidence=["sheet has no non-empty rows"],
        )

    best = max(candidates, key=lambda row: (row.score, row.business_hits, row.non_empty_ratio))
    data_start_row = _first_non_empty_row_after(sheet, best.row_number) or best.row_number + 1
    title_row = _title_row_before(sheet, best.row_number)
    confidence = min(1.0, best.score / 4.0)
    evidence = [
        f"row {best.row_number} non-empty ratio {best.non_empty_ratio:.2f}",
        f"row {best.row_number} text ratio {best.text_ratio:.2f}",
        f"row {best.row_number} business term hits {best.business_hits}",
    ]
    if best.next_row_data_shape:
        evidence.append(f"row {data_start_row} looks like data")
    if title_row is not None:
        evidence.append(f"row {title_row} looks like title")

    return HeaderDetectionResult(
        sheet_name=sheet.name,
        title_row=title_row,
        header_row=best.row_number,
        data_start_row=data_start_row,
        confidence=confidence,
        evidence=evidence,
    )


def _score_row(sheet: SheetData, row_number: int) -> RowScore:
    values = _row_values(sheet, row_number)
    non_empty = [_normalize(value) for value in values if not _is_blank(value)]
    if not non_empty:
        return RowScore(row_number, 0.0, 0.0, 0.0, 0, False)

    non_empty_ratio = len(non_empty) / max(1, sheet.max_column)
    text_values = [value for value in non_empty if isinstance(value, str)]
    text_ratio = len(text_values) / len(non_empty)
    business_hits = _business_term_hits(non_empty)
    next_row_data_shape = _looks_like_data_row(sheet, row_number + 1, len(non_empty))
    score = (
        non_empty_ratio * 1.2
        + text_ratio * 0.9
        + min(business_hits, 4) * 0.35
        + (0.8 if next_row_data_shape else 0.0)
    )
    return RowScore(
        row_number=row_number,
        score=score,
        non_empty_ratio=non_empty_ratio,
        text_ratio=text_ratio,
        business_hits=business_hits,
        next_row_data_shape=next_row_data_shape,
    )


def _row_values(sheet: SheetData, row_number: int) -> list[Any]:
    if row_number < 1 or row_number > len(sheet.cells):
        return []
    return [cell.display_value for cell in sheet.cells[row_number - 1]]


def _normalize(value: Any) -> Any:
    if isinstance(value, str):
        return value.strip()
    return value


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _business_term_hits(values: list[Any]) -> int:
    text = " ".join(str(value).lower() for value in values)
    return sum(1 for term in BUSINESS_TERMS if term in text)


def _looks_like_data_row(sheet: SheetData, row_number: int, header_width: int) -> bool:
    values = [value for value in _row_values(sheet, row_number) if not _is_blank(value)]
    if not values:
        return False
    comparable_width = len(values) >= max(1, int(header_width * 0.5))
    has_data_type = any(not isinstance(_normalize(value), str) for value in values)
    has_identifier = any(any(char.isdigit() for char in str(value)) for value in values)
    return comparable_width and (has_data_type or has_identifier)


def _first_non_empty_row_after(sheet: SheetData, row_number: int) -> int | None:
    for candidate in range(row_number + 1, sheet.max_row + 1):
        if any(not _is_blank(value) for value in _row_values(sheet, candidate)):
            return candidate
    return None


def _title_row_before(sheet: SheetData, header_row: int) -> int | None:
    for row_number in range(header_row - 1, 0, -1):
        values = [value for value in _row_values(sheet, row_number) if not _is_blank(value)]
        if values:
            return row_number
    return None
