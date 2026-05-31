# Excel Matcher MVP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the full local Excel intelligent business matching MVP from `milestone_1.md`, including parsing, header detection, profiling, dictionary storage, staged matching, template learning, Streamlit, FastAPI, generated validation samples, and `pytest` verification.

**Architecture:** Build a thin-entry layered Python package. Core Pydantic models and services hold all business behavior; Streamlit and FastAPI call the same services. Matching runs the fixed stage order `Exact -> Dictionary -> RapidFuzz -> Embedding/FAISS -> LLM`, with per-stage diagnostics and skip reasons.

**Tech Stack:** Python 3.11+, Pydantic v2, openpyxl, pandas, RapidFuzz, sentence-transformers, FAISS, SQLite, FastAPI, Streamlit, pytest.

---

## Scope Check

This is a broad MVP, but the subsystems are coupled by one vertical workflow: upload workbook, analyze sheets, match fields, confirm mappings, and learn templates. Keep it as one plan, split into independently reviewable tasks with commits after each task.

## Target File Structure

- Create: `pyproject.toml` for dependencies and test configuration.
- Create: `README.md` with setup, run, and verification commands.
- Create: `.gitignore` to keep generated data, caches, and local databases out of git.
- Create: `excel_matcher/__init__.py` package marker.
- Create: `excel_matcher/config.py` runtime settings and default paths.
- Create: `excel_matcher/models.py` Pydantic contracts shared by services, API, and UI.
- Create: `excel_matcher/excel/parser.py` openpyxl workbook parser.
- Create: `excel_matcher/excel/header_detector.py` heuristic header detector.
- Create: `excel_matcher/excel/profiler.py` column profiler.
- Create: `excel_matcher/storage/schema.py` SQLite schema creation.
- Create: `excel_matcher/storage/default_fields.py` built-in standard fields and aliases.
- Create: `excel_matcher/storage/sqlite_store.py` repository functions.
- Create: `excel_matcher/matcher/normalization.py` field text normalization helpers.
- Create: `excel_matcher/matcher/exact_matcher.py` exact matcher.
- Create: `excel_matcher/matcher/dictionary_matcher.py` dictionary matcher.
- Create: `excel_matcher/matcher/fuzzy_matcher.py` RapidFuzz matcher.
- Create: `excel_matcher/vector/faiss_store.py` embedding model and FAISS index wrapper.
- Create: `excel_matcher/matcher/embedding_matcher.py` Embedding/FAISS stage adapter.
- Create: `excel_matcher/matcher/llm_matcher.py` `codex exec` semantic scoring stage.
- Create: `excel_matcher/matcher/fusion_matcher.py` staged matcher and confidence fusion.
- Create: `excel_matcher/services/workbook_service.py` workbook analysis service.
- Create: `excel_matcher/services/matching_service.py` matching orchestration service.
- Create: `excel_matcher/services/template_service.py` template signature, save, and reuse service.
- Create: `excel_matcher/api/app.py` FastAPI app.
- Create: `excel_matcher/ui/streamlit_app.py` Streamlit UI.
- Create: `tests/sample_generator.py` deterministic validation workbook generator.
- Create: `tests/test_models.py`, `tests/test_parser.py`, `tests/test_header_detector.py`, `tests/test_profiler.py`, `tests/test_storage.py`, `tests/test_matchers.py`, `tests/test_vector_and_llm_skip.py`, `tests/test_services.py`, `tests/test_sample_accuracy.py`, and `tests/test_api.py`.

## Task 1: Project Scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `.gitignore`
- Create: `README.md`
- Create: `excel_matcher/__init__.py`
- Test: `pytest --collect-only`

- [ ] **Step 1: Create package and tooling files**

Create `pyproject.toml`:

```toml
[project]
name = "excel-matcher"
version = "0.1.0"
description = "Local Excel intelligent business matching MVP"
requires-python = ">=3.11"
dependencies = [
  "fastapi>=0.115.0",
  "httpx>=0.27.0",
  "openpyxl>=3.1.5",
  "pandas>=2.2.0",
  "pydantic>=2.8.0",
  "python-multipart>=0.0.9",
  "rapidfuzz>=3.9.0",
  "sentence-transformers>=3.0.0",
  "streamlit>=1.36.0",
  "uvicorn>=0.30.0",
]

[project.optional-dependencies]
faiss = ["faiss-cpu>=1.8.0"]
test = ["pytest>=8.2.0"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["."]
addopts = "-q"
```

Create `.gitignore`:

```gitignore
.DS_Store
.pytest_cache/
__pycache__/
*.pyc
.venv/
data/
tests/fixtures/generated/
```

Create `README.md`:

```markdown
# Excel Matcher MVP

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e ".[test,faiss]"
```

## Test

```bash
pytest
```

## Run API

```bash
uvicorn excel_matcher.api.app:app --reload
```

## Run UI

```bash
streamlit run excel_matcher/ui/streamlit_app.py
```
```

Create `excel_matcher/__init__.py`:

```python
"""Local Excel intelligent business matching MVP."""

__all__ = ["__version__"]

__version__ = "0.1.0"
```

- [ ] **Step 2: Run collection**

Run: `pytest --collect-only`

Expected: command runs and reports no tests collected or importable test modules once they exist.

- [ ] **Step 3: Commit scaffolding**

```bash
git add pyproject.toml .gitignore README.md excel_matcher/__init__.py
git commit -m "chore: scaffold excel matcher project"
```

## Task 2: Shared Models And Config

**Files:**
- Create: `excel_matcher/config.py`
- Create: `excel_matcher/models.py`
- Create: `tests/test_models.py`

- [ ] **Step 1: Write model tests**

Create `tests/test_models.py`:

```python
from excel_matcher.models import (
    CellData,
    ColumnProfile,
    DataType,
    FieldMatchCandidate,
    FieldMappingResult,
    HeaderDetectionResult,
    SheetData,
    StageResult,
    StageStatus,
    WorkbookData,
)


def test_workbook_model_serializes_nested_cells():
    workbook = WorkbookData(
        path="test.xlsx",
        sheets=[
            SheetData(
                name="订单",
                max_row=2,
                max_column=2,
                hidden_columns=["B"],
                merged_ranges=["A1:B1"],
                cells=[
                    [
                        CellData(row=1, column=1, coordinate="A1", raw_value="标题", display_value="标题"),
                        CellData(row=1, column=2, coordinate="B1", raw_value=None, display_value=None, merged_anchor="A1"),
                    ]
                ],
            )
        ],
    )

    dumped = workbook.model_dump()

    assert dumped["sheets"][0]["hidden_columns"] == ["B"]
    assert dumped["sheets"][0]["cells"][0][1]["merged_anchor"] == "A1"


def test_stage_result_preserves_skip_reason():
    stage = StageResult(stage="embedding", status=StageStatus.SKIPPED, reason="model unavailable")

    assert stage.status == StageStatus.SKIPPED
    assert stage.reason == "model unavailable"
    assert stage.candidates == []


def test_mapping_result_marks_review_when_confidence_is_low():
    result = FieldMappingResult(
        excel_field="购买方",
        target_field="customer_name",
        confidence=0.72,
        needs_review=True,
        candidates=[FieldMatchCandidate(target_field="customer_name", score=0.72, source="fuzzy")],
        stage_results=[],
    )

    assert result.needs_review is True
    assert result.candidates[0].source == "fuzzy"


def test_column_profile_uses_supported_data_type_enum():
    profile = ColumnProfile(
        column_name="金额",
        column_index=3,
        data_type=DataType.NUMBER,
        null_rate=0.0,
        unique_rate=1.0,
        samples=[1000, 2000],
    )

    assert profile.data_type == DataType.NUMBER


def test_header_detection_result_has_required_rows():
    result = HeaderDetectionResult(sheet_name="订单", title_row=1, header_row=2, data_start_row=3, confidence=0.9)

    assert result.header_row == 2
    assert result.data_start_row == 3
```

- [ ] **Step 2: Run model tests to verify failure**

Run: `pytest tests/test_models.py -q`

Expected: FAIL with `ModuleNotFoundError` or missing model names.

- [ ] **Step 3: Create config**

Create `excel_matcher/config.py`:

```python
from pathlib import Path
from pydantic import BaseModel, Field


class Settings(BaseModel):
    data_dir: Path = Field(default=Path("data"))
    sqlite_path: Path = Field(default=Path("data/excel_matcher.db"))
    faiss_index_path: Path = Field(default=Path("data/faiss/index.faiss"))
    faiss_metadata_path: Path = Field(default=Path("data/faiss/metadata.json"))
    embedding_model_name: str = "BAAI/bge-small-zh-v1.5"
    embedding_top_k: int = 5
    enable_llm: bool = True
    llm_timeout_seconds: int = 60
    review_threshold: float = 0.75
    close_candidate_delta: float = 0.08
```

- [ ] **Step 4: Create shared models**

Create `excel_matcher/models.py`:

```python
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class DataType(str, Enum):
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    DATE = "DATE"
    BOOLEAN = "BOOLEAN"


class StageStatus(str, Enum):
    COMPLETED = "completed"
    SKIPPED = "skipped"
    FAILED = "failed"


class CellData(BaseModel):
    row: int
    column: int
    coordinate: str
    raw_value: Any = None
    display_value: Any = None
    merged_anchor: str | None = None


class SheetData(BaseModel):
    name: str
    max_row: int
    max_column: int
    hidden_columns: list[str] = Field(default_factory=list)
    merged_ranges: list[str] = Field(default_factory=list)
    cells: list[list[CellData]] = Field(default_factory=list)


class WorkbookData(BaseModel):
    path: str
    sheets: list[SheetData] = Field(default_factory=list)


class HeaderDetectionResult(BaseModel):
    sheet_name: str
    header_row: int
    data_start_row: int
    title_row: int | None = None
    confidence: float
    evidence: list[str] = Field(default_factory=list)


class ColumnProfile(BaseModel):
    column_name: str
    column_index: int
    data_type: DataType = DataType.TEXT
    null_rate: float = 0.0
    unique_rate: float = 0.0
    samples: list[Any] = Field(default_factory=list)
    evidence: list[str] = Field(default_factory=list)


class StandardField(BaseModel):
    key: str
    display_name: str
    domain: str
    description: str = ""
    aliases: list[str] = Field(default_factory=list)


class FieldMatchCandidate(BaseModel):
    target_field: str
    score: float
    source: str
    reason: str = ""


class StageResult(BaseModel):
    stage: str
    status: StageStatus
    reason: str = ""
    candidates: list[FieldMatchCandidate] = Field(default_factory=list)
    diagnostics: dict[str, Any] = Field(default_factory=dict)


class FieldMappingResult(BaseModel):
    excel_field: str
    target_field: str | None = None
    confidence: float = 0.0
    needs_review: bool = True
    candidates: list[FieldMatchCandidate] = Field(default_factory=list)
    stage_results: list[StageResult] = Field(default_factory=list)


class TemplateSignature(BaseModel):
    sheet_name: str
    signature: str
    normalized_fields: list[str]
    data_types: list[DataType] = Field(default_factory=list)
```

- [ ] **Step 5: Run model tests**

Run: `pytest tests/test_models.py -q`

Expected: PASS.

- [ ] **Step 6: Commit models**

```bash
git add excel_matcher/config.py excel_matcher/models.py tests/test_models.py
git commit -m "feat: add shared models and settings"
```

## Task 3: Excel Parser

**Files:**
- Create: `excel_matcher/excel/__init__.py`
- Create: `excel_matcher/excel/parser.py`
- Create: `tests/test_parser.py`

- [ ] **Step 1: Write parser tests**

Create `tests/test_parser.py`:

```python
from pathlib import Path

from openpyxl import Workbook

from excel_matcher.excel.parser import parse_workbook


def make_structured_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "订单"
    ws.merge_cells("A1:C1")
    ws["A1"] = "销售订单导入模板"
    ws["A2"] = "订单号"
    ws["B2"] = "客户名称"
    ws["C2"] = "金额"
    ws["A3"] = "SO001"
    ws["B3"] = "腾讯"
    ws["C3"] = 1000
    ws.column_dimensions["C"].hidden = True
    wb.create_sheet("空白说明")
    wb.save(path)


def test_parse_workbook_preserves_sheets_cells_merges_and_hidden_columns(tmp_path):
    file_path = tmp_path / "test.xlsx"
    make_structured_workbook(file_path)

    workbook = parse_workbook(file_path)

    assert workbook.path == str(file_path)
    assert [sheet.name for sheet in workbook.sheets] == ["订单", "空白说明"]
    first = workbook.sheets[0]
    assert first.hidden_columns == ["C"]
    assert "A1:C1" in first.merged_ranges
    assert first.cells[0][0].display_value == "销售订单导入模板"
    assert first.cells[0][1].merged_anchor == "A1"
    assert first.cells[2][2].raw_value == 1000
```

- [ ] **Step 2: Run parser tests to verify failure**

Run: `pytest tests/test_parser.py -q`

Expected: FAIL because `excel_matcher.excel.parser` does not exist.

- [ ] **Step 3: Implement parser**

Create `excel_matcher/excel/__init__.py`:

```python
"""Excel parsing and analysis modules."""
```

Create `excel_matcher/excel/parser.py`. Implement:

```python
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_matcher.models import CellData, SheetData, WorkbookData


def parse_workbook(path: str | Path) -> WorkbookData:
    workbook_path = Path(path)
    wb = load_workbook(workbook_path, data_only=True)
    sheets: list[SheetData] = []
    for ws in wb.worksheets:
        merged_anchor_by_coordinate = _merged_anchor_map(ws.merged_cells.ranges)
        cells: list[list[CellData]] = []
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
            parsed_row: list[CellData] = []
            for cell in row:
                anchor = merged_anchor_by_coordinate.get(cell.coordinate)
                raw_value: Any = cell.value
                display_value = raw_value
                if anchor and anchor != cell.coordinate:
                    display_value = ws[anchor].value
                parsed_row.append(
                    CellData(
                        row=cell.row,
                        column=cell.column,
                        coordinate=cell.coordinate,
                        raw_value=raw_value,
                        display_value=display_value,
                        merged_anchor=anchor,
                    )
                )
            cells.append(parsed_row)
        hidden_columns = [
            get_column_letter(index)
            for index in range(1, ws.max_column + 1)
            if ws.column_dimensions[get_column_letter(index)].hidden
        ]
        sheets.append(
            SheetData(
                name=ws.title,
                max_row=ws.max_row,
                max_column=ws.max_column,
                hidden_columns=hidden_columns,
                merged_ranges=[str(cell_range) for cell_range in ws.merged_cells.ranges],
                cells=cells,
            )
        )
    return WorkbookData(path=str(workbook_path), sheets=sheets)


def _merged_anchor_map(ranges) -> dict[str, str]:
    anchors: dict[str, str] = {}
    for merged_range in ranges:
        anchor = merged_range.start_cell.coordinate
        for row_index, column_index in merged_range.cells:
            anchors[f"{get_column_letter(column_index)}{row_index}"] = anchor
    return anchors
```

- [ ] **Step 4: Run parser tests**

Run: `pytest tests/test_parser.py -q`

Expected: PASS.

- [ ] **Step 5: Commit parser**

```bash
git add excel_matcher/excel/__init__.py excel_matcher/excel/parser.py tests/test_parser.py
git commit -m "feat: parse excel workbook structure"
```

## Task 4: Header Detection

**Files:**
- Create: `excel_matcher/excel/header_detector.py`
- Create: `tests/test_header_detector.py`

- [ ] **Step 1: Write header detector tests**

Create `tests/test_header_detector.py`:

```python
from excel_matcher.excel.header_detector import detect_header
from excel_matcher.models import CellData, SheetData


def make_sheet(rows):
    cells = []
    for row_index, row_values in enumerate(rows, start=1):
        cells.append(
            [
                CellData(
                    row=row_index,
                    column=column_index,
                    coordinate=f"{column_index}:{row_index}",
                    raw_value=value,
                    display_value=value,
                )
                for column_index, value in enumerate(row_values, start=1)
            ]
        )
    return SheetData(name="订单", max_row=len(rows), max_column=len(rows[0]), cells=cells)


def test_detect_header_after_title_row():
    sheet = make_sheet(
        [
            ["销售订单导入模板", None, None],
            ["订单号", "客户名称", "金额"],
            ["SO001", "腾讯", 1000],
            ["SO002", "阿里", 2000],
        ]
    )

    result = detect_header(sheet)

    assert result.title_row == 1
    assert result.header_row == 2
    assert result.data_start_row == 3
    assert result.confidence >= 0.8


def test_detect_header_skips_blank_rows():
    sheet = make_sheet(
        [
            [None, None, None],
            ["库存模板", None, None],
            [None, None, None],
            ["SKU", "商品名称", "库存数量"],
            ["A001", "键盘", 20],
        ]
    )

    result = detect_header(sheet)

    assert result.header_row == 4
    assert result.data_start_row == 5
```

- [ ] **Step 2: Run header tests to verify failure**

Run: `pytest tests/test_header_detector.py -q`

Expected: FAIL because `detect_header` does not exist.

- [ ] **Step 3: Implement heuristic detector**

Create `excel_matcher/excel/header_detector.py` with:

```python
from excel_matcher.models import HeaderDetectionResult, SheetData

BUSINESS_TERMS = {
    "订单",
    "客户",
    "金额",
    "日期",
    "数量",
    "商品",
    "库存",
    "财务",
    "发票",
    "联系人",
    "SKU",
    "产品",
}


def detect_header(sheet: SheetData) -> HeaderDetectionResult:
    best_row = 1
    best_score = -1.0
    evidence: list[str] = []
    for row_index, row in enumerate(sheet.cells, start=1):
        score, row_evidence = _score_row(sheet, row_index)
        if score > best_score:
            best_score = score
            best_row = row_index
            evidence = row_evidence
    title_row = _find_title_row(sheet, best_row)
    confidence = max(0.0, min(1.0, best_score))
    return HeaderDetectionResult(
        sheet_name=sheet.name,
        title_row=title_row,
        header_row=best_row,
        data_start_row=min(best_row + 1, sheet.max_row),
        confidence=confidence,
        evidence=evidence,
    )


def _score_row(sheet: SheetData, row_index: int) -> tuple[float, list[str]]:
    row = sheet.cells[row_index - 1]
    values = [cell.display_value for cell in row]
    non_empty = [value for value in values if value not in (None, "")]
    if not non_empty:
        return 0.0, ["empty row"]
    non_empty_ratio = len(non_empty) / max(1, len(values))
    text_ratio = sum(isinstance(value, str) for value in non_empty) / len(non_empty)
    term_hits = sum(any(term.lower() in str(value).lower() for term in BUSINESS_TERMS) for value in non_empty)
    term_score = min(0.4, term_hits * 0.12)
    next_data_score = _next_row_data_score(sheet, row_index)
    title_penalty = -0.2 if len(non_empty) == 1 and row_index < sheet.max_row else 0.0
    score = (0.25 * non_empty_ratio) + (0.25 * text_ratio) + term_score + (0.3 * next_data_score) + title_penalty
    return score, [
        f"non_empty_ratio={non_empty_ratio:.2f}",
        f"text_ratio={text_ratio:.2f}",
        f"term_hits={term_hits}",
        f"next_data_score={next_data_score:.2f}",
    ]


def _next_row_data_score(sheet: SheetData, row_index: int) -> float:
    if row_index >= sheet.max_row:
        return 0.0
    values = [cell.display_value for cell in sheet.cells[row_index]]
    non_empty = [value for value in values if value not in (None, "")]
    if not non_empty:
        return 0.0
    non_text_count = sum(not isinstance(value, str) for value in non_empty)
    mixed_data_count = sum(_looks_like_code_or_number(value) for value in non_empty)
    return min(1.0, (non_text_count + mixed_data_count) / max(1, len(non_empty)))


def _looks_like_code_or_number(value) -> bool:
    text = str(value)
    return any(character.isdigit() for character in text)


def _find_title_row(sheet: SheetData, header_row: int) -> int | None:
    for row_index in range(header_row - 1, 0, -1):
        values = [cell.display_value for cell in sheet.cells[row_index - 1]]
        non_empty = [value for value in values if value not in (None, "")]
        if len(non_empty) == 1:
            return row_index
    return None
```

- [ ] **Step 4: Run header tests**

Run: `pytest tests/test_header_detector.py -q`

Expected: PASS.

- [ ] **Step 5: Commit detector**

```bash
git add excel_matcher/excel/header_detector.py tests/test_header_detector.py
git commit -m "feat: detect excel header rows"
```

## Task 5: Column Profiler

**Files:**
- Create: `excel_matcher/excel/profiler.py`
- Create: `tests/test_profiler.py`

- [ ] **Step 1: Write profiler tests**

Create `tests/test_profiler.py`:

```python
from datetime import date

from excel_matcher.excel.profiler import profile_columns
from excel_matcher.models import CellData, DataType, HeaderDetectionResult, SheetData


def make_sheet():
    rows = [
        ["订单号", "客户名称", "金额", "下单日期", "是否付款"],
        ["SO001", "腾讯", 1000, date(2026, 1, 1), "是"],
        ["SO002", "阿里", 2000, date(2026, 1, 2), "否"],
        ["SO003", None, None, "2026-01-03", "Y"],
    ]
    cells = []
    for row_index, row_values in enumerate(rows, start=1):
        cells.append(
            [
                CellData(row=row_index, column=column_index, coordinate=f"{column_index}:{row_index}", raw_value=value, display_value=value)
                for column_index, value in enumerate(row_values, start=1)
            ]
        )
    return SheetData(name="订单", max_row=len(rows), max_column=len(rows[0]), cells=cells)


def test_profile_columns_infers_types_and_rates():
    header = HeaderDetectionResult(sheet_name="订单", header_row=1, data_start_row=2, confidence=0.9)

    profiles = profile_columns(make_sheet(), header)

    by_name = {profile.column_name: profile for profile in profiles}
    assert by_name["客户名称"].data_type == DataType.TEXT
    assert by_name["金额"].data_type == DataType.NUMBER
    assert by_name["下单日期"].data_type == DataType.DATE
    assert by_name["是否付款"].data_type == DataType.BOOLEAN
    assert by_name["客户名称"].null_rate == 1 / 3
    assert by_name["订单号"].unique_rate == 1.0
    assert by_name["客户名称"].samples == ["腾讯", "阿里"]
```

- [ ] **Step 2: Run profiler tests to verify failure**

Run: `pytest tests/test_profiler.py -q`

Expected: FAIL because `profile_columns` does not exist.

- [ ] **Step 3: Implement profiler**

Create `excel_matcher/excel/profiler.py`:

```python
from datetime import date, datetime
from typing import Any

from excel_matcher.models import ColumnProfile, DataType, HeaderDetectionResult, SheetData

BOOLEAN_TEXT = {"是", "否", "y", "n", "yes", "no", "true", "false", "1", "0"}


def profile_columns(sheet: SheetData, header_result: HeaderDetectionResult) -> list[ColumnProfile]:
    header = sheet.cells[header_result.header_row - 1]
    data_rows = sheet.cells[header_result.data_start_row - 1 :]
    profiles: list[ColumnProfile] = []
    for index, header_cell in enumerate(header):
        column_name = str(header_cell.display_value or "").strip()
        if not column_name:
            column_name = f"column_{index + 1}"
        values = [row[index].display_value for row in data_rows if index < len(row)]
        non_null = [value for value in values if value not in (None, "")]
        null_rate = 1.0 - (len(non_null) / max(1, len(values)))
        unique_rate = len({str(value) for value in non_null}) / max(1, len(non_null))
        profiles.append(
            ColumnProfile(
                column_name=column_name,
                column_index=index + 1,
                data_type=_infer_data_type(non_null),
                null_rate=null_rate,
                unique_rate=unique_rate,
                samples=non_null[:5],
                evidence=[f"non_null={len(non_null)}", f"total={len(values)}"],
            )
        )
    return profiles


def _infer_data_type(values: list[Any]) -> DataType:
    if not values:
        return DataType.TEXT
    checks = [
        (DataType.BOOLEAN, _is_boolean),
        (DataType.DATE, _is_date),
        (DataType.NUMBER, _is_number),
    ]
    for data_type, checker in checks:
        hits = sum(checker(value) for value in values)
        if hits / len(values) >= 0.8:
            return data_type
    return DataType.TEXT


def _is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    try:
        float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return False
    return True


def _is_date(value: Any) -> bool:
    if isinstance(value, (date, datetime)):
        return True
    text = str(value)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            datetime.strptime(text, fmt)
        except ValueError:
            continue
        return True
    return False


def _is_boolean(value: Any) -> bool:
    if isinstance(value, bool):
        return True
    return str(value).strip().lower() in BOOLEAN_TEXT
```

- [ ] **Step 4: Run profiler tests**

Run: `pytest tests/test_profiler.py -q`

Expected: PASS.

- [ ] **Step 5: Commit profiler**

```bash
git add excel_matcher/excel/profiler.py tests/test_profiler.py
git commit -m "feat: profile excel columns"
```

## Task 6: SQLite Storage And Built-In Dictionary

**Files:**
- Create: `excel_matcher/storage/__init__.py`
- Create: `excel_matcher/storage/default_fields.py`
- Create: `excel_matcher/storage/schema.py`
- Create: `excel_matcher/storage/sqlite_store.py`
- Create: `tests/test_storage.py`

- [ ] **Step 1: Write storage tests**

Create `tests/test_storage.py`:

```python
from excel_matcher.storage.sqlite_store import SQLiteStore


def test_seed_and_load_standard_fields(tmp_path):
    store = SQLiteStore(tmp_path / "test.db")
    store.initialize()
    store.seed_default_fields()

    fields = store.list_standard_fields()
    customer = next(field for field in fields if field.key == "customer_name")

    assert customer.display_name == "客户名称"
    assert "购买方" in customer.aliases
    assert customer.domain in {"order", "crm", "finance"}


def test_save_and_load_confirmed_template(tmp_path):
    store = SQLiteStore(tmp_path / "test.db")
    store.initialize()
    store.seed_default_fields()

    store.save_template(
        signature="订单号|购买方|金额",
        sheet_name="订单",
        mappings={"购买方": "customer_name", "金额": "amount"},
    )

    template = store.find_template("订单号|购买方|金额")

    assert template is not None
    assert template.mappings["购买方"] == "customer_name"
```

- [ ] **Step 2: Run storage tests to verify failure**

Run: `pytest tests/test_storage.py -q`

Expected: FAIL because `SQLiteStore` does not exist.

- [ ] **Step 3: Add default field data**

Create `excel_matcher/storage/default_fields.py`:

```python
from excel_matcher.models import StandardField

DEFAULT_FIELDS = [
    {"key": "order_no", "display_name": "订单号", "domain": "order", "description": "销售或采购订单编号", "aliases": ["订单编号", "单号", "销售单号"]},
    {"key": "customer_name", "display_name": "客户名称", "domain": "order", "description": "客户、购买方或企业名称", "aliases": ["购买方", "企业名称", "客户名", "公司名称"]},
    {"key": "amount", "display_name": "金额", "domain": "finance", "description": "订单、发票或交易金额", "aliases": ["订单金额", "含税金额", "应收金额", "总金额"]},
    {"key": "order_date", "display_name": "下单日期", "domain": "order", "description": "订单创建或业务发生日期", "aliases": ["订单日期", "日期", "业务日期"]},
    {"key": "product_name", "display_name": "商品名称", "domain": "inventory", "description": "商品或产品名称", "aliases": ["产品名称", "物料名称", "品名"]},
    {"key": "quantity", "display_name": "数量", "domain": "order", "description": "订单或明细数量", "aliases": ["订购数量", "采购数量", "销售数量"]},
    {"key": "sku", "display_name": "SKU", "domain": "inventory", "description": "库存商品编码", "aliases": ["商品编码", "物料编码", "产品编码"]},
    {"key": "inventory_qty", "display_name": "库存数量", "domain": "inventory", "description": "当前可用库存数量", "aliases": ["库存", "现存量", "可用库存"]},
    {"key": "contact_name", "display_name": "联系人", "domain": "crm", "description": "客户联系人姓名", "aliases": ["联系人姓名", "客户联系人"]},
    {"key": "phone", "display_name": "电话", "domain": "crm", "description": "联系电话或手机号", "aliases": ["手机号", "联系电话", "手机"]},
    {"key": "invoice_no", "display_name": "发票号", "domain": "finance", "description": "发票编号", "aliases": ["发票号码", "票据号"]},
    {"key": "paid_status", "display_name": "付款状态", "domain": "finance", "description": "是否付款或结清", "aliases": ["是否付款", "是否支付", "支付状态"]},
    {"key": "account_name", "display_name": "账户名称", "domain": "finance", "description": "银行或财务账户名称", "aliases": ["开户名", "银行账户名"]},
    {"key": "department", "display_name": "部门", "domain": "crm", "description": "部门或组织单元", "aliases": ["所属部门", "业务部门"]},
]


def default_standard_fields() -> list[StandardField]:
    return [StandardField(**field) for field in DEFAULT_FIELDS]
```

- [ ] **Step 4: Add schema and repository**

Create `excel_matcher/storage/schema.py`:

```python
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS standard_fields (
    key TEXT PRIMARY KEY,
    display_name TEXT NOT NULL,
    domain TEXT NOT NULL,
    description TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS field_aliases (
    field_key TEXT NOT NULL,
    alias TEXT NOT NULL,
    PRIMARY KEY (field_key, alias),
    FOREIGN KEY (field_key) REFERENCES standard_fields(key)
);

CREATE TABLE IF NOT EXISTS templates (
    signature TEXT PRIMARY KEY,
    sheet_name TEXT NOT NULL,
    mappings_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
"""
```

Create `excel_matcher/storage/sqlite_store.py` with this public interface:

```python
from dataclasses import dataclass
import json
import sqlite3
from pathlib import Path

from excel_matcher.models import StandardField
from excel_matcher.storage.default_fields import default_standard_fields
from excel_matcher.storage.schema import SCHEMA_SQL


@dataclass(frozen=True)
class TemplateRecord:
    signature: str
    sheet_name: str
    mappings: dict[str, str]


class SQLiteStore:
    def __init__(self, path: str | Path):
        self.path = Path(path)

    def connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        return conn

    def initialize(self) -> None:
        with self.connect() as conn:
            conn.executescript(SCHEMA_SQL)

    def seed_default_fields(self) -> None:
        fields = default_standard_fields()
        with self.connect() as conn:
            for field in fields:
                conn.execute(
                    "INSERT OR REPLACE INTO standard_fields(key, display_name, domain, description) VALUES (?, ?, ?, ?)",
                    (field.key, field.display_name, field.domain, field.description),
                )
                for alias in field.aliases:
                    conn.execute(
                        "INSERT OR IGNORE INTO field_aliases(field_key, alias) VALUES (?, ?)",
                        (field.key, alias),
                    )

    def list_standard_fields(self) -> list[StandardField]:
        with self.connect() as conn:
            rows = conn.execute("SELECT * FROM standard_fields ORDER BY key").fetchall()
            alias_rows = conn.execute("SELECT field_key, alias FROM field_aliases ORDER BY alias").fetchall()
        aliases: dict[str, list[str]] = {}
        for row in alias_rows:
            aliases.setdefault(row["field_key"], []).append(row["alias"])
        return [
            StandardField(
                key=row["key"],
                display_name=row["display_name"],
                domain=row["domain"],
                description=row["description"],
                aliases=aliases.get(row["key"], []),
            )
            for row in rows
        ]

    def save_template(self, signature: str, sheet_name: str, mappings: dict[str, str]) -> None:
        with self.connect() as conn:
            conn.execute(
                """
                INSERT INTO templates(signature, sheet_name, mappings_json, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(signature) DO UPDATE SET
                  sheet_name=excluded.sheet_name,
                  mappings_json=excluded.mappings_json,
                  updated_at=CURRENT_TIMESTAMP
                """,
                (signature, sheet_name, json.dumps(mappings, ensure_ascii=False)),
            )

    def find_template(self, signature: str) -> TemplateRecord | None:
        with self.connect() as conn:
            row = conn.execute("SELECT * FROM templates WHERE signature = ?", (signature,)).fetchone()
        if row is None:
            return None
        return TemplateRecord(
            signature=row["signature"],
            sheet_name=row["sheet_name"],
            mappings=json.loads(row["mappings_json"]),
        )
```

- [ ] **Step 5: Run storage tests**

Run: `pytest tests/test_storage.py -q`

Expected: PASS.

- [ ] **Step 6: Commit storage**

```bash
git add excel_matcher/storage tests/test_storage.py
git commit -m "feat: add sqlite dictionary storage"
```

## Task 7: Exact, Dictionary, And RapidFuzz Matchers

**Files:**
- Create: `excel_matcher/matcher/__init__.py`
- Create: `excel_matcher/matcher/normalization.py`
- Create: `excel_matcher/matcher/exact_matcher.py`
- Create: `excel_matcher/matcher/dictionary_matcher.py`
- Create: `excel_matcher/matcher/fuzzy_matcher.py`
- Create: `tests/test_matchers.py`

- [ ] **Step 1: Write matcher tests**

Create `tests/test_matchers.py`:

```python
from excel_matcher.matcher.dictionary_matcher import DictionaryMatcher
from excel_matcher.matcher.exact_matcher import ExactMatcher
from excel_matcher.matcher.fuzzy_matcher import FuzzyMatcher
from excel_matcher.models import StandardField, StageStatus


FIELDS = [
    StandardField(key="customer_name", display_name="客户名称", domain="order", description="客户或购买方名称", aliases=["购买方", "企业名称"]),
    StandardField(key="amount", display_name="金额", domain="order", description="订单金额", aliases=["订单金额", "含税金额"]),
]


def test_exact_matcher_matches_alias_after_normalization():
    result = ExactMatcher(FIELDS).match(" 购买方 ")

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert result.candidates[0].score == 1.0


def test_dictionary_matcher_uses_aliases():
    result = DictionaryMatcher(FIELDS).match("企业名称")

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"


def test_fuzzy_matcher_scores_near_field_name():
    result = FuzzyMatcher(FIELDS, threshold=70).match("客户名")

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert 0.7 <= result.candidates[0].score <= 1.0
```

- [ ] **Step 2: Run matcher tests to verify failure**

Run: `pytest tests/test_matchers.py -q`

Expected: FAIL because matcher modules do not exist.

- [ ] **Step 3: Implement normalization and matchers**

Create the matcher modules. `normalize_field_name(text)` should strip whitespace, lowercase ASCII text, remove common punctuation including Chinese parentheses, and remove suffixes such as `必填`, `可选`, and `*`. `ExactMatcher.match()` and `DictionaryMatcher.match()` return `StageResult(status=completed)` with candidates or an empty completed result. `FuzzyMatcher.match()` uses `rapidfuzz.fuzz.WRatio` over display names and aliases, normalizes scores to `0.0-1.0`, sorts descending, and keeps up to five candidates.

- [ ] **Step 4: Run matcher tests**

Run: `pytest tests/test_matchers.py -q`

Expected: PASS.

- [ ] **Step 5: Commit text matchers**

```bash
git add excel_matcher/matcher tests/test_matchers.py
git commit -m "feat: add text matching stages"
```

## Task 8: Embedding And FAISS Stage With Skip Behavior

**Files:**
- Create: `excel_matcher/vector/__init__.py`
- Create: `excel_matcher/vector/faiss_store.py`
- Create: `excel_matcher/matcher/embedding_matcher.py`
- Create: `tests/test_vector_and_llm_skip.py`

- [ ] **Step 1: Write embedding skip tests**

Create the embedding portion of `tests/test_vector_and_llm_skip.py`:

```python
from excel_matcher.matcher.embedding_matcher import EmbeddingMatcher
from excel_matcher.models import StageStatus, StandardField


def test_embedding_matcher_skips_when_vector_store_unavailable():
    fields = [StandardField(key="customer_name", display_name="客户名称", domain="order", aliases=["购买方"])]
    matcher = EmbeddingMatcher(fields=fields, vector_store=None)

    result = matcher.match("购买方")

    assert result.status == StageStatus.SKIPPED
    assert "vector store" in result.reason.lower()
    assert result.candidates == []
```

- [ ] **Step 2: Run embedding skip test to verify failure**

Run: `pytest tests/test_vector_and_llm_skip.py::test_embedding_matcher_skips_when_vector_store_unavailable -q`

Expected: FAIL because `EmbeddingMatcher` does not exist.

- [ ] **Step 3: Implement FAISS wrapper and embedding matcher**

Create `excel_matcher/vector/faiss_store.py` with `FaissVectorStore` methods `build(fields)`, `save()`, `load()`, `is_available()`, and `query(text, top_k)`. Import `faiss` and `sentence_transformers` inside methods so tests can run without installed optional packages. Create `excel_matcher/matcher/embedding_matcher.py` so `vector_store=None` returns `StageResult(stage="embedding", status=skipped, reason="vector store is not available")`; otherwise query TopK and return normalized candidates.

- [ ] **Step 4: Run embedding skip test**

Run: `pytest tests/test_vector_and_llm_skip.py::test_embedding_matcher_skips_when_vector_store_unavailable -q`

Expected: PASS.

- [ ] **Step 5: Commit embedding stage**

```bash
git add excel_matcher/vector excel_matcher/matcher/embedding_matcher.py tests/test_vector_and_llm_skip.py
git commit -m "feat: add embedding matcher skip behavior"
```

## Task 9: LLM Semantic Scoring Stage

**Files:**
- Create: `excel_matcher/matcher/llm_matcher.py`
- Modify: `tests/test_vector_and_llm_skip.py`

- [ ] **Step 1: Add LLM tests**

Append to `tests/test_vector_and_llm_skip.py`:

```python
from excel_matcher.matcher.llm_matcher import CodexExecLLMMatcher
from excel_matcher.models import ColumnProfile, DataType, FieldMatchCandidate


def test_llm_matcher_skips_without_candidates():
    matcher = CodexExecLLMMatcher(enabled=True)
    profile = ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT, samples=["腾讯"])

    result = matcher.match(profile, candidates=[])

    assert result.status == StageStatus.SKIPPED
    assert "candidate" in result.reason.lower()


def test_llm_matcher_parses_json_from_runner():
    def runner(prompt: str, timeout: int) -> str:
        assert "购买方" in prompt
        return '{"target_field":"customer_name","semantic_score":0.96,"reason":"购买方表示客户"}'

    matcher = CodexExecLLMMatcher(enabled=True, runner=runner, timeout_seconds=5)
    profile = ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT, samples=["腾讯"])
    candidates = [FieldMatchCandidate(target_field="customer_name", score=0.88, source="embedding")]

    result = matcher.match(profile, candidates=candidates)

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert result.candidates[0].score == 0.96
```

- [ ] **Step 2: Run LLM tests to verify failure**

Run: `pytest tests/test_vector_and_llm_skip.py -q`

Expected: FAIL because `CodexExecLLMMatcher` does not exist.

- [ ] **Step 3: Implement LLM matcher**

Create `excel_matcher/matcher/llm_matcher.py`. The class `CodexExecLLMMatcher` accepts `enabled`, `timeout_seconds`, and an injectable `runner`. If disabled, return skipped with reason `llm scoring is disabled`. If no candidates are passed, return skipped with reason `llm scoring requires candidates`. The default runner calls `subprocess.run(["codex", "exec", prompt], capture_output=True, text=True, timeout=timeout_seconds, check=False)`. Parse JSON from stdout, require `target_field` and `semantic_score`, clamp score to `0.0-1.0`, and return `StageResult(stage="llm", status=completed)`. Invalid JSON, timeout, or non-zero exit returns skipped with a specific reason.

- [ ] **Step 4: Run LLM tests**

Run: `pytest tests/test_vector_and_llm_skip.py -q`

Expected: PASS.

- [ ] **Step 5: Commit LLM scorer**

```bash
git add excel_matcher/matcher/llm_matcher.py tests/test_vector_and_llm_skip.py
git commit -m "feat: add codex exec semantic scorer"
```

## Task 10: Fusion Matcher And Matching Service

**Files:**
- Create: `excel_matcher/matcher/fusion_matcher.py`
- Create: `excel_matcher/services/__init__.py`
- Create: `excel_matcher/services/matching_service.py`
- Create: `tests/test_services.py`

- [ ] **Step 1: Write fusion and service tests**

Create `tests/test_services.py`:

```python
from excel_matcher.matcher.fusion_matcher import FusionMatcher
from excel_matcher.models import ColumnProfile, DataType, StageStatus, StandardField
from excel_matcher.services.matching_service import match_profiles


FIELDS = [
    StandardField(key="customer_name", display_name="客户名称", domain="order", description="客户", aliases=["购买方"]),
    StandardField(key="amount", display_name="金额", domain="order", description="金额", aliases=["订单金额"]),
]


def test_fusion_runs_stages_in_fixed_order_and_keeps_skip_reason():
    matcher = FusionMatcher(standard_fields=FIELDS, vector_store=None, enable_llm=True)
    profile = ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT, samples=["腾讯"])

    result = matcher.match(profile)

    assert [stage.stage for stage in result.stage_results] == ["exact", "dictionary", "rapidfuzz", "embedding", "llm"]
    assert result.target_field == "customer_name"
    assert result.confidence >= 0.95
    embedding = next(stage for stage in result.stage_results if stage.stage == "embedding")
    assert embedding.status == StageStatus.SKIPPED
    assert embedding.reason


def test_match_profiles_returns_one_mapping_per_profile():
    profiles = [
        ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT, samples=["腾讯"]),
        ColumnProfile(column_name="订单金额", column_index=2, data_type=DataType.NUMBER, samples=[1000]),
    ]

    results = match_profiles(profiles, standard_fields=FIELDS, vector_store=None, enable_llm=False)

    assert [result.target_field for result in results] == ["customer_name", "amount"]
```

- [ ] **Step 2: Run fusion tests to verify failure**

Run: `pytest tests/test_services.py -q`

Expected: FAIL because `FusionMatcher` and `match_profiles` do not exist.

- [ ] **Step 3: Implement fusion matcher**

Create `excel_matcher/matcher/fusion_matcher.py`. The matcher must always append stage results in this order: `exact`, `dictionary`, `rapidfuzz`, `embedding`, `llm`. Use exact and dictionary score `1.0` and confidence `0.98`; use RapidFuzz, embedding, and LLM scores directly after normalization. Pick the highest-priority candidate by sorting with priority `exact=5`, `dictionary=4`, `llm=3`, `embedding=2`, `rapidfuzz=1`, then score. Set `needs_review` when confidence is below settings threshold or top two final candidates differ by less than settings close delta.

- [ ] **Step 4: Implement matching service**

Create `excel_matcher/services/__init__.py`:

```python
"""Service orchestration layer."""
```

Create `excel_matcher/services/matching_service.py` with `match_profiles(profiles, standard_fields, vector_store=None, enable_llm=True, settings=None)` that creates one `FusionMatcher` and returns a list of `FieldMappingResult`.

- [ ] **Step 5: Run fusion tests**

Run: `pytest tests/test_services.py -q`

Expected: PASS.

- [ ] **Step 6: Commit fusion**

```bash
git add excel_matcher/matcher/fusion_matcher.py excel_matcher/services tests/test_services.py
git commit -m "feat: fuse staged field matching"
```

## Task 11: Workbook Analysis And Template Learning Services

**Files:**
- Create: `excel_matcher/services/workbook_service.py`
- Create: `excel_matcher/services/template_service.py`
- Modify: `tests/test_services.py`

- [ ] **Step 1: Add service tests**

Append to `tests/test_services.py`:

```python
from excel_matcher.models import HeaderDetectionResult
from excel_matcher.services.template_service import build_template_signature, template_similarity


def test_template_signature_is_stable_for_same_fields():
    profiles = [
        ColumnProfile(column_name="订单号", column_index=1, data_type=DataType.TEXT),
        ColumnProfile(column_name="购买方", column_index=2, data_type=DataType.TEXT),
        ColumnProfile(column_name="金额", column_index=3, data_type=DataType.NUMBER),
    ]

    first = build_template_signature(sheet_name="订单", profiles=profiles)
    second = build_template_signature(sheet_name="导入模板", profiles=profiles)

    assert first.normalized_fields == second.normalized_fields
    assert first.signature == second.signature


def test_template_similarity_scores_identical_structure_high():
    profiles = [
        ColumnProfile(column_name="订单号", column_index=1, data_type=DataType.TEXT),
        ColumnProfile(column_name="购买方", column_index=2, data_type=DataType.TEXT),
    ]
    first = build_template_signature("订单", profiles)
    second = build_template_signature("订单", profiles)

    assert template_similarity(first, second) == 1.0
```

- [ ] **Step 2: Run template tests to verify failure**

Run: `pytest tests/test_services.py -q`

Expected: FAIL because template service functions do not exist.

- [ ] **Step 3: Implement workbook service**

Create `excel_matcher/services/workbook_service.py` with `analyze_workbook(path)` that calls `parse_workbook`, then for each sheet calls `detect_header` and `profile_columns`, and returns a Pydantic model or dictionary with workbook, header results, and profiles keyed by sheet name.

- [ ] **Step 4: Implement template service**

Create `excel_matcher/services/template_service.py` with `build_template_signature(sheet_name, profiles)` and `template_similarity(left, right)`. The signature should ignore sheet name for the hash, normalize field names, include field order, field count, and data type sequence. Similarity should combine exact signature match, Jaccard similarity of normalized fields, and ordered overlap. Identical profiles return `1.0`.

- [ ] **Step 5: Run service tests**

Run: `pytest tests/test_services.py -q`

Expected: PASS.

- [ ] **Step 6: Commit services**

```bash
git add excel_matcher/services/workbook_service.py excel_matcher/services/template_service.py tests/test_services.py
git commit -m "feat: add workbook and template services"
```

## Task 12: Generated Excel Samples And Accuracy Tests

**Files:**
- Create: `tests/sample_generator.py`
- Create: `tests/test_sample_accuracy.py`

- [ ] **Step 1: Write sample accuracy tests**

Create `tests/test_sample_accuracy.py`:

```python
from excel_matcher.excel.header_detector import detect_header
from excel_matcher.excel.parser import parse_workbook
from excel_matcher.excel.profiler import profile_columns
from excel_matcher.services.matching_service import match_profiles
from excel_matcher.storage.default_fields import default_standard_fields
from tests.sample_generator import generate_validation_samples


def test_generated_samples_cover_required_count(tmp_path):
    samples = generate_validation_samples(tmp_path)

    assert len(samples) >= 20
    assert {"order", "crm", "inventory", "finance"}.issubset({sample.domain for sample in samples})


def test_header_detection_accuracy_on_generated_samples(tmp_path):
    samples = generate_validation_samples(tmp_path)
    correct = 0

    for sample in samples:
        workbook = parse_workbook(sample.path)
        sheet = workbook.sheets[0]
        result = detect_header(sheet)
        correct += int(result.header_row == sample.expected_header_row and result.data_start_row == sample.expected_data_start_row)

    assert correct / len(samples) >= 0.80


def test_base_field_matching_accuracy_on_generated_samples(tmp_path):
    samples = generate_validation_samples(tmp_path)
    fields = default_standard_fields()
    total = 0
    correct = 0

    for sample in samples:
        workbook = parse_workbook(sample.path)
        sheet = workbook.sheets[0]
        header = detect_header(sheet)
        profiles = profile_columns(sheet, header)
        results = match_profiles(profiles, fields, vector_store=None, enable_llm=False)
        by_excel_field = {result.excel_field: result.target_field for result in results}
        for excel_field, expected_target in sample.expected_mappings.items():
            total += 1
            correct += int(by_excel_field.get(excel_field) == expected_target)

    assert correct / total >= 0.85
```

- [ ] **Step 2: Run sample tests to verify failure**

Run: `pytest tests/test_sample_accuracy.py -q`

Expected: FAIL because `tests.sample_generator` and `default_standard_fields` do not exist.

- [ ] **Step 3: Implement sample generator**

Create `tests/sample_generator.py` with a `ValidationSample` dataclass and `generate_validation_samples(output_dir)`. Generate at least 20 `.xlsx` files with openpyxl. Cover four domains, merged titles, blank rows, hidden columns, multiple sheets, shifted header rows, changed column order, aliases such as `购买方`, `企业名称`, `订单金额`, `SKU`, `库存数量`, `联系人`, `发票号`, and ambiguous fields. Each sample includes `path`, `domain`, `expected_header_row`, `expected_data_start_row`, and `expected_mappings`.

- [ ] **Step 4: Add default field helper**

In `excel_matcher/storage/default_fields.py`, add `default_standard_fields()` that returns `list[StandardField]` converted from `DEFAULT_FIELDS`.

- [ ] **Step 5: Run sample accuracy tests**

Run: `pytest tests/test_sample_accuracy.py -q`

Expected: PASS with header accuracy at least `0.80` and base field matching accuracy at least `0.85`.

- [ ] **Step 6: Commit validation samples**

```bash
git add tests/sample_generator.py tests/test_sample_accuracy.py excel_matcher/storage/default_fields.py
git commit -m "test: add generated excel validation samples"
```

## Task 13: FastAPI Interface

**Files:**
- Create: `excel_matcher/api/__init__.py`
- Create: `excel_matcher/api/app.py`
- Create: `tests/test_api.py`

- [ ] **Step 1: Write API tests**

Create `tests/test_api.py`:

```python
from fastapi.testclient import TestClient

from excel_matcher.api.app import app


def test_health_reports_components():
    client = TestClient(app)

    response = client.get("/health")

    assert response.status_code == 200
    body = response.json()
    assert "sqlite" in body
    assert "embedding" in body
    assert "llm" in body


def test_fields_endpoint_returns_built_in_dictionary():
    client = TestClient(app)

    response = client.get("/fields")

    assert response.status_code == 200
    body = response.json()
    assert any(field["key"] == "customer_name" for field in body)
```

- [ ] **Step 2: Run API tests to verify failure**

Run: `pytest tests/test_api.py -q`

Expected: FAIL because `excel_matcher.api.app` does not exist.

- [ ] **Step 3: Implement API app**

Create `excel_matcher/api/__init__.py`:

```python
"""FastAPI entry point."""
```

Create `excel_matcher/api/app.py` with a `FastAPI(title="Excel Matcher MVP")` app. Add `GET /health` returning component statuses with `available` and `reason` fields. Add `GET /fields` returning `default_standard_fields()`. Add `POST /workbooks/analyze`, `POST /fields/match`, and `POST /templates` routes using service functions; for file upload, save the upload to a temporary `.xlsx` path before calling `analyze_workbook`.

- [ ] **Step 4: Run API tests**

Run: `pytest tests/test_api.py -q`

Expected: PASS.

- [ ] **Step 5: Commit API**

```bash
git add excel_matcher/api tests/test_api.py
git commit -m "feat: expose fastapi endpoints"
```

## Task 14: Streamlit UI

**Files:**
- Create: `excel_matcher/ui/__init__.py`
- Create: `excel_matcher/ui/streamlit_app.py`
- Modify: `README.md`

- [ ] **Step 1: Create Streamlit package**

Create `excel_matcher/ui/__init__.py`:

```python
"""Streamlit user interface."""
```

- [ ] **Step 2: Implement Streamlit app**

Create `excel_matcher/ui/streamlit_app.py`. It must upload `.xlsx`, save to a temporary file, call `analyze_workbook`, display workbook summary, display detected header and data start rows per sheet, display column profiles, run matching through `match_profiles`, display mapping results with confidence and `needs_review`, show every stage status and skip reason, allow manual target-field selection through `st.selectbox`, and save confirmed mappings through `SQLiteStore.save_template()`.

- [ ] **Step 3: Update run instructions**

Update `README.md` so the UI section includes:

```markdown
The Streamlit app shows each scoring stage in order: Exact, Dictionary, RapidFuzz, Embedding/FAISS, and LLM. If a stage is unavailable, the UI displays the skip reason returned by the service layer.
```

- [ ] **Step 4: Smoke-check imports**

Run: `python -m compileall excel_matcher`

Expected: all Python files compile without syntax errors.

- [ ] **Step 5: Commit UI**

```bash
git add excel_matcher/ui README.md
git commit -m "feat: add streamlit review interface"
```

## Task 15: Full Verification And Polish

**Files:**
- Verify: full repository
- Test: full repository

- [ ] **Step 1: Run full test suite**

Run: `pytest`

Expected: PASS. If this fails, stop the task and create a focused follow-up fix task with the failing test name, the expected behavior, and the exact file paths involved.

- [ ] **Step 2: Run import and entrypoint checks**

Run: `python -m compileall excel_matcher tests`

Expected: all files compile.

Run: `python -c "from excel_matcher.api.app import app; print(app.title)"`

Expected output contains `Excel Matcher MVP`.

- [ ] **Step 3: Check git status**

Run: `git status --short`

Expected: only intentionally untracked local files remain, such as `.DS_Store` or original user-provided files that were not part of implementation commits.

- [ ] **Step 4: Finish verification task**

Run: `git status --short`

Expected: no uncommitted implementation files. Do not create an empty commit for a verification-only pass.
