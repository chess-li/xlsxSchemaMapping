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
    normalized_fields: list[str] = Field(default_factory=list)
    data_types: list[DataType] = Field(default_factory=list)
