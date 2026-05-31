from pathlib import Path

from excel_matcher.config import Settings
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
    TemplateSignature,
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
                        CellData(
                            row=1,
                            column=1,
                            coordinate="A1",
                            raw_value="标题",
                            display_value="标题",
                        ),
                        CellData(
                            row=1,
                            column=2,
                            coordinate="B1",
                            raw_value=None,
                            display_value="标题",
                            merged_anchor="A1",
                        ),
                    ]
                ],
            )
        ],
    )

    dumped = workbook.model_dump()

    assert dumped["sheets"][0]["hidden_columns"] == ["B"]
    assert dumped["sheets"][0]["cells"][0][1]["merged_anchor"] == "A1"
    assert dumped["sheets"][0]["cells"][0][1]["display_value"] == "标题"


def test_stage_result_preserves_skip_reason():
    stage = StageResult(
        stage="embedding",
        status=StageStatus.SKIPPED,
        reason="model unavailable",
    )

    assert stage.status == StageStatus.SKIPPED
    assert stage.reason == "model unavailable"
    assert stage.candidates == []


def test_mapping_result_marks_review_when_confidence_is_low():
    result = FieldMappingResult(
        excel_field="购买方",
        target_field="customer_name",
        confidence=0.72,
        needs_review=True,
        candidates=[
            FieldMatchCandidate(
                target_field="customer_name",
                score=0.72,
                source="rapidfuzz",
                reason="close text match",
            )
        ],
        stage_results=[],
    )

    assert result.needs_review is True
    assert result.candidates[0].source == "rapidfuzz"


def test_column_profile_uses_supported_data_type_enum():
    profile = ColumnProfile(
        column_name="金额",
        column_index=3,
        data_type=DataType.NUMBER,
        null_rate=0.0,
        unique_rate=1.0,
        samples=[1000, 2000],
        evidence=["all non-empty values are numeric"],
    )

    assert profile.data_type == DataType.NUMBER
    assert profile.evidence == ["all non-empty values are numeric"]


def test_header_detection_result_has_required_rows():
    result = HeaderDetectionResult(
        sheet_name="订单",
        title_row=1,
        header_row=2,
        data_start_row=3,
        confidence=0.9,
        evidence=["row 2 has business terms"],
    )

    assert result.header_row == 2
    assert result.data_start_row == 3
    assert result.evidence == ["row 2 has business terms"]


def test_template_signature_preserves_structure_fields():
    signature = TemplateSignature(
        sheet_name="订单",
        signature="abc123",
        normalized_fields=["order_no", "customer_name", "amount"],
        data_types=[DataType.TEXT, DataType.TEXT, DataType.NUMBER],
    )

    assert signature.sheet_name == "订单"
    assert signature.normalized_fields == ["order_no", "customer_name", "amount"]
    assert signature.data_types == [DataType.TEXT, DataType.TEXT, DataType.NUMBER]


def test_settings_provides_default_paths_and_feature_flags():
    settings = Settings()

    assert settings.data_dir == Path("data")
    assert settings.sqlite_path == Path("data/excel_matcher.db")
    assert settings.faiss_index_path == Path("data/faiss/index.faiss")
    assert settings.faiss_metadata_path == Path("data/faiss/metadata.json")
    assert settings.embedding_model_name == "BAAI/bge-small-zh-v1.5"
    assert settings.embedding_top_k == 5
    assert settings.enable_embedding is False
    assert settings.enable_llm is False
    assert settings.llm_timeout_seconds == 60
    assert settings.review_threshold == 0.75
    assert settings.close_candidate_delta == 0.08
