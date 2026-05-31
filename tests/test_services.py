from openpyxl import Workbook

from excel_matcher.config import Settings
from excel_matcher.matcher.fusion_matcher import FusionMatcher
from excel_matcher.models import ColumnProfile, DataType, StageStatus, StandardField
from excel_matcher.services.matching_service import match_profiles
from excel_matcher.services.template_service import (
    build_template_signature,
    template_similarity,
)
from excel_matcher.services.workbook_service import analyze_workbook


FIELDS = [
    StandardField(
        key="customer_name",
        display_name="客户名称",
        domain="order",
        aliases=["购买方", "企业名称"],
    ),
    StandardField(
        key="amount",
        display_name="金额",
        domain="order",
        aliases=["订单金额", "含税金额"],
    ),
    StandardField(
        key="inventory_qty",
        display_name="库存数量",
        domain="inventory",
        aliases=["库存", "现存量"],
    ),
]


def test_phase1_fusion_runs_deterministic_stages_and_marks_semantic_disabled():
    matcher = FusionMatcher(
        standard_fields=FIELDS,
        enable_embedding=False,
        enable_llm=False,
        settings=Settings(review_threshold=0.75),
    )
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯"],
    )

    result = matcher.match(profile)

    assert [stage.stage for stage in result.stage_results] == [
        "exact",
        "dictionary",
        "rapidfuzz",
        "embedding",
        "llm",
    ]
    assert result.target_field == "customer_name"
    assert result.confidence == 0.98
    assert result.needs_review is False
    assert result.stage_results[3].status == StageStatus.SKIPPED
    assert result.stage_results[3].reason == "embedding semantic scoring is disabled in phase 1"
    assert result.stage_results[4].reason == "llm semantic scoring is disabled in phase 1"


def test_phase1_matching_does_not_use_data_content_semantics():
    matcher = FusionMatcher(standard_fields=FIELDS, enable_embedding=False, enable_llm=False)
    profile = ColumnProfile(
        column_name="金额",
        column_index=2,
        data_type=DataType.NUMBER,
        samples=["腾讯", "阿里"],
    )

    result = matcher.match(profile)

    assert result.target_field == "amount"
    assert result.candidates[0].source in {"exact", "dictionary"}


def test_match_profiles_returns_one_mapping_per_profile():
    profiles = [
        ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT),
        ColumnProfile(column_name="订单金额", column_index=2, data_type=DataType.NUMBER),
    ]

    results = match_profiles(profiles, FIELDS, enable_embedding=False, enable_llm=False)

    assert [result.excel_field for result in results] == ["购买方", "订单金额"]
    assert [result.target_field for result in results] == ["customer_name", "amount"]


def test_analyze_workbook_returns_workbook_headers_and_profiles_by_sheet(tmp_path):
    path = tmp_path / "analysis.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单"
    sheet["A1"] = "客户名称"
    sheet["B1"] = "金额"
    sheet["A2"] = "腾讯"
    sheet["B2"] = 1000
    workbook.save(path)

    analysis = analyze_workbook(path)

    assert analysis["workbook"].sheets[0].name == "订单"
    assert analysis["headers"]["订单"].header_row == 1
    assert [profile.column_name for profile in analysis["profiles"]["订单"]] == [
        "客户名称",
        "金额",
    ]


def test_template_signature_is_stable_for_same_fields():
    profiles = [
        ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT),
        ColumnProfile(column_name="金额", column_index=2, data_type=DataType.NUMBER),
    ]

    left = build_template_signature("订单", profiles)
    right = build_template_signature("任意Sheet名", profiles)

    assert left.signature == right.signature
    assert left.sheet_name == "订单"
    assert right.sheet_name == "任意Sheet名"


def test_template_similarity_scores_identical_structure_high():
    profiles = [
        ColumnProfile(column_name="购买方", column_index=1, data_type=DataType.TEXT),
        ColumnProfile(column_name="金额", column_index=2, data_type=DataType.NUMBER),
    ]
    signature = build_template_signature("订单", profiles)

    assert template_similarity(signature, signature) == 1.0
