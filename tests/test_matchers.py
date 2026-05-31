from excel_matcher.matcher.dictionary_matcher import DictionaryMatcher
from excel_matcher.matcher.exact_matcher import ExactMatcher
from excel_matcher.matcher.fuzzy_matcher import FuzzyMatcher
from excel_matcher.matcher.normalization import normalize_field_name
from excel_matcher.models import StageStatus, StandardField


FIELDS = [
    StandardField(
        key="order_no",
        display_name="订单号",
        domain="order",
        aliases=["订单编号", "单号", "销售单号"],
    ),
    StandardField(
        key="customer_name",
        display_name="客户名称",
        domain="order",
        aliases=["购买方", "企业名称", "客户名", "公司名称"],
    ),
    StandardField(
        key="amount",
        display_name="金额",
        domain="order",
        aliases=["订单金额", "含税金额", "应收金额", "总金额"],
    ),
]


def test_normalize_field_name_removes_suffix_markers_and_common_punctuation():
    assert normalize_field_name(" SKU-Code（必填）* ") == "skucode"
    assert normalize_field_name(" 客户 名称 (可选) ") == "客户名称"


def test_exact_matcher_matches_display_names_aliases_and_keys():
    matcher = ExactMatcher(FIELDS)

    display_result = matcher.match("客户名称")
    alias_result = matcher.match(" 购买方（必填） ")
    key_result = matcher.match("CUSTOMER_NAME")

    assert display_result.status == StageStatus.COMPLETED
    assert display_result.candidates[0].target_field == "customer_name"
    assert display_result.candidates[0].score == 1.0
    assert alias_result.candidates[0].target_field == "customer_name"
    assert key_result.candidates[0].target_field == "customer_name"


def test_dictionary_matcher_returns_alias_candidate_with_dictionary_source():
    result = DictionaryMatcher(FIELDS).match("企业名称")

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert result.candidates[0].score == 1.0
    assert result.candidates[0].source == "dictionary"


def test_fuzzy_matcher_sorts_top_candidates_by_normalized_score():
    result = FuzzyMatcher(FIELDS).match("客户名称必填")

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert 0.0 <= result.candidates[0].score <= 1.0
    assert len(result.candidates) <= 5
    assert result.candidates == sorted(
        result.candidates,
        key=lambda candidate: candidate.score,
        reverse=True,
    )
