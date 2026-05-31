from excel_matcher.services.matching_service import match_profiles
from excel_matcher.services.workbook_service import analyze_workbook
from excel_matcher.storage.default_fields import get_default_fields
from tests.sample_generator import generate_validation_samples


REQUIRED_ALIASES = {"购买方", "企业名称", "订单金额", "SKU", "库存数量", "联系人", "发票号"}


def test_generated_samples_cover_required_count(tmp_path):
    samples = generate_validation_samples(tmp_path)

    assert len(samples) >= 20
    assert {sample.domain for sample in samples} >= {"order", "crm", "inventory", "finance"}
    assert all(sample.path.exists() for sample in samples)
    generated_fields = {
        excel_field
        for sample in samples
        for excel_field in sample.expected_mappings
    }
    assert REQUIRED_ALIASES <= generated_fields


def test_header_detection_accuracy_on_generated_samples(tmp_path):
    samples = generate_validation_samples(tmp_path)
    correct = 0

    for sample in samples:
        analysis = analyze_workbook(sample.path)
        sheet_name = analysis["workbook"].sheets[0].name
        header = analysis["headers"][sheet_name]
        if (
            header.header_row == sample.expected_header_row
            and header.data_start_row == sample.expected_data_start_row
        ):
            correct += 1

    assert correct / len(samples) >= 0.80


def test_base_field_matching_accuracy_on_generated_samples(tmp_path):
    samples = generate_validation_samples(tmp_path)
    standard_fields = get_default_fields()
    total = 0
    correct = 0

    for sample in samples:
        analysis = analyze_workbook(sample.path)
        sheet_name = analysis["workbook"].sheets[0].name
        mappings = match_profiles(
            analysis["profiles"][sheet_name],
            standard_fields,
            enable_embedding=False,
            enable_llm=False,
        )
        by_excel_field = {mapping.excel_field: mapping.target_field for mapping in mappings}
        for excel_field, expected_target in sample.expected_mappings.items():
            total += 1
            if by_excel_field.get(excel_field) == expected_target:
                correct += 1

    assert correct / total >= 0.85
