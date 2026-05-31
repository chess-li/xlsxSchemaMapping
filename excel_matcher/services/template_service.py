import hashlib
import json

from excel_matcher.matcher.normalization import normalize_field_name
from excel_matcher.models import ColumnProfile, TemplateSignature


def build_template_signature(
    sheet_name: str,
    profiles: list[ColumnProfile],
) -> TemplateSignature:
    normalized_fields = [normalize_field_name(profile.column_name) for profile in profiles]
    data_types = [profile.data_type for profile in profiles]
    payload = {
        "count": len(profiles),
        "fields": normalized_fields,
        "data_types": [data_type.value for data_type in data_types],
    }
    signature = hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()
    return TemplateSignature(
        sheet_name=sheet_name,
        signature=signature,
        normalized_fields=normalized_fields,
        data_types=data_types,
    )


def template_similarity(left: TemplateSignature, right: TemplateSignature) -> float:
    if left.signature == right.signature:
        return 1.0
    fields_jaccard = _jaccard(left.normalized_fields, right.normalized_fields)
    ordered_overlap = _ordered_overlap(left.normalized_fields, right.normalized_fields)
    type_overlap = _ordered_overlap(
        [data_type.value for data_type in left.data_types],
        [data_type.value for data_type in right.data_types],
    )
    return round(fields_jaccard * 0.45 + ordered_overlap * 0.45 + type_overlap * 0.10, 6)


def _jaccard(left: list[str], right: list[str]) -> float:
    left_set = set(left)
    right_set = set(right)
    if not left_set and not right_set:
        return 1.0
    return len(left_set & right_set) / len(left_set | right_set)


def _ordered_overlap(left: list[str], right: list[str]) -> float:
    if not left and not right:
        return 1.0
    width = max(len(left), len(right))
    matches = sum(1 for index in range(min(len(left), len(right))) if left[index] == right[index])
    return matches / width
