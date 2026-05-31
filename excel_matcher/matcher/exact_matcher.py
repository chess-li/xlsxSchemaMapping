from excel_matcher.matcher.normalization import normalize_field_name
from excel_matcher.models import FieldMatchCandidate, StageResult, StageStatus, StandardField


class ExactMatcher:
    def __init__(self, fields: list[StandardField]):
        self._index: dict[str, StandardField] = {}
        for field in fields:
            for value in [field.key, field.display_name, *field.aliases]:
                normalized = normalize_field_name(value)
                if normalized:
                    self._index.setdefault(normalized, field)

    def match(self, field_name: str) -> StageResult:
        normalized = normalize_field_name(field_name)
        field = self._index.get(normalized)
        candidates = []
        if field is not None:
            candidates.append(
                FieldMatchCandidate(
                    target_field=field.key,
                    score=1.0,
                    source="exact",
                    reason="normalized field name matched exactly",
                )
            )
        return StageResult(
            stage="exact",
            status=StageStatus.COMPLETED,
            candidates=candidates,
        )
