from rapidfuzz import fuzz

from excel_matcher.matcher.normalization import normalize_field_name
from excel_matcher.models import FieldMatchCandidate, StageResult, StageStatus, StandardField


class FuzzyMatcher:
    def __init__(self, fields: list[StandardField], limit: int = 5):
        self.fields = fields
        self.limit = limit

    def match(self, field_name: str) -> StageResult:
        normalized_query = normalize_field_name(field_name)
        candidates = []
        for field in self.fields:
            score, matched_text = self._best_score(normalized_query, field)
            if score > 0:
                candidates.append(
                    FieldMatchCandidate(
                        target_field=field.key,
                        score=score,
                        source="rapidfuzz",
                        reason=f"best fuzzy match: {matched_text}",
                    )
                )
        candidates.sort(key=lambda candidate: candidate.score, reverse=True)
        return StageResult(
            stage="rapidfuzz",
            status=StageStatus.COMPLETED,
            candidates=candidates[: self.limit],
        )

    def _best_score(self, normalized_query: str, field: StandardField) -> tuple[float, str]:
        best_score = 0.0
        best_text = ""
        for text in [field.display_name, *field.aliases]:
            normalized_candidate = normalize_field_name(text)
            score = fuzz.WRatio(normalized_query, normalized_candidate) / 100
            if score > best_score:
                best_score = score
                best_text = text
        return best_score, best_text
