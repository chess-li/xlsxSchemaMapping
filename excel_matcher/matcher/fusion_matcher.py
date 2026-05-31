from dataclasses import dataclass

from excel_matcher.config import Settings
from excel_matcher.matcher.dictionary_matcher import DictionaryMatcher
from excel_matcher.matcher.exact_matcher import ExactMatcher
from excel_matcher.matcher.fuzzy_matcher import FuzzyMatcher
from excel_matcher.models import (
    ColumnProfile,
    FieldMappingResult,
    FieldMatchCandidate,
    StageResult,
    StageStatus,
    StandardField,
)


PHASE1_PRIORITIES = {
    "exact": 3,
    "dictionary": 2,
    "rapidfuzz": 1,
}


@dataclass(frozen=True)
class RankedCandidate:
    candidate: FieldMatchCandidate
    confidence: float
    priority: int


class FusionMatcher:
    def __init__(
        self,
        standard_fields: list[StandardField],
        enable_embedding: bool = False,
        enable_llm: bool = False,
        settings: Settings | None = None,
    ):
        self.standard_fields = standard_fields
        self.enable_embedding = enable_embedding
        self.enable_llm = enable_llm
        self.settings = settings or Settings()
        self.exact_matcher = ExactMatcher(standard_fields)
        self.dictionary_matcher = DictionaryMatcher(standard_fields)
        self.fuzzy_matcher = FuzzyMatcher(standard_fields)

    def match(self, profile: ColumnProfile) -> FieldMappingResult:
        stage_results = [
            self.exact_matcher.match(profile.column_name),
            self.dictionary_matcher.match(profile.column_name),
            self.fuzzy_matcher.match(profile.column_name),
            self._embedding_stage_result(),
            self._llm_stage_result(),
        ]
        ranked_candidates = self._rank_candidates(stage_results)
        best = ranked_candidates[0] if ranked_candidates else None
        candidates = [ranked.candidate for ranked in ranked_candidates]
        confidence = best.confidence if best is not None else 0.0
        target_field = best.candidate.target_field if best is not None else None
        return FieldMappingResult(
            excel_field=profile.column_name,
            target_field=target_field,
            confidence=confidence,
            needs_review=self._needs_review(ranked_candidates),
            candidates=candidates,
            stage_results=stage_results,
        )

    def _embedding_stage_result(self) -> StageResult:
        if not self.enable_embedding:
            return StageResult(
                stage="embedding",
                status=StageStatus.SKIPPED,
                reason="embedding semantic scoring is disabled in phase 1",
            )
        return StageResult(
            stage="embedding",
            status=StageStatus.SKIPPED,
            reason="embedding matcher is not configured",
        )

    def _llm_stage_result(self) -> StageResult:
        if not self.enable_llm:
            return StageResult(
                stage="llm",
                status=StageStatus.SKIPPED,
                reason="llm semantic scoring is disabled in phase 1",
            )
        return StageResult(
            stage="llm",
            status=StageStatus.SKIPPED,
            reason="llm matcher is not configured",
        )

    def _rank_candidates(self, stage_results: list[StageResult]) -> list[RankedCandidate]:
        best_by_target: dict[str, RankedCandidate] = {}
        for stage_result in stage_results:
            priority = PHASE1_PRIORITIES.get(stage_result.stage, 0)
            if priority == 0 or stage_result.status != StageStatus.COMPLETED:
                continue
            for candidate in stage_result.candidates:
                confidence = 0.98 if stage_result.stage in {"exact", "dictionary"} else candidate.score
                ranked = RankedCandidate(candidate=candidate, confidence=confidence, priority=priority)
                existing = best_by_target.get(candidate.target_field)
                if existing is None or (ranked.priority, ranked.confidence) > (
                    existing.priority,
                    existing.confidence,
                ):
                    best_by_target[candidate.target_field] = ranked
        return sorted(
            best_by_target.values(),
            key=lambda ranked: (ranked.priority, ranked.confidence),
            reverse=True,
        )

    def _needs_review(self, ranked_candidates: list[RankedCandidate]) -> bool:
        if not ranked_candidates:
            return True
        if ranked_candidates[0].confidence < self.settings.review_threshold:
            return True
        if len(ranked_candidates) < 2:
            return False
        return (
            ranked_candidates[0].confidence - ranked_candidates[1].confidence
            < self.settings.close_candidate_delta
        )
