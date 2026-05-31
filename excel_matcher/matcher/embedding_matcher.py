from excel_matcher.models import ColumnProfile, StageResult, StageStatus, StandardField


def profile_to_semantic_text(profile: ColumnProfile) -> str:
    sample_text = " | ".join(str(sample) for sample in profile.samples[:5])
    return f"field={profile.column_name}; type={profile.data_type.value}; samples={sample_text}"


class EmbeddingMatcher:
    def __init__(self, fields: list[StandardField], vector_store, top_k: int = 5):
        self.fields = fields
        self.vector_store = vector_store
        self.top_k = top_k

    def match(self, profile: ColumnProfile) -> StageResult:
        if self.vector_store is None or not self.vector_store.is_available():
            return StageResult(
                stage="embedding",
                status=StageStatus.SKIPPED,
                reason="vector store is not available",
            )
        candidates = self.vector_store.query(profile_to_semantic_text(profile), self.top_k)
        return StageResult(
            stage="embedding",
            status=StageStatus.COMPLETED,
            candidates=candidates,
        )
