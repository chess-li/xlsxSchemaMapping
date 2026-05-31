from pathlib import Path

from excel_matcher.matcher.embedding_matcher import EmbeddingMatcher, profile_to_semantic_text
from excel_matcher.matcher.llm_matcher import CodexExecLLMMatcher
from excel_matcher.models import ColumnProfile, DataType, FieldMatchCandidate, StageStatus, StandardField
from excel_matcher.vector.faiss_store import FaissVectorStore


def test_profile_to_semantic_text_includes_field_name_type_and_samples():
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯", "阿里", "字节", "百度", "京东", "小米"],
    )

    text = profile_to_semantic_text(profile)

    assert "购买方" in text
    assert "TEXT" in text
    assert "腾讯" in text
    assert "京东" in text
    assert "小米" not in text


def test_embedding_matcher_skips_when_vector_store_unavailable():
    fields = [
        StandardField(
            key="customer_name",
            display_name="客户名称",
            domain="order",
            aliases=["购买方"],
        )
    ]
    matcher = EmbeddingMatcher(fields=fields, vector_store=None)
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯"],
    )

    result = matcher.match(profile)

    assert result.status == StageStatus.SKIPPED
    assert "vector store" in result.reason.lower()
    assert result.candidates == []


def test_embedding_matcher_uses_profile_semantic_text_for_query():
    class FakeVectorStore:
        def __init__(self):
            self.query_text = ""
            self.query_top_k = 0

        def is_available(self):
            return True

        def query(self, text: str, top_k: int):
            self.query_text = text
            self.query_top_k = top_k
            return [
                FieldMatchCandidate(
                    target_field="customer_name",
                    score=0.91,
                    source="embedding",
                    reason="semantic",
                )
            ]

    store = FakeVectorStore()
    fields = [
        StandardField(
            key="customer_name",
            display_name="客户名称",
            domain="order",
            aliases=["购买方"],
        )
    ]
    matcher = EmbeddingMatcher(fields=fields, vector_store=store, top_k=3)
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯"],
    )

    result = matcher.match(profile)

    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert "购买方" in store.query_text
    assert "腾讯" in store.query_text
    assert store.query_top_k == 3


def test_faiss_vector_store_initializes_without_importing_optional_dependencies(tmp_path):
    store = FaissVectorStore(
        model_name="BAAI/bge-small-zh-v1.5",
        index_path=tmp_path / "index.faiss",
        metadata_path=tmp_path / "metadata.json",
    )

    assert store.model_name == "BAAI/bge-small-zh-v1.5"
    assert store.index_path == Path(tmp_path / "index.faiss")
    assert store.metadata_path == Path(tmp_path / "metadata.json")
    assert store.is_available() is False


def test_llm_matcher_skips_without_candidates():
    matcher = CodexExecLLMMatcher(enabled=True)
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯"],
    )

    result = matcher.match(profile, candidates=[])

    assert result.status == StageStatus.SKIPPED
    assert result.reason == "llm semantic scoring requires candidates"


def test_llm_matcher_prompt_includes_data_samples_and_parses_json():
    captured = {}

    def runner(prompt: str, timeout: int) -> str:
        captured["prompt"] = prompt
        captured["timeout"] = timeout
        return '{"target_field":"customer_name","semantic_score":0.96,"reason":"购买方和样例腾讯表示客户"}'

    matcher = CodexExecLLMMatcher(enabled=True, runner=runner, timeout_seconds=5)
    profile = ColumnProfile(
        column_name="购买方",
        column_index=1,
        data_type=DataType.TEXT,
        samples=["腾讯"],
    )
    candidates = [
        FieldMatchCandidate(
            target_field="customer_name",
            score=0.88,
            source="embedding",
        )
    ]

    result = matcher.match(profile, candidates=candidates)

    assert "购买方" in captured["prompt"]
    assert "TEXT" in captured["prompt"]
    assert "腾讯" in captured["prompt"]
    assert "customer_name" in captured["prompt"]
    assert captured["timeout"] == 5
    assert result.status == StageStatus.COMPLETED
    assert result.candidates[0].target_field == "customer_name"
    assert result.candidates[0].score == 0.96
    assert result.candidates[0].source == "llm"
