import json
from pathlib import Path

import numpy as np

from excel_matcher.models import FieldMatchCandidate, StandardField


class FaissVectorStore:
    def __init__(self, model_name: str, index_path, metadata_path):
        self.model_name = model_name
        self.index_path = Path(index_path)
        self.metadata_path = Path(metadata_path)
        self.model = None
        self.index = None
        self.metadata: list[dict[str, str]] = []

    def is_available(self) -> bool:
        return self.index is not None and bool(self.metadata)

    def build(self, fields: list[StandardField]) -> None:
        import faiss

        model = self._load_model()
        texts, metadata = self._field_texts(fields)
        embeddings = np.asarray(model.encode(texts), dtype="float32")
        faiss.normalize_L2(embeddings)
        index = faiss.IndexFlatIP(embeddings.shape[1])
        index.add(embeddings)
        self.index = index
        self.metadata = metadata

    def save(self) -> None:
        if self.index is None:
            return
        import faiss

        self.index_path.parent.mkdir(parents=True, exist_ok=True)
        self.metadata_path.parent.mkdir(parents=True, exist_ok=True)
        faiss.write_index(self.index, str(self.index_path))
        self.metadata_path.write_text(
            json.dumps(self.metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def load(self) -> None:
        if not self.index_path.exists() or not self.metadata_path.exists():
            return
        import faiss

        self.index = faiss.read_index(str(self.index_path))
        self.metadata = json.loads(self.metadata_path.read_text(encoding="utf-8"))

    def query(self, text: str, top_k: int) -> list[FieldMatchCandidate]:
        if not self.is_available():
            return []
        import faiss

        model = self._load_model()
        query_embedding = np.asarray(model.encode([text]), dtype="float32")
        faiss.normalize_L2(query_embedding)
        scores, indices = self.index.search(query_embedding, top_k)
        candidates = []
        for score, index in zip(scores[0], indices[0]):
            if index < 0:
                continue
            metadata = self.metadata[index]
            candidates.append(
                FieldMatchCandidate(
                    target_field=metadata["field_key"],
                    score=float(score),
                    source="embedding",
                    reason=f"{metadata['source_type']}: {metadata['source_text']}",
                )
            )
        return candidates

    def _load_model(self):
        if self.model is None:
            from sentence_transformers import SentenceTransformer

            self.model = SentenceTransformer(self.model_name)
        return self.model

    def _field_texts(self, fields: list[StandardField]) -> tuple[list[str], list[dict[str, str]]]:
        texts = []
        metadata = []
        for field in fields:
            entries = [
                ("key", field.key),
                ("display_name", field.display_name),
                ("description", field.description),
                *[("alias", alias) for alias in field.aliases],
            ]
            for source_type, source_text in entries:
                if not source_text:
                    continue
                texts.append(source_text)
                metadata.append(
                    {
                        "field_key": field.key,
                        "domain": field.domain,
                        "source_text": source_text,
                        "source_type": source_type,
                    }
                )
        return texts, metadata
