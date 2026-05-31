from pathlib import Path

from pydantic import BaseModel, Field


class Settings(BaseModel):
    data_dir: Path = Field(default=Path("data"))
    sqlite_path: Path = Field(default=Path("data/excel_matcher.db"))
    faiss_index_path: Path = Field(default=Path("data/faiss/index.faiss"))
    faiss_metadata_path: Path = Field(default=Path("data/faiss/metadata.json"))
    embedding_model_name: str = "BAAI/bge-small-zh-v1.5"
    embedding_top_k: int = 5
    enable_embedding: bool = False
    enable_llm: bool = False
    llm_timeout_seconds: int = 60
    review_threshold: float = 0.75
    close_candidate_delta: float = 0.08
