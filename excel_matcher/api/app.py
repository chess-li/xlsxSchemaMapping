import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, UploadFile
from fastapi.encoders import jsonable_encoder
from pydantic import BaseModel, Field

from excel_matcher.config import Settings
from excel_matcher.models import ColumnProfile
from excel_matcher.services.matching_service import match_profiles
from excel_matcher.services.workbook_service import analyze_workbook
from excel_matcher.storage.default_fields import get_default_fields
from excel_matcher.storage.sqlite_store import SQLiteStore


app = FastAPI(title="Excel Matcher MVP")


class MatchRequest(BaseModel):
    profiles: list[ColumnProfile]
    enable_embedding: bool = False
    enable_llm: bool = False


class TemplateSaveRequest(BaseModel):
    signature: str
    sheet_name: str
    mappings: dict[str, str] = Field(default_factory=dict)


@app.get("/health")
def health():
    sqlite_path = _sqlite_path()
    return {
        "components": {
            "sqlite": {
                "available": True,
                "reason": f"sqlite path configured at {sqlite_path}",
            },
            "embedding": {
                "available": False,
                "reason": "embedding semantic scoring is disabled by default",
            },
            "llm": {
                "available": False,
                "reason": "llm semantic scoring is disabled by default",
            },
        }
    }


@app.get("/fields")
def fields():
    return {"fields": jsonable_encoder(get_default_fields())}


@app.post("/workbooks/analyze")
async def analyze_uploaded_workbook(file: UploadFile):
    suffix = Path(file.filename or "workbook.xlsx").suffix or ".xlsx"
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
            temporary_path = Path(temporary_file.name)
            temporary_file.write(await file.read())
        return jsonable_encoder(analyze_workbook(temporary_path))
    finally:
        if temporary_path is not None and temporary_path.exists():
            os.unlink(temporary_path)


@app.post("/fields/match")
def match_fields(request: MatchRequest):
    mappings = match_profiles(
        request.profiles,
        get_default_fields(),
        enable_embedding=request.enable_embedding,
        enable_llm=request.enable_llm,
    )
    return {"mappings": jsonable_encoder(mappings)}


@app.post("/templates")
def save_template(request: TemplateSaveRequest):
    store = SQLiteStore(_sqlite_path())
    store.initialize()
    store.save_template(
        signature=request.signature,
        sheet_name=request.sheet_name,
        mappings=request.mappings,
    )
    return {"saved": True}


def _sqlite_path() -> Path:
    return Path(getattr(app.state, "sqlite_path", Settings().sqlite_path))
