# Excel Matcher MVP

Local Excel intelligent business matching MVP for parsing workbooks, detecting headers, profiling columns, matching business fields, and saving confirmed templates.

## Setup

```bash
python3.11 -m venv .venv
.venv/bin/python -m pip install -e ".[test]"
```

Phase 2 semantic dependencies:

```bash
.venv/bin/python -m pip install -e ".[test,semantic]"
```

## Test

```bash
.venv/bin/python -m pytest
```

## Run API

```bash
.venv/bin/uvicorn excel_matcher.api.app:app --reload
```

## Run UI

```bash
.venv/bin/streamlit run excel_matcher/ui/streamlit_app.py
```

## Phase 2 Semantic Setup

The semantic stack uses `BAAI/bge-small-zh-v1.5` through sentence-transformers and stores FAISS files under `data/faiss/`.
