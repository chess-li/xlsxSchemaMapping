# Excel Matcher MVP Goal Run Plan

> **For Codex Goal workers:** This plan is optimized for long-running, resumable goal execution. Start or continue a goal with the objective below, then work through the first unchecked work package. Keep this file as the progress ledger. Do not mark the goal complete until the final verification gate passes.

**Goal Objective:** Implement and verify the local Excel intelligent business matching MVP from `milestone_1.md`, delivering a deterministic Phase 1 and an optional semantic Phase 2.

**Architecture:** Use a thin-entry layered Python package. Core Pydantic models and services hold business behavior; Streamlit and FastAPI call the same services. Phase 1 runs `Exact -> Dictionary -> RapidFuzz` and reports `Embedding/FAISS` plus `LLM` as skipped. Phase 2 upgrades the same path with optional Embedding/FAISS and `codex exec` semantic scoring using both field names and sample values.

**Tech Stack:** Python 3.11+, Pydantic v2, openpyxl, pandas, RapidFuzz, SQLite, FastAPI, Streamlit, pytest. Phase 2 adds optional `sentence-transformers`, FAISS, and `codex exec`.

**Reference Spec:** `docs/superpowers/specs/2026-05-31-excel-matcher-mvp-design.md`

---

## Goal Run Protocol

1. Before each execution turn, run `git status --short` and identify user changes. Do not overwrite unrelated changes.
2. Read this file and resume at the first unchecked work package or failed gate.
3. For each work package, follow the listed RED/PASS loop: write tests first, confirm failure, implement, confirm pass, then commit.
4. Keep commits small and scoped to the package. If a package is purely verification, do not create an empty commit.
5. If a command fails because dependencies are missing, install only the needed extras for the current phase. Phase 1 must use `pip install -e ".[test]"`; Phase 2 may use `pip install -e ".[test,semantic]"`.
6. If a test exposes a design mismatch, fix the package rather than rewriting the acceptance criteria unless the spec clearly disagrees.
7. If blocked for the same external reason across three goal turns, record the blocker under "Blocked Notes" and stop with the goal marked blocked.
8. After every completed package, update the checklist in this plan before moving on.

## Progress Ledger

- [x] G0: Rehydrate workspace and confirm baseline
- [x] G1: Project scaffold, settings, and shared models
- [x] G2: Excel parsing, header detection, and profiling
- [x] G3: SQLite storage and deterministic text matchers
- [x] G4: Phase 1 fusion, workbook services, and template learning
- [x] G5: Generated validation samples and accuracy gates
- [x] G6: FastAPI and Streamlit entry points
- [x] G7: Phase 1 verification gate
- [x] G8: Phase 2 Embedding/FAISS semantic stage
- [x] G9: Phase 2 LLM semantic stage and fusion upgrade
- [x] G10: Phase 2 verification gate

## Phase Rules

**Phase 1 deterministic rules:**

- Do not import, load, or call `sentence-transformers`, FAISS, `codex exec`, or any LLM provider.
- Field matching uses field names only: exact, dictionary aliases, and RapidFuzz.
- `ColumnProfile.samples` may be collected and displayed, but matching must not use sample values for semantic meaning.
- UI and API must show the full stage order and skipped reasons for disabled semantic stages.

**Phase 2 semantic rules:**

- Semantic dependencies are optional and must not be imported at module import time.
- Embedding/FAISS and LLM stages may use `ColumnProfile.samples`.
- Tests for semantic stages must use fakes or runner injection and must not download models or invoke real `codex exec`.
- All Phase 1 tests must continue to pass with `enable_embedding=False` and `enable_llm=False`.

## Target File Map

Create or modify these files only as needed for the listed packages:

- `pyproject.toml`: project metadata, dependencies, optional extras, pytest config.
- `.gitignore`: generated data, caches, virtualenv, local databases.
- `README.md`: setup, test, API, UI, and Phase 2 semantic instructions.
- `excel_matcher/__init__.py`: package marker and version.
- `excel_matcher/config.py`: runtime settings and default paths.
- `excel_matcher/models.py`: Pydantic contracts shared by services, API, and UI.
- `excel_matcher/excel/parser.py`: openpyxl workbook parser.
- `excel_matcher/excel/header_detector.py`: heuristic header detector.
- `excel_matcher/excel/profiler.py`: column profile builder.
- `excel_matcher/storage/default_fields.py`: built-in standard fields and aliases.
- `excel_matcher/storage/schema.py`: SQLite DDL.
- `excel_matcher/storage/sqlite_store.py`: repository API.
- `excel_matcher/matcher/normalization.py`: field text normalization.
- `excel_matcher/matcher/exact_matcher.py`: exact matcher stage.
- `excel_matcher/matcher/dictionary_matcher.py`: dictionary alias matcher stage.
- `excel_matcher/matcher/fuzzy_matcher.py`: RapidFuzz matcher stage.
- `excel_matcher/matcher/fusion_matcher.py`: staged matcher and confidence fusion.
- `excel_matcher/matcher/embedding_matcher.py`: Phase 2 embedding stage.
- `excel_matcher/matcher/llm_matcher.py`: Phase 2 `codex exec` LLM stage.
- `excel_matcher/vector/faiss_store.py`: optional FAISS index wrapper.
- `excel_matcher/services/workbook_service.py`: parse, detect headers, profile columns.
- `excel_matcher/services/matching_service.py`: matching orchestration.
- `excel_matcher/services/template_service.py`: signature, save, and similarity helpers.
- `excel_matcher/api/app.py`: FastAPI app.
- `excel_matcher/ui/streamlit_app.py`: Streamlit review UI.
- `tests/sample_generator.py`: deterministic `.xlsx` validation fixture generator.
- `tests/test_models.py`, `tests/test_parser.py`, `tests/test_header_detector.py`, `tests/test_profiler.py`, `tests/test_storage.py`, `tests/test_matchers.py`, `tests/test_services.py`, `tests/test_sample_accuracy.py`, `tests/test_api.py`, `tests/test_semantic_stages.py`.

---

## G0: Rehydrate Workspace And Confirm Baseline

**Objective:** Establish current workspace state before implementation.

**Commands:**

```bash
git status --short
find . -maxdepth 2 -type f | sort | sed 's#^\./##' | head -200
```

**Acceptance:**

- Current branch and dirty files are understood.
- Existing user changes are not reverted.
- If implementation files already exist, inspect them and continue from the current state rather than recreating blindly.

**Commit:** None.

---

## G1: Project Scaffold, Settings, And Shared Models

**Objective:** Create the package, dependency metadata, settings, and shared contracts.

**Files:**

- Create: `pyproject.toml`, `.gitignore`, `README.md`, `excel_matcher/__init__.py`
- Create: `excel_matcher/config.py`, `excel_matcher/models.py`
- Create: `tests/test_models.py`

**Tests To Write First:**

- `test_workbook_model_serializes_nested_cells`
- `test_stage_result_preserves_skip_reason`
- `test_mapping_result_marks_review_when_confidence_is_low`
- `test_column_profile_uses_supported_data_type_enum`
- `test_header_detection_result_has_required_rows`

**Model Contract:**

- `DataType`: `TEXT`, `NUMBER`, `DATE`, `BOOLEAN`.
- `StageStatus`: `completed`, `skipped`, `failed`.
- `CellData`: `row`, `column`, `coordinate`, `raw_value`, `display_value`, optional `merged_anchor`.
- `SheetData`: `name`, `max_row`, `max_column`, `hidden_columns`, `merged_ranges`, nested `cells`.
- `WorkbookData`: `path`, `sheets`.
- `HeaderDetectionResult`: `sheet_name`, `header_row`, `data_start_row`, optional `title_row`, `confidence`, `evidence`.
- `ColumnProfile`: `column_name`, `column_index`, `data_type`, `null_rate`, `unique_rate`, `samples`, `evidence`.
- `StandardField`: `key`, `display_name`, `domain`, `description`, `aliases`.
- `FieldMatchCandidate`: `target_field`, `score`, `source`, `reason`.
- `StageResult`: `stage`, `status`, `reason`, `candidates`, `diagnostics`.
- `FieldMappingResult`: `excel_field`, optional `target_field`, `confidence`, `needs_review`, `candidates`, `stage_results`.
- `TemplateSignature`: `sheet_name`, `signature`, `normalized_fields`, `data_types`.

**Settings Contract:**

- `data_dir=Path("data")`
- `sqlite_path=Path("data/excel_matcher.db")`
- `faiss_index_path=Path("data/faiss/index.faiss")`
- `faiss_metadata_path=Path("data/faiss/metadata.json")`
- `embedding_model_name="BAAI/bge-small-zh-v1.5"`
- `embedding_top_k=5`
- `enable_embedding=False`
- `enable_llm=False`
- `llm_timeout_seconds=60`
- `review_threshold=0.75`
- `close_candidate_delta=0.08`

**Commands:**

```bash
pytest tests/test_models.py -q
```

Expected RED before implementation: missing module or missing model names.

Expected PASS after implementation: all model tests pass.

**Commit:**

```bash
git add pyproject.toml .gitignore README.md excel_matcher/__init__.py excel_matcher/config.py excel_matcher/models.py tests/test_models.py
git commit -m "feat: add shared models and settings"
```

---

## G2: Excel Parsing, Header Detection, And Profiling

**Objective:** Turn `.xlsx` workbooks into structured sheet data, detect header rows, and profile columns.

**Files:**

- Create: `excel_matcher/excel/__init__.py`
- Create: `excel_matcher/excel/parser.py`
- Create: `excel_matcher/excel/header_detector.py`
- Create: `excel_matcher/excel/profiler.py`
- Create: `tests/test_parser.py`, `tests/test_header_detector.py`, `tests/test_profiler.py`

**Parser Acceptance:**

- Uses `openpyxl.load_workbook(..., data_only=True)`.
- Preserves sheet order, cell coordinates, raw values, display values, merged ranges, merged anchors, and hidden columns.
- For non-anchor merged cells, `display_value` resolves from the anchor cell.

**Header Detector Acceptance:**

- Detects header after a title row.
- Skips blank rows.
- Returns `title_row`, `header_row`, `data_start_row`, `confidence`, and evidence.
- Heuristic considers non-empty ratio, text ratio, business-term hits, and next-row data shape.

**Profiler Acceptance:**

- Builds one `ColumnProfile` per header cell.
- Infers `TEXT`, `NUMBER`, `DATE`, and `BOOLEAN`.
- Computes `null_rate` and `unique_rate`.
- Keeps up to five non-null sample values in original order.
- Falls back to `column_N` for blank headers.

**Commands:**

```bash
pytest tests/test_parser.py tests/test_header_detector.py tests/test_profiler.py -q
```

Expected RED before implementation: missing modules/functions.

Expected PASS after implementation: parser, detector, and profiler tests pass.

**Commit:**

```bash
git add excel_matcher/excel tests/test_parser.py tests/test_header_detector.py tests/test_profiler.py
git commit -m "feat: analyze excel workbook structure"
```

---

## G3: SQLite Storage And Deterministic Text Matchers

**Objective:** Add built-in standard fields, template persistence, normalization, and deterministic matching stages.

**Files:**

- Create: `excel_matcher/storage/__init__.py`
- Create: `excel_matcher/storage/default_fields.py`
- Create: `excel_matcher/storage/schema.py`
- Create: `excel_matcher/storage/sqlite_store.py`
- Create: `excel_matcher/matcher/__init__.py`
- Create: `excel_matcher/matcher/normalization.py`
- Create: `excel_matcher/matcher/exact_matcher.py`
- Create: `excel_matcher/matcher/dictionary_matcher.py`
- Create: `excel_matcher/matcher/fuzzy_matcher.py`
- Create: `tests/test_storage.py`, `tests/test_matchers.py`

**Default Field Minimum Set:**

- `order_no`: display `订单号`, aliases `订单编号`, `单号`, `销售单号`.
- `customer_name`: display `客户名称`, aliases `购买方`, `企业名称`, `客户名`, `公司名称`.
- `amount`: display `金额`, aliases `订单金额`, `含税金额`, `应收金额`, `总金额`.
- `order_date`: display `下单日期`, aliases `订单日期`, `日期`, `业务日期`.
- `product_name`: display `商品名称`, aliases `产品名称`, `物料名称`, `品名`.
- `quantity`: display `数量`, aliases `订购数量`, `采购数量`, `销售数量`.
- `sku`: display `SKU`, aliases `商品编码`, `物料编码`, `产品编码`.
- `inventory_qty`: display `库存数量`, aliases `库存`, `现存量`, `可用库存`.
- `contact_name`: display `联系人`, aliases `联系人姓名`, `客户联系人`.
- `phone`: display `电话`, aliases `手机号`, `联系电话`, `手机`.
- `invoice_no`: display `发票号`, aliases `发票号码`, `票据号`.
- `paid_status`: display `付款状态`, aliases `是否付款`, `是否支付`, `支付状态`.
- `account_name`: display `账户名称`, aliases `开户名`, `银行账户名`.
- `department`: display `部门`, aliases `所属部门`, `业务部门`.

**Storage Acceptance:**

- SQLite schema contains `standard_fields`, `field_aliases`, and `templates`.
- `SQLiteStore.initialize()` creates schema.
- `seed_default_fields()` inserts standard fields and aliases idempotently.
- `list_standard_fields()` returns `list[StandardField]` with aliases populated.
- `save_template(signature, sheet_name, mappings)` upserts JSON mappings.
- `find_template(signature)` returns a record with `signature`, `sheet_name`, and `mappings`, or `None`.

**Matcher Acceptance:**

- `normalize_field_name(text)` strips whitespace, lowercases ASCII, removes common punctuation including Chinese parentheses, and removes suffix markers such as `必填`, `可选`, and `*`.
- `ExactMatcher.match(field_name)` matches display names and aliases after normalization with score `1.0`.
- `DictionaryMatcher.match(field_name)` uses aliases and display names from standard fields.
- `FuzzyMatcher.match(field_name)` uses `rapidfuzz.fuzz.WRatio`, normalizes scores to `0.0-1.0`, sorts descending, and keeps up to five candidates over display names and aliases.

**Commands:**

```bash
pytest tests/test_storage.py tests/test_matchers.py -q
```

Expected RED before implementation: missing storage and matcher modules.

Expected PASS after implementation: storage and matcher tests pass.

**Commit:**

```bash
git add excel_matcher/storage excel_matcher/matcher tests/test_storage.py tests/test_matchers.py
git commit -m "feat: add sqlite dictionary and text matchers"
```

---

## G4: Phase 1 Fusion, Workbook Services, And Template Learning

**Objective:** Orchestrate deterministic matching and provide reusable service APIs.

**Files:**

- Create: `excel_matcher/matcher/fusion_matcher.py`
- Create: `excel_matcher/services/__init__.py`
- Create: `excel_matcher/services/matching_service.py`
- Create: `excel_matcher/services/workbook_service.py`
- Create: `excel_matcher/services/template_service.py`
- Create or modify: `tests/test_services.py`

**Tests To Write First:**

- `test_phase1_fusion_runs_deterministic_stages_and_marks_semantic_disabled`
- `test_phase1_matching_does_not_use_data_content_semantics`
- `test_match_profiles_returns_one_mapping_per_profile`
- `test_template_signature_is_stable_for_same_fields`
- `test_template_similarity_scores_identical_structure_high`

**Fusion Acceptance:**

- `FusionMatcher.__init__` accepts `standard_fields`, `enable_embedding=False`, `enable_llm=False`, and optional `settings`.
- `match(profile)` always returns five stage results in this order: `exact`, `dictionary`, `rapidfuzz`, `embedding`, `llm`.
- Phase 1 calls exact, dictionary, and RapidFuzz using only `profile.column_name`.
- When embedding is disabled, stage is skipped with reason `embedding semantic scoring is disabled in phase 1`.
- When LLM is disabled, stage is skipped with reason `llm semantic scoring is disabled in phase 1`.
- Candidate priority in Phase 1 is `exact=3`, `dictionary=2`, `rapidfuzz=1`.
- Exact and dictionary winners use confidence `0.98`; RapidFuzz uses its score.
- `needs_review=True` when confidence is below `Settings.review_threshold`.

**Service Acceptance:**

- `match_profiles(profiles, standard_fields, enable_embedding=False, enable_llm=False, settings=None)` returns one `FieldMappingResult` per profile.
- `analyze_workbook(path)` calls parser, header detector, and profiler for every sheet and returns structured workbook, headers, and profiles keyed by sheet name or equivalent Pydantic structure.
- `build_template_signature(sheet_name, profiles)` ignores sheet name for the hash, normalizes field names, and includes field order, count, and data type sequence.
- `template_similarity(left, right)` combines exact signature match, Jaccard field similarity, and ordered overlap. Identical profiles return `1.0`.

**Commands:**

```bash
pytest tests/test_services.py -q
```

Expected RED before implementation: missing fusion and service modules.

Expected PASS after implementation: all service tests pass.

**Commit:**

```bash
git add excel_matcher/matcher/fusion_matcher.py excel_matcher/services tests/test_services.py
git commit -m "feat: add phase 1 matching services"
```

---

## G5: Generated Validation Samples And Accuracy Gates

**Objective:** Generate deterministic workbook fixtures and enforce minimum Phase 1 accuracy.

**Files:**

- Create: `tests/sample_generator.py`
- Create: `tests/test_sample_accuracy.py`
- Modify if needed: `excel_matcher/storage/default_fields.py`, header detector, profiler, or matchers.

**Tests To Write First:**

- `test_generated_samples_cover_required_count`
- `test_header_detection_accuracy_on_generated_samples`
- `test_base_field_matching_accuracy_on_generated_samples`

**Sample Generator Acceptance:**

- Provides `ValidationSample` dataclass with `path`, `domain`, `expected_header_row`, `expected_data_start_row`, and `expected_mappings`.
- `generate_validation_samples(output_dir)` creates at least 20 `.xlsx` files with openpyxl.
- Samples cover `order`, `crm`, `inventory`, and `finance`.
- Include merged titles, blank rows, hidden columns, multiple sheets, shifted header rows, changed column order, aliases, and ambiguous fields.
- Include aliases such as `购买方`, `企业名称`, `订单金额`, `SKU`, `库存数量`, `联系人`, and `发票号`.

**Accuracy Acceptance:**

- Header detection accuracy on generated samples is at least `0.80`.
- Base field matching accuracy on generated samples is at least `0.85`.
- Accuracy tests use Phase 1 matching only: `enable_embedding=False`, `enable_llm=False`.

**Commands:**

```bash
pytest tests/test_sample_accuracy.py -q
```

Expected RED before implementation: missing sample generator.

Expected PASS after implementation: coverage and accuracy thresholds pass.

**Commit:**

```bash
git add tests/sample_generator.py tests/test_sample_accuracy.py excel_matcher/storage/default_fields.py excel_matcher/excel excel_matcher/matcher
git commit -m "test: add generated excel validation samples"
```

Only include touched implementation files in the commit.

---

## G6: FastAPI And Streamlit Entry Points

**Objective:** Expose the shared services through API and UI entry points.

**Files:**

- Create: `excel_matcher/api/__init__.py`
- Create: `excel_matcher/api/app.py`
- Create: `excel_matcher/ui/__init__.py`
- Create: `excel_matcher/ui/streamlit_app.py`
- Create: `tests/test_api.py`
- Modify: `README.md`

**API Tests To Write First:**

- `test_health_reports_components`
- `test_fields_endpoint_returns_built_in_dictionary`

**API Acceptance:**

- `FastAPI(title="Excel Matcher MVP")`.
- `GET /health` returns component statuses for `sqlite`, `embedding`, and `llm`, each with `available` and `reason`.
- `GET /fields` returns built-in fields.
- `POST /workbooks/analyze` accepts an uploaded `.xlsx`, saves it to a temporary path, and returns analysis.
- `POST /fields/match` accepts profiles and matching flags, returns mapping results.
- `POST /templates` saves confirmed mappings through `SQLiteStore.save_template()`.

**UI Acceptance:**

- Uploads `.xlsx`, saves to a temporary file, and calls `analyze_workbook`.
- Displays workbook summary, detected header rows, data start rows, and column profiles.
- Runs `match_profiles` and displays target field, confidence, review flag, and stage results.
- Shows each stage in order: Exact, Dictionary, RapidFuzz, Embedding/FAISS, LLM.
- Shows skipped reasons for unavailable stages.
- Allows manual target-field selection with `st.selectbox`.
- Saves confirmed mappings through `SQLiteStore.save_template()`.

**Commands:**

```bash
pytest tests/test_api.py -q
python -m compileall excel_matcher
```

Expected RED before API implementation: missing `excel_matcher.api.app`.

Expected PASS after implementation: API tests pass and compileall succeeds.

**Commit:**

```bash
git add excel_matcher/api excel_matcher/ui tests/test_api.py README.md
git commit -m "feat: expose api and streamlit interfaces"
```

---

## G7: Phase 1 Verification Gate

**Objective:** Prove deterministic MVP is complete and does not rely on semantic dependencies.

**Commands:**

```bash
pytest
python -m compileall excel_matcher tests
python -c "from excel_matcher.api.app import app; print(app.title)"
git status --short
```

**Acceptance:**

- Full test suite passes using `pip install -e ".[test]"`.
- Test run does not download embedding models.
- Test run does not invoke `codex exec`.
- Compile checks pass.
- API title output contains `Excel Matcher MVP`.
- `git status --short` shows no uncommitted implementation files.

**If Verification Fails:**

- Do not proceed to Phase 2.
- Create a focused fix inside the current phase with the failing test name, expected behavior, and exact file paths.
- After fixing, rerun the full Phase 1 gate.

**Commit:** None if verification only. If fixes were needed, commit the focused fix.

---

## G8: Phase 2 Embedding/FAISS Semantic Stage

**Objective:** Add optional embedding semantic matching while keeping imports lazy and tests offline.

**Files:**

- Create: `excel_matcher/vector/__init__.py`
- Create: `excel_matcher/vector/faiss_store.py`
- Create: `excel_matcher/matcher/embedding_matcher.py`
- Create or modify: `tests/test_semantic_stages.py`
- Modify: `README.md`

**Tests To Write First:**

- `test_profile_to_semantic_text_includes_field_name_type_and_samples`
- `test_embedding_matcher_skips_when_vector_store_unavailable`
- `test_embedding_matcher_uses_profile_semantic_text_for_query`

**Embedding Acceptance:**

- `profile_to_semantic_text(profile)` includes field name, data type, and up to five sample values.
- `EmbeddingMatcher(fields, vector_store, top_k=5).match(profile)` returns skipped when vector store is missing or unavailable.
- With an available fake vector store, matcher queries using semantic text and returns completed candidates.

**FAISS Store Acceptance:**

- `FaissVectorStore.__init__(model_name, index_path, metadata_path)` stores paths and initializes `model`, `index`, and metadata.
- Imports `faiss` and `SentenceTransformer` inside methods only.
- `is_available()` returns true only when index and metadata are loaded.
- `build(fields)` embeds each field key, display name, description, and alias; creates normalized inner-product FAISS index; writes metadata with `field_key`, `domain`, `source_text`, and `source_type`.
- `save()` writes index and metadata files under configured paths.
- `load()` reads index and metadata only when both files exist.
- `query(text, top_k)` embeds query text and returns `FieldMatchCandidate` entries with source `embedding`.

**README Acceptance:**

- Adds Phase 2 setup: `pip install -e ".[test,semantic]"`.
- Documents model `BAAI/bge-small-zh-v1.5` and FAISS files under `data/faiss/`.

**Commands:**

```bash
pytest tests/test_semantic_stages.py::test_profile_to_semantic_text_includes_field_name_type_and_samples tests/test_semantic_stages.py::test_embedding_matcher_skips_when_vector_store_unavailable tests/test_semantic_stages.py::test_embedding_matcher_uses_profile_semantic_text_for_query -q
```

Expected RED before implementation: missing `embedding_matcher`.

Expected PASS after implementation: tests pass without model downloads.

**Commit:**

```bash
git add excel_matcher/vector excel_matcher/matcher/embedding_matcher.py tests/test_semantic_stages.py README.md
git commit -m "feat: add phase 2 embedding semantic stage"
```

---

## G9: Phase 2 LLM Semantic Stage And Fusion Upgrade

**Objective:** Add injectable `codex exec` semantic scoring and integrate semantic stages into fusion.

**Files:**

- Create: `excel_matcher/matcher/llm_matcher.py`
- Modify: `excel_matcher/matcher/fusion_matcher.py`
- Modify: `excel_matcher/services/matching_service.py`
- Modify: `tests/test_semantic_stages.py`, `tests/test_services.py`

**LLM Tests To Write First:**

- `test_llm_matcher_skips_without_candidates`
- `test_llm_matcher_prompt_includes_data_samples_and_parses_json`

**Fusion Test To Write First:**

- `test_phase2_fusion_calls_embedding_then_llm_with_profile_samples`

**LLM Acceptance:**

- Class name: `CodexExecLLMMatcher`.
- Constructor: `__init__(self, enabled: bool = True, timeout_seconds: int = 60, runner=None)`.
- Public method: `match(self, profile: ColumnProfile, candidates: list[FieldMatchCandidate]) -> StageResult`.
- Disabled matcher returns skipped reason `llm semantic scoring is disabled`.
- Empty candidates return skipped reason `llm semantic scoring requires candidates`.
- Prompt includes profile field name, data type, up to five sample values, and candidate target fields.
- Default runner calls `subprocess.run(["codex", "exec", prompt], capture_output=True, text=True, timeout=timeout_seconds, check=False)`.
- Non-zero exit, timeout, or invalid JSON returns skipped with reason starting `codex exec`.
- Valid JSON contains `target_field` and `semantic_score`; clamp score to `0.0-1.0`; return completed candidate with source `llm`.

**Fusion Upgrade Acceptance:**

- `FusionMatcher.__init__` accepts `enable_embedding`, `enable_llm`, `embedding_matcher=None`, and `llm_matcher=None`.
- When `enable_embedding=True` and matcher exists, call `embedding_matcher.match(profile)` after RapidFuzz.
- When `enable_llm=True` and matcher exists, call `llm_matcher.match(profile, semantic_candidates)` after embedding.
- `semantic_candidates` are embedding candidates when present; otherwise the best candidates from earlier deterministic stages.
- Stage order remains `exact`, `dictionary`, `rapidfuzz`, `embedding`, `llm`.
- Ranking priority becomes `exact=5`, `dictionary=4`, `llm=3`, `embedding=2`, `rapidfuzz=1`.
- `match_profiles` forwards semantic flags and injected matchers while preserving Phase 1 defaults as false.

**Commands:**

```bash
pytest tests/test_semantic_stages.py::test_llm_matcher_skips_without_candidates tests/test_semantic_stages.py::test_llm_matcher_prompt_includes_data_samples_and_parses_json -q
pytest tests/test_services.py::test_phase2_fusion_calls_embedding_then_llm_with_profile_samples -q
pytest tests/test_services.py -q
```

Expected RED before implementation: missing LLM matcher and fusion injection support.

Expected PASS after implementation: semantic tests and services tests pass without real `codex exec`.

**Commit:**

```bash
git add excel_matcher/matcher/llm_matcher.py excel_matcher/matcher/fusion_matcher.py excel_matcher/services/matching_service.py tests/test_semantic_stages.py tests/test_services.py
git commit -m "feat: enable phase 2 semantic fusion"
```

---

## G10: Phase 2 Verification Gate

**Objective:** Prove full MVP passes with deterministic defaults and offline semantic tests.

**Commands:**

```bash
pytest
python -m compileall excel_matcher tests
pytest tests/test_semantic_stages.py tests/test_services.py::test_phase2_fusion_calls_embedding_then_llm_with_profile_samples -q
git status --short
```

**Acceptance:**

- Full test suite passes.
- Compile checks pass.
- Semantic-specific tests pass using fakes or runner injection.
- No test downloads embedding models.
- No test invokes real `codex exec`.
- `git status --short` shows no uncommitted implementation files.

**Goal Completion Rule:**

Only after this gate passes may the goal be marked complete. If Phase 2 is intentionally deferred, mark the goal complete only for a narrowed Phase 1 objective and leave G8-G10 unchecked with a clear note.

---

## Blocked Notes

Record blockers here only when progress cannot continue without user input or an external state change.

- None.

## Completion Summary Template

Use this when finishing a goal run:

```text
Completed packages:
- G...

Verification:
- pytest: PASS
- python -m compileall excel_matcher tests: PASS
- semantic offline tests: PASS or not run because Phase 2 deferred

Commits:
- <hash> <message>

Remaining:
- None, or list deferred Phase 2 work.
```
