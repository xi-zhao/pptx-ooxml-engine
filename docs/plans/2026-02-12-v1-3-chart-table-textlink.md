# pptx-ooxml-engine v1.3 Chart/Table/Text-Link Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Release `v1.3.0` with chart update, table style/sizing, and run-level text hyperlink operations.

**Architecture:** Extend operation model/schema with atomic mutations over existing chart/table/text objects; reuse existing target-resolution patterns (`shape_name`/`shape_index`, `table_name`/`table_index`) for consistency and LLM-call stability.

**Tech Stack:** Python 3.10+, python-pptx, pydantic, pytest.

## Scope

New operations:
- `update_chart_data`
- `set_table_style`
- `set_table_row_col_size`
- `set_text_hyperlink`

Out of scope:
- section management
- theme/master mutation
- chart visual theme customization beyond data update

## Tasks

### Task 1: Failing tests first

**Files:**
- `tests/test_engine.py`
- `tests/test_models.py`

Steps:
1. Add end-to-end test covering all new operations together.
2. Add model parse/validation coverage.
3. Run `pytest -q tests/test_engine.py tests/test_models.py` and confirm failures.

### Task 2: Implement operations

**Files:**
- `src/pptx_ooxml_engine/models.py`
- `src/pptx_ooxml_engine/engine.py`

Steps:
1. Add typed models and validators for the 4 new ops.
2. Implement handlers:
- chart replace-data pipeline
- table-wide style mutation
- table row/column size mutation
- run-level hyperlink mutation
3. Re-run focused tests until green.

### Task 3: Schema/docs/version sync

**Files:**
- `src/pptx_ooxml_engine/schemas/ops.v1.json`
- `tests/test_schema.py`
- `README.en.md`
- `README.zh-CN.md`
- `docs/specification.md`
- `src/pptx_ooxml_engine/__init__.py`
- `pyproject.toml`

Steps:
1. Add schema specs for new ops.
2. Update docs/signatures/examples text.
3. Bump version to `1.3.0`.

### Task 4: Verify and release

Steps:
1. Run:
- `pytest -q`
- `./examples/run_examples.sh`
2. Commit and push current branch.
3. Sync `xi-zhao/pptx-ooxml-engine` main.
4. Tag and publish `v1.3.0`.
