# pptx-ooxml-engine v1.2 Content Finishing Roadmap & Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Close key execution-layer gaps after v1.1 by adding media/table/link operations needed for enterprise deck production.

**Architecture:** Keep planner/runtime separation. Add atomic operations that are deterministic and composable by LLM planners. Avoid GUI-only features and keep semantics at native OOXML level.

**Tech Stack:** Python 3.10+, python-pptx, pydantic, pytest.

## 1) Gap Analysis (post v1.1)

### P0 (next batch)
- Replace existing image while preserving frame geometry.
- Table cell-level editing and merge operations.
- Shape hyperlink operation.

### P1 (near-term)
- Table border/fill/alignment style ops.
- Existing chart data/style mutation ops.
- Shape grouping/ungrouping (if feasible with OOXML tree operations).

### P2 (mid-term)
- Section management (presentation section list).
- Theme/master mutation ops (colors/fonts/background presets).
- Quality checks: text overflow/readability metrics as report API.

### P3 (optional platform adapters)
- Export adapters (PDF/image) via external tools only.
- Font/media compression adapters.

## 2) v1.2 Scope

Operations to add:
- `replace_image`
- `set_table_cell`
- `merge_table_cells`
- `set_shape_hyperlink`

Non-goals for v1.2:
- chart mutation
- theme/master editing
- section management

## 3) Implementation Tasks

### Task 1: Add failing tests

**Files:**
- Modify: `tests/test_engine.py`
- Modify: `tests/test_models.py`

Steps:
1. Add end-to-end test covering all four new operations.
2. Add model parse/validation tests.
3. Run:
   - `pytest -q tests/test_engine.py tests/test_models.py`
4. Confirm failures due to unsupported ops.

### Task 2: Implement models and handlers

**Files:**
- Modify: `src/pptx_ooxml_engine/models.py`
- Modify: `src/pptx_ooxml_engine/engine.py`

Steps:
1. Add models + validators:
- `ReplaceImageOp`
- `SetTableCellOp`
- `MergeTableCellsOp`
- `SetShapeHyperlinkOp`
2. Implement engine handlers with strict target resolution and bounds checks.
3. Re-run targeted tests until green.

### Task 3: Update schema/docs/version

**Files:**
- Modify: `src/pptx_ooxml_engine/schemas/ops.v1.json`
- Modify: `tests/test_schema.py`
- Modify: `README.en.md`
- Modify: `README.zh-CN.md`
- Modify: `docs/specification.md`
- Modify: `src/pptx_ooxml_engine/__init__.py`
- Modify: `pyproject.toml`

Steps:
1. Add schema entries for 4 new ops.
2. Update docs and signatures.
3. Bump version to `1.2.0`.

### Task 4: Verify and release

Steps:
1. Run:
- `pytest -q`
- `./examples/run_examples.sh`
2. Commit and push working branch.
3. Sync GitHub `main`.
4. Tag and publish `v1.2.0`.
