# pptx-ooxml-engine v1.1 Placeholder/Chart/Layout Ops Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Release `v1.1.0` with native OOXML operations for placeholder filling, chart creation, shape geometry adjustment, and z-order control.

**Architecture:** Extend operation union/types/schema first, then implement deterministic handlers in `engine.py` using `python-pptx` + minimal OOXML tree reorder for z-order.

**Tech Stack:** Python 3.10+, python-pptx, pydantic, pytest.

### Task 1: Write failing tests for new ops

**Files:**
- Modify: `tests/test_engine.py`
- Modify: `tests/test_models.py`

**Steps:**
1. Add `test_apply_ops_v1_1_placeholder_chart_and_layer_ops`.
2. Add parse/model test coverage for new op types.
3. Run:
   - `pytest -q tests/test_engine.py tests/test_models.py`
4. Confirm failures due to unsupported ops.

### Task 2: Implement models and engine handlers

**Files:**
- Modify: `src/pptx_ooxml_engine/models.py`
- Modify: `src/pptx_ooxml_engine/engine.py`

**Steps:**
1. Add operation models:
   - `fill_placeholder`
   - `set_shape_geometry`
   - `set_shape_z_order`
   - `add_chart`
2. Implement engine handlers:
   - placeholder locate by index/type
   - text/image placeholder fill
   - geometry adjustment for existing shapes
   - z-order manipulation in shape tree
   - chart insertion using `CategoryChartData`
3. Run targeted tests until green.

### Task 3: Update schema/docs/version

**Files:**
- Modify: `src/pptx_ooxml_engine/schemas/ops.v1.json`
- Modify: `tests/test_schema.py`
- Modify: `README.en.md`
- Modify: `README.zh-CN.md`
- Modify: `docs/specification.md`
- Modify: `src/pptx_ooxml_engine/__init__.py`
- Modify: `pyproject.toml`

**Steps:**
1. Add JSON schema blocks for new operations.
2. Update docs/spec with signatures and behavior notes.
3. Bump version to `1.1.0`.

### Task 4: Full verification and release

**Files:**
- N/A (verification + git/release operations)

**Steps:**
1. Run:
   - `pytest -q`
   - `./examples/run_examples.sh`
2. Commit and push branch updates.
3. Sync GitHub repo `xi-zhao/pptx-ooxml-engine` `main`.
4. Tag and publish `v1.1.0`.
