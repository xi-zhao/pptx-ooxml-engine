# pptx-ooxml-engine v1.0 Core Ops Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Deliver `v1.0.0` with native OOXML-first structure/content/layout operations suitable for LLM orchestration and community release.

**Architecture:** Keep `engine` as an operation dispatcher with deterministic, schema-validated ops. Extend operation models and JSON schema in lock-step, then implement each op with `python-pptx` plus minimal direct OOXML edits where needed (list bullets/numbering). Preserve template/master semantics.

**Tech Stack:** Python 3.10+, `python-pptx`, `pydantic`, pytest.

### Task 1: Add failing tests for v1 core operations

**Files:**
- Modify: `tests/test_engine.py`
- Modify: `tests/test_models.py`

**Step 1: Write failing tests**
- Add tests covering:
- `apply_ops` no longer requires `pptx-copy-ops` when `copy_slide` is unused.
- New ops end-to-end:
- `add_textbox` with paragraphs (multi-level list style).
- `set_shape_text` on existing shape.
- `add_shape`.
- `align_shapes`.
- `distribute_shapes`.
- `add_table`.
- `add_image`.
- `set_slide_background`.

**Step 2: Run tests and observe failures**

Run: `pytest -q tests/test_engine.py tests/test_models.py`
Expected: Failures due to unsupported op/model/schema behavior.

### Task 2: Implement new models and engine handlers

**Files:**
- Modify: `src/pptx_ooxml_engine/models.py`
- Modify: `src/pptx_ooxml_engine/engine.py`

**Step 1: Add operation models**
- Add typed models and validators for:
- `add_textbox`
- `set_shape_text`
- `add_image`
- `add_shape`
- `add_table`
- `set_slide_background`
- `align_shapes`
- `distribute_shapes`

**Step 2: Add engine helpers and dispatch**
- Implement robust helpers for:
- color parsing
- shape targeting by name/index
- text paragraph write pipeline
- bullet/number OOXML paragraph markers
- geometry conversions via inches
- Implement op handlers and register in dispatcher.
- Change copy-op loading to lazy path: only import `pptx-copy-ops` when a `copy_slide` op is executed.

**Step 3: Re-run focused tests**

Run: `pytest -q tests/test_engine.py tests/test_models.py`
Expected: PASS.

### Task 3: Update schema, docs, examples, and API version

**Files:**
- Modify: `src/pptx_ooxml_engine/schemas/ops.v1.json`
- Modify: `docs/specification.md`
- Modify: `README.en.md`
- Modify: `README.zh-CN.md`
- Modify: `src/pptx_ooxml_engine/__init__.py`
- Modify: `pyproject.toml`
- Modify: `examples/ops/README.md`
- Modify: `examples/ops/01_template_only_generate.json`
- Modify: `examples/ops/02_template_plus_reuse_library.json`
- Modify: `src/pptx_ooxml_engine/examples_runner.py`

**Step 1: Schema + docs**
- Add JSON schema entries for all new operations with required fields and enums.
- Update spec/readmes with operation definitions and usage snippets.

**Step 2: Examples + version**
- Ensure examples demonstrate both structure/reuse and new generated-content operations.
- Bump package version to `1.0.0` consistently in code and metadata.

**Step 3: Run docs/examples related tests**

Run: `pytest -q tests/test_schema.py tests/test_examples.py tests/test_examples_runner.py tests/test_api_surface.py tests/test_cli.py`
Expected: PASS.

### Task 4: Full verification and release

**Files:**
- No code changes expected unless verification uncovers regressions.

**Step 1: Full test and example generation**

Run:
- `pytest -q`
- `./examples/run_examples.sh`

Expected: all tests pass, examples generated successfully.

**Step 2: Release workflow**
- Commit all changes.
- Push working branch.
- Sync `pptx-ooxml-engine` GitHub repo `main`.
- Tag and publish `v1.0.0` release with changelog highlights.
