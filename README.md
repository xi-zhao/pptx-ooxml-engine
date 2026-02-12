# pptx-ooxml-engine

Model-agnostic native PPTX OOXML execution engine.

`pptx-ooxml-engine` executes structured operations against `.pptx` files with native OOXML manipulation only:

- no HTML-to-PPTX conversion
- no embedded LLM orchestration
- deterministic operation runtime

Full specification: `docs/specification.md`

`pptx-ooxml-engine` 是一个面向 Agent/LLM 的底层执行引擎，只负责“操作执行”，不负责“智能决策”。

- 全程原生 OOXML
- 不走 HTML 转换
- 不绑定任何模型或编排框架

## Positioning / 定位

This package is the **execution layer**.

Use it under any planner (LLM, rules, workflow engine) that produces operation JSON.

本库定位是 **执行层**，建议与上层“排版规划层”解耦：

- `pptx-ooxml-engine`: execute `ops`
- upper planner: decide *which* ops to run
- planner output can include explicit top-level fields:
  - `template_pptx` (master/layout template)
  - `reuse_slide_libraries` (reusable page libraries)

## Features (V0.1)

- `copy_slide` (`part` / `shape`)
- `create_slide_on_layout`
- `rewrite_text`
- OOXML verification before output
- CLI + Python API

## Install

```bash
pip install pptx-ooxml-engine
```

For slide-copy operations, install `pptx-copy-ops` as runtime dependency in your environment.

如果你要使用 `copy_slide`，运行环境里需要可用的 `pptx-copy-ops`。

## Python API

```python
from pptx_ooxml_engine import generate_pptx

result = generate_pptx(
    template_pptx="resources/theme1.pptx",  # this is the master/layout template
    ops={
        "reuse_slide_libraries": ["resources/theme2.pptx"],
        "operations": [
            {"op": "copy_slide", "reuse_library_index": 0, "source_slide_index": 0, "mode": "part"},
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "新建页标题", "body": "正文内容"},
            {"op": "rewrite_text", "slide_index": 0, "find": "旧词", "replace": "新词"},
        ],
    },
    output_pptx="output/demo.pptx",
    verify=True,
)
print(result.output_path)
```

## CLI

```bash
python -m pptx_ooxml_engine.cli \
  --ops-file ops.json \
  --output output/demo.pptx \
  --verify

# ops.json can explicitly include template and page libraries:
# {
#   "template_pptx": "resources/theme1.pptx",
#   "reuse_slide_libraries": ["resources/reuse_libraries/library_a.pptx"],
#   "operations": [...]
# }
```

See runnable examples in `examples/ops/`.

Quick demo run:

```bash
cd pptx-ooxml-engine
./examples/run_examples.sh
```

## Ops Schema

- bundled schema loader: `from pptx_ooxml_engine import load_ops_schema`
- schema file: `src/pptx_ooxml_engine/schemas/ops.v1.json`

## Architecture

- `models.py`: typed operation models + parsing
- `engine.py`: deterministic operation executor
- `verify.py`: OOXML integrity checks
- `schema.py`: schema loading helpers

## Design Rules

- Execution-only core (no planner logic)
- `template` means slide master/layout template PPTX, not a source slide library page
- Explicit copy mode:
  - `part`: max visual fidelity
  - `shape`: template-unification friendly
- Verify before finalize

## Test

```bash
cd pptx-ooxml-engine
pytest -q
```

## Build

```bash
cd pptx-ooxml-engine
python -m build
```

## License

MIT
