# pptx-ooxml-engine

Native OOXML execution engine for generating and editing `.pptx` files.

`pptx-ooxml-engine` runs structured operation plans against PowerPoint files with:

- native OOXML manipulation
- no HTML-to-PPTX conversion
- deterministic execution
- model-agnostic runtime

Full specification: `docs/specification.md`

## Positioning

This package is the **execution layer**.

It does not implement LLM planning logic.  
It executes operation plans produced by upstream planners/agents/workflows.

## Core Concepts

- `template_pptx`: master/layout template (not a reusable page)
- `reuse_slide_libraries`: reusable slide library files
- `operations`: ordered atomic operations

## Features (v0.1)

- `copy_slide` (`part` / `shape`)
- `create_slide_on_layout`
- `rewrite_text`
- `delete_slide`
- `move_slide`
- `set_slide_size` (`16:9` / `4:3` / custom)
- `set_slide_layout`
- `set_notes`
- `verify_pptx`
- Python API + CLI
- Runnable examples

## Installation

```bash
pip install pptx-ooxml-engine
```

For `copy_slide`, runtime requires `pptx-copy-ops`.

## Python API

```python
from pptx_ooxml_engine import generate_pptx

result = generate_pptx(
    template_pptx="resources/theme1.pptx",
    ops={
        "reuse_slide_libraries": ["resources/theme2.pptx"],
        "operations": [
            {"op": "delete_slide", "slide_index": 0},
            {"op": "set_slide_size", "preset": "16:9"},
            {"op": "copy_slide", "reuse_library_index": 0, "source_slide_index": 0, "mode": "shape"},
            {"op": "move_slide", "from_index": 0, "to_index": 1},
            {"op": "set_slide_layout", "slide_index": 1, "layout_index": 1},
            {"op": "set_notes", "slide_index": 1, "text": "Speaker notes here"},
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "New Slide", "body": "Body"},
            {"op": "rewrite_text", "slide_index": 1, "find": "Old", "replace": "New"}
        ]
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
```

`--template` is optional if `template_pptx` already exists in `ops.json`.

## Examples

```bash
cd pptx-ooxml-engine
./examples/run_examples.sh
```

## Testing

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
