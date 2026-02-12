# pptx-ooxml-engine

Native OOXML execution engine for generating and editing `.pptx`.

`pptx-ooxml-engine` executes structured operation plans on PowerPoint files with:

- native OOXML operations
- no HTML-to-PPTX conversion
- deterministic runtime behavior
- model-agnostic execution layer

Full spec: `docs/specification.md`

## Positioning

This package is the **execution layer** for planners/agents.

- It does not decide presentation strategy.
- It executes validated operation plans.
- `template_pptx` means a master/layout template (not reusable slide pages).

## Features (v1.2.0)

Structure:
- `copy_slide` (`part` / `shape`)
- `create_slide_on_layout`
- `delete_slide`
- `move_slide`
- `set_slide_size`
- `set_slide_layout`
- `set_notes`

Content:
- `rewrite_text`
- `add_textbox` (supports paragraph styles and list types)
- `set_shape_text`
- `add_image` (`stretch` / `contain` / `cover`)
- `add_shape`
- `add_table`
- `set_slide_background`
- `fill_placeholder`
- `add_chart`
- `replace_image`
- `set_table_cell`
- `merge_table_cells`
- `set_shape_hyperlink`

Layout:
- `align_shapes`
- `distribute_shapes`
- `set_shape_geometry`
- `set_shape_z_order`

Quality:
- `verify_pptx`
- Python API + CLI
- runnable examples

`pptx-copy-ops` is required only when `copy_slide` is used.

## Installation

```bash
pip install pptx-ooxml-engine
```

## Python API

```python
from pptx_ooxml_engine import generate_pptx

result = generate_pptx(
    template_pptx="resources/theme1.pptx",
    ops={
        "reuse_slide_libraries": ["resources/theme2.pptx"],
        "operations": [
            {"op": "delete_slide", "slide_index": 0},
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "Executive Summary"},
            {"op": "set_slide_background", "slide_index": 0, "color_hex": "0B1D3A"},
            {
                "op": "add_textbox",
                "slide_index": 0,
                "x_inches": 0.8,
                "y_inches": 1.6,
                "width_inches": 5.8,
                "height_inches": 2.8,
                "paragraphs": [
                    {"text": "Key Points", "font_size_pt": 24, "bold": True},
                    {"text": "Reusable slide library enabled", "list_type": "bullet", "level": 1},
                    {"text": "Native OOXML composition", "list_type": "number", "level": 1}
                ]
            },
            {"op": "copy_slide", "reuse_library_index": 0, "source_slide_index": 0, "mode": "shape"}
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
  --template resources/theme1.pptx \
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
