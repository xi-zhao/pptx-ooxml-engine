# Ops Examples

This folder contains minimal operation-plan examples for `pptx-ooxml-engine`.

## Files

- `01_template_only_generate.json`:
  - Generate from template only (no reuse library).
  - Demonstrates: `delete_slide`, `create_slide_on_layout`, `set_slide_background`,
    `add_shape`, `align_shapes`, `distribute_shapes`, `add_textbox`, `add_table`.
- `02_template_plus_reuse_library.json`:
  - Generate with template + reusable slide library.
  - Demonstrates `copy_slide` plus newly generated pages via `add_textbox` and `add_shape`.

## Run

```bash
cd pptx-ooxml-engine
./examples/run_examples.sh
```

Notes:
- `--template` is optional if `template_pptx` is already present in `ops-file`.
- `template_pptx` means a master/layout template PPTX, not a reusable page itself.
- `pptx-copy-ops` is only required for examples that include `copy_slide`.
- Generated files default to `examples/generated/` (override by passing output dir to script).
