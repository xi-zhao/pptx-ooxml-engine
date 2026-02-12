# Ops Examples

This folder contains minimal operation-plan examples for `pptx-ooxml-engine`.

## Files

- `01_template_only_generate.json`:
  - Generate slides from a master/layout template only.
- `02_template_plus_reuse_library.json`:
  - Generate with both template and reusable slide library.
  - Uses `copy_slide` + `reuse_library_index`.

## Run

```bash
cd pptx-ooxml-engine
./examples/run_examples.sh
```

Notes:
- `--template` is optional if `template_pptx` is already present in `ops-file`.
- `template_pptx` means a master/layout template PPTX, not a reusable page itself.
- Generated files default to `examples/generated/` (override by passing output dir to script).
