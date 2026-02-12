# pptx-ooxml-engine

用于生成和改写 `.pptx` 的原生 OOXML 执行引擎。

`pptx-ooxml-engine` 通过结构化操作计划执行 PPT 处理，特点是：

- 原生 OOXML 操作
- 不经过 HTML 转换
- 执行结果可确定、可复现
- 不绑定任何模型框架

完整规格文档：`docs/specification.md`

## 定位

本库是 **执行层**，用于承接上层规划器/Agent 输出的操作计划。

- 不负责生成策略或内容规划
- 只负责按计划执行
- `template_pptx` 指母版模板（master/layout），不是复用页

## 功能（v1.2.0）

结构管理：
- `copy_slide`（`part` / `shape`）
- `create_slide_on_layout`
- `delete_slide`
- `move_slide`
- `set_slide_size`
- `set_slide_layout`
- `set_notes`

内容生成与编辑：
- `rewrite_text`
- `add_textbox`（支持段落样式与列表类型）
- `set_shape_text`
- `add_image`（`stretch` / `contain` / `cover`）
- `add_shape`
- `add_table`
- `set_slide_background`
- `fill_placeholder`
- `add_chart`
- `replace_image`
- `set_table_cell`
- `merge_table_cells`
- `set_shape_hyperlink`

排版：
- `align_shapes`
- `distribute_shapes`
- `set_shape_geometry`
- `set_shape_z_order`

质量保障：
- `verify_pptx`
- Python API + CLI
- 可运行示例

仅在使用 `copy_slide` 时才需要 `pptx-copy-ops` 依赖。

## 安装

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
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "执行摘要"},
            {"op": "set_slide_background", "slide_index": 0, "color_hex": "0B1D3A"},
            {
                "op": "add_textbox",
                "slide_index": 0,
                "x_inches": 0.8,
                "y_inches": 1.6,
                "width_inches": 5.8,
                "height_inches": 2.8,
                "paragraphs": [
                    {"text": "关键要点", "font_size_pt": 24, "bold": True},
                    {"text": "支持复用页库", "list_type": "bullet", "level": 1},
                    {"text": "原生 OOXML 组合", "list_type": "number", "level": 1}
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

如果 `ops.json` 内已包含 `template_pptx`，可不传 `--template`。

## 示例

```bash
cd pptx-ooxml-engine
./examples/run_examples.sh
```

## 测试

```bash
cd pptx-ooxml-engine
pytest -q
```

## 构建

```bash
cd pptx-ooxml-engine
python -m build
```

## 许可证

MIT
