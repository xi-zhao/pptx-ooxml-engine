# pptx-ooxml-engine Specification (v1.2.0)

## 1. Purpose / 目标

`pptx-ooxml-engine` 是一个原生 OOXML 执行引擎，用于按结构化操作计划生成与改写 `.pptx` 文件。

核心目标：
- 原生操作 PPTX，不走 HTML 渲染链路
- 适配 LLM/Agent 调用（输入是结构化 op plan）
- 保持执行确定性和可验证性

## 2. Positioning / 定位

这是执行层，不是规划层。

- 上层负责：大纲、内容策略、页库匹配、研究信息融合
- 本层负责：按顺序执行 ops，产出可打开且结构正确的 PPTX

`template_pptx` 语义：
- 指母版/版式模板（master/layout template）
- 不是复用页库

## 3. Runtime & Dependencies / 运行依赖

- Python >= 3.10
- `python-pptx>=1.0.2`
- `pydantic>=2.8.0`
- `pptx-copy-ops`：仅在 `copy_slide` 操作出现时按需加载

## 4. Data Model / 数据模型

顶层对象 `OperationPlan`：

```json
{
  "template_pptx": "path/to/template.pptx",
  "reuse_slide_libraries": ["path/to/lib1.pptx"],
  "operations": []
}
```

字段：
- `template_pptx`：可选。模板路径，可被 CLI `--template` 覆盖。
- `reuse_slide_libraries`：可选。复用页库路径数组。
- `operations`：必填。按顺序执行的操作列表。

## 5. Operations / 操作集

### 5.1 Structure Ops / 结构操作

- `copy_slide`
- `create_slide_on_layout`
- `delete_slide`
- `move_slide`
- `set_slide_size`
- `set_slide_layout`
- `set_notes`

### 5.2 Content Ops / 内容操作

- `rewrite_text`
- `add_textbox`
- `set_shape_text`
- `add_image`
- `add_shape`
- `add_table`
- `set_slide_background`
- `fill_placeholder`
- `add_chart`
- `replace_image`
- `set_table_cell`
- `merge_table_cells`
- `set_shape_hyperlink`

### 5.3 Layout Ops / 排版操作

- `align_shapes`
- `distribute_shapes`
- `set_shape_geometry`
- `set_shape_z_order`

## 6. Operation Signatures / 参数签名

### `copy_slide`
- Required:
- `op: "copy_slide"`
- `source_slide_index: int >= 0`
- Optional:
- `source_path: str` 或 `reuse_library_index: int >= 0`（二选一）
- `mode: "part" | "shape"`（默认 `part`）

### `create_slide_on_layout`
- `op: "create_slide_on_layout"`
- `layout_index: int >= 0`（默认 0）
- `title?: str`
- `body?: str`

### `rewrite_text`
- `op: "rewrite_text"`
- `slide_index: int >= 0`
- `find: str`
- `replace: str`
- `shape_name?: str`
- `occurrence?: "first" | "all"`（默认 `all`）

### `delete_slide`
- `op: "delete_slide"`
- `slide_index: int >= 0`

### `move_slide`
- `op: "move_slide"`
- `from_index: int >= 0`
- `to_index: int >= 0`

### `set_slide_size`
- `op: "set_slide_size"`
- Required one of:
- `preset: "16:9" | "4:3"`
- `width_inches: float > 0` + `height_inches: float > 0`

### `set_slide_layout`
- `op: "set_slide_layout"`
- `slide_index: int >= 0`
- `layout_index: int >= 0`

### `set_notes`
- `op: "set_notes"`
- `slide_index: int >= 0`
- `text: str`

### `add_textbox`
- `op: "add_textbox"`
- `slide_index: int >= 0`
- `x_inches, y_inches >= 0`
- `width_inches, height_inches > 0`
- Required one of:
- `text: str`
- `paragraphs: ParagraphSpec[]`
- Optional:
- `name?: str`
- `vertical_anchor?: "top" | "middle" | "bottom"`
- `word_wrap?: bool`

### `set_shape_text`
- `op: "set_shape_text"`
- `slide_index: int >= 0`
- Required one of:
- `shape_name: str`
- `shape_index: int >= 0`
- Required one of:
- `text: str`
- `paragraphs: ParagraphSpec[]`
- Optional:
- `vertical_anchor?: "top" | "middle" | "bottom"`
- `word_wrap?: bool`

### `add_image`
- `op: "add_image"`
- `slide_index: int >= 0`
- `image_path: str`
- `x_inches, y_inches >= 0`
- `width_inches, height_inches > 0`
- `fit?: "stretch" | "contain" | "cover"`（默认 `stretch`）
- `name?: str`

### `add_shape`
- `op: "add_shape"`
- `slide_index: int >= 0`
- `shape_type: "rect" | "round_rect" | "ellipse" | "right_arrow" | "line"`
- `x_inches, y_inches >= 0`
- `width_inches, height_inches > 0`
- Optional:
- `name?: str`
- `text?: str`
- `fill_color_hex?: RRGGBB | #RRGGBB`
- `line_color_hex?: RRGGBB | #RRGGBB`
- `line_width_pt?: float > 0`
- `text_color_hex?: RRGGBB | #RRGGBB`
- `font_size_pt?: float > 0`

### `add_table`
- `op: "add_table"`
- `slide_index: int >= 0`
- `x_inches, y_inches >= 0`
- `width_inches, height_inches > 0`
- `data: string[][]`（非空二维数组）
- Optional:
- `header?: bool`（默认 `false`，`true` 时首行加粗）
- `name?: str`
- `font_size_pt?: float > 0`

### `set_slide_background`
- `op: "set_slide_background"`
- `slide_index: int >= 0`
- `color_hex: RRGGBB | #RRGGBB`

### `fill_placeholder`
- `op: "fill_placeholder"`
- `slide_index: int >= 0`
- Required one of:
- `placeholder_idx: int >= 0`
- `placeholder_type: "title" | "body" | "subtitle" | "picture" | "object"`
- Required one of:
- `text: str`
- `paragraphs: ParagraphSpec[]`
- `image_path: str`

### `add_chart`
- `op: "add_chart"`
- `slide_index: int >= 0`
- `chart_type: "column_clustered" | "line" | "pie"`
- `x_inches, y_inches >= 0`
- `width_inches, height_inches > 0`
- `categories: string[]`（非空）
- `series: ChartSeriesSpec[]`（非空，且每个 series 的 values 长度必须与 categories 一致）
- `name?: str`

### `replace_image`
- `op: "replace_image"`
- `slide_index: int >= 0`
- Required one of:
- `shape_name: str`
- `shape_index: int >= 0`
- `image_path: str`
- `fit?: "stretch" | "contain" | "cover"`（默认 `stretch`）

### `set_table_cell`
- `op: "set_table_cell"`
- `slide_index: int >= 0`
- Required one of:
- `table_name: str`
- `table_index: int >= 0`
- `row: int >= 0`
- `col: int >= 0`
- Optional content/style:
- `text?: str`
- `bold?: bool`
- `italic?: bool`
- `font_size_pt?: float > 0`
- `text_color_hex?: RRGGBB | #RRGGBB`
- `fill_color_hex?: RRGGBB | #RRGGBB`
- `alignment?: "left" | "center" | "right" | "justify"`

### `merge_table_cells`
- `op: "merge_table_cells"`
- `slide_index: int >= 0`
- Required one of:
- `table_name: str`
- `table_index: int >= 0`
- `start_row, start_col, end_row, end_col: int >= 0`
- `end_row >= start_row` 且 `end_col >= start_col`

### `set_shape_hyperlink`
- `op: "set_shape_hyperlink"`
- `slide_index: int >= 0`
- Required one of:
- `shape_name: str`
- `shape_index: int >= 0`
- `url: str`

### `align_shapes`
- `op: "align_shapes"`
- `slide_index: int >= 0`
- `shape_names: string[]`（至少 2 个）
- `align: "left" | "center" | "right" | "top" | "middle" | "bottom"`
- `reference?: "first" | "slide"`（默认 `first`）

### `distribute_shapes`
- `op: "distribute_shapes"`
- `slide_index: int >= 0`
- `shape_names: string[]`（至少 3 个）
- `direction: "horizontal" | "vertical"`

### `set_shape_geometry`
- `op: "set_shape_geometry"`
- `slide_index: int >= 0`
- Required one of:
- `shape_name: str`
- `shape_index: int >= 0`
- Required at least one:
- `x_inches >= 0`
- `y_inches >= 0`
- `width_inches > 0`
- `height_inches > 0`

### `set_shape_z_order`
- `op: "set_shape_z_order"`
- `slide_index: int >= 0`
- Required one of:
- `shape_name: str`
- `shape_index: int >= 0`
- `action: "bring_to_front" | "send_to_back" | "bring_forward" | "send_backward"`

### ParagraphSpec

- `text: str`
- `level?: int(0..8)`
- `list_type?: "none" | "bullet" | "number"`
- `font_size_pt?: float > 0`
- `bold?: bool`
- `italic?: bool`
- `color_hex?: RRGGBB | #RRGGBB`
- `alignment?: "left" | "center" | "right" | "justify"`
- `line_spacing?: float > 0`
- `space_before_pt?: float >= 0`
- `space_after_pt?: float >= 0`

### ChartSeriesSpec

- `name: str`
- `values: float[]`（非空）

## 7. Execution Semantics / 执行语义

入口：
- `apply_ops(...)`
- `generate_pptx(...)`

执行流程：
1. 解析并验证 plan（Pydantic）
2. 解析模板路径（优先级：显式参数 > 兼容输入 > plan 字段）
3. 按序执行 operations
4. 输出 PPTX
5. 可选 `verify_pptx` 校验

`copy_slide` 依赖加载策略：
- 仅当 operations 中存在 `copy_slide` 时加载 `pptx-copy-ops`
- 无 `copy_slide` 时不需要该依赖

## 8. Verification / 校验

`verify_pptx(path)` 做结构一致性校验，主要包括：
- 能否被 `python-pptx` 打开
- slide -> layout -> master 关系完整性
- dangling relationship 检查
- 已使用 master 是否在 `presentation.xml` 注册

## 9. Public API / 对外 API

Python：
- `apply_ops(...) -> ApplyResult`
- `generate_pptx(...) -> ApplyResult`
- `parse_plan(raw) -> OperationPlan`
- `parse_ops(raw) -> list[Operation]`
- `load_ops_schema(version="v1") -> dict`
- `verify_pptx(path) -> VerifyReport`

CLI：

```bash
python -m pptx_ooxml_engine.cli \
  --template path/to/template.pptx \
  --ops-file ops.json \
  --output output.pptx \
  --verify
```

参数：
- `--ops-file`：必填，ops JSON
- `--output`：必填，输出文件
- `--template`：可选，覆盖 `ops.template_pptx`
- `--verify`：可选，执行校验
- `--no-strict-verify`：可选，校验报错不终止
- `--version`：输出版本

## 10. Error Model / 错误模型

典型错误：
- `ValueError`：操作参数不合法、文本找不到、shape 定位失败
- `IndexError`：slide/layout/shape 索引越界
- `FileNotFoundError`：图片路径不存在
- `ModuleNotFoundError`：执行 `copy_slide` 但缺少 `pptx-copy-ops`

## 11. Determinism / 确定性

在同样模板文件、同样 ops、同样依赖版本下，输出结果应稳定可复现。

## 12. Layering Recommendation / 分层建议

推荐三层：
- Planner Layer：用户意图 -> 大纲 -> op plan
- Retrieval Layer：知识库/复用页库/deep research
- Execution Layer：`pptx-ooxml-engine`

该库仅承担第三层职责，保持可组合性与可维护性。
