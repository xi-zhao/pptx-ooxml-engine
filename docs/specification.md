# pptx-ooxml-engine Specification (v0.1.0)

Status: Draft (implementation-aligned)  
Last Updated: 2026-02-12

## 1. Purpose / 目标

`pptx-ooxml-engine` 是一个 **原生 OOXML 的 PPTX 生成与改写执行引擎**。

它的核心职责是：

- 接收结构化操作计划（operation plan）
- 在 `.pptx` 上按顺序执行操作
- 产出新的 `.pptx`
- 在需要时进行结构完整性校验

它**不负责**：

- LLM 内容生成
- 知识库检索与推理
- 页面排版策略规划（上层 planner 负责）

## 2. Scope / 边界定义

### 2.1 In Scope

- 基于模板（master/layout）进行页面生成和改写
- 跨文件复制页面（`copy_slide`）
- 通过 layout 新建页面（`create_slide_on_layout`）
- 文本替换改写（`rewrite_text`）
- 删除页面（`delete_slide`）
- 页面重排（`move_slide`）
- 页面尺寸设置（`set_slide_size`）
- 页面版式重设（`set_slide_layout`）
- 备注写入（`set_notes`）
- OOXML 结构级校验（`verify_pptx`）
- Python API + CLI 执行入口

### 2.2 Out of Scope

- 智能版式决策（例如“选哪个 layout 最优”）
- 深度图文自动设计
- 复杂动画与时间线编排
- 外部 Office GUI 自动化验证（当前版本未内置）

## 3. Terminology / 术语

- `template_pptx`: **母版模板文件**（master/layout 来源），不是复用页本身。
- `reuse_slide_libraries`: 可复用页面库列表（历史 PPT 或页库文件）。
- `operation`: 单个原子操作（copy/create/rewrite/delete/move/size/layout/notes）。
- `operation plan`: 顶层执行计划对象，包含模板、页库和操作列表。
- `verify`: 执行后结构校验步骤。

## 4. Runtime Dependencies / 运行依赖

- Python `>=3.10`
- `python-pptx >= 1.0.2`
- `pydantic >= 2.8.0`
- `pptx-copy-ops`（仅 `copy_slide` 所需）

## 5. Data Model / 数据模型

顶层计划对象（`OperationPlan`）：

```json
{
  "template_pptx": "path/to/template.pptx",
  "reuse_slide_libraries": ["path/to/reuse_lib_1.pptx"],
  "operations": []
}
```

字段规范：

| Field | Type | Required | Default | Description |
|---|---|---|---|---|
| `template_pptx` | `string \| null` | No | `null` | master/layout 模板路径 |
| `reuse_slide_libraries` | `string[]` | No | `[]` | 复用页库路径列表 |
| `operations` | `Operation[]` | Yes | - | 执行操作列表（顺序执行） |

## 6. Operation Specs / 操作规格

## 6.1 `copy_slide`

用途：将外部页面复制到当前输出演示文稿。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"copy_slide"` | Yes | 操作类型 |
| `source_path` | `string \| null` | Conditional | 源 PPTX 路径（与 `reuse_library_index` 二选一） |
| `reuse_library_index` | `int \| null` (`>=0`) | Conditional | 引用 `reuse_slide_libraries` 的索引（与 `source_path` 二选一） |
| `source_slide_index` | `int` (`>=0`) | Yes | 源页 0-based 索引 |
| `mode` | `"shape" \| "part"` | No | 默认 `part` |

语义：

- `mode=part`: 高保真复制，可能引入额外 layout/master。
- `mode=shape`: 倾向模板统一，视觉可能略有差异。

错误条件：

- `source_path` 和 `reuse_library_index` 同时缺失。
- `reuse_library_index` 越界。
- `source_slide_index` 越界（由底层复制库抛错）。

## 6.2 `create_slide_on_layout`

用途：基于当前模板指定 layout 新建页面。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"create_slide_on_layout"` | Yes | 操作类型 |
| `layout_index` | `int` (`>=0`) | No | 默认 `0` |
| `title` | `string \| null` | No | 如有标题占位符则写入 |
| `body` | `string \| null` | No | 优先写 BODY 占位符，否则写入第一个可写文本框 |

错误条件：

- `layout_index` 越界。

## 6.3 `rewrite_text`

用途：在指定页面中进行文本替换。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"rewrite_text"` | Yes | 操作类型 |
| `slide_index` | `int` (`>=0`) | Yes | 目标页 0-based 索引 |
| `find` | `string` | Yes | 查找文本 |
| `replace` | `string` | Yes | 替换文本 |
| `shape_name` | `string \| null` | No | 限定 shape 名称 |
| `occurrence` | `"first" \| "all"` | No | 默认 `all` |

行为说明：

- 替换以 `shape.text_frame.text` 粒度执行（段落样式不做细粒保留）。
- 若找不到可替换文本，抛 `ValueError`。

## 6.4 `delete_slide`

用途：按索引删除当前输出中的页面。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"delete_slide"` | Yes | 操作类型 |
| `slide_index` | `int` (`>=0`) | Yes | 要删除页面的 0-based 索引 |

行为说明：

- 删除通过移除 slide 关系与 slide id 列表条目完成。
- 典型用途是移除模板初始页。

错误条件：

- `slide_index` 越界。

## 6.5 `move_slide`

用途：在演示文稿内部移动页面顺序。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"move_slide"` | Yes | 操作类型 |
| `from_index` | `int` (`>=0`) | Yes | 原位置索引 |
| `to_index` | `int` (`>=0`) | Yes | 目标位置索引 |

错误条件：

- `from_index` 或 `to_index` 越界。

## 6.6 `set_slide_size`

用途：设置演示文稿画布尺寸。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"set_slide_size"` | Yes | 操作类型 |
| `preset` | `"16:9" \| "4:3" \| null` | Conditional | 预设尺寸（与自定义尺寸二选一） |
| `width_inches` | `float \| null` (`>0`) | Conditional | 自定义宽度（英寸） |
| `height_inches` | `float \| null` (`>0`) | Conditional | 自定义高度（英寸） |

行为说明：

- `preset=16:9` -> `13.333 x 7.5` 英寸
- `preset=4:3` -> `10 x 7.5` 英寸

错误条件：

- 同时设置 `preset` 与 `width/height`
- 未设置 `preset` 且未提供 `width/height`

## 6.7 `set_slide_layout`

用途：将指定页面的 layout 关系重绑到目标 layout。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"set_slide_layout"` | Yes | 操作类型 |
| `slide_index` | `int` (`>=0`) | Yes | 目标页面 |
| `layout_index` | `int` (`>=0`) | Yes | 目标 layout 索引 |

错误条件：

- `slide_index` 越界
- `layout_index` 越界

## 6.8 `set_notes`

用途：设置页面备注（讲稿区）文本。

字段：

| Field | Type | Required | Description |
|---|---|---|---|
| `op` | `"set_notes"` | Yes | 操作类型 |
| `slide_index` | `int` (`>=0`) | Yes | 目标页面 |
| `text` | `string` | Yes | 备注文本 |

错误条件：

- `slide_index` 越界。

## 7. Execution Semantics / 执行语义

引擎入口：`apply_ops(...)` 或 `generate_pptx(...)`。

执行顺序：

1. 解析 operation plan
2. 解析模板路径（优先级见 7.1）
3. 加载目标模板为输出基底（保留已有页）
4. 按 `operations` 顺序逐条执行
5. 保存到 `output_pptx`
6. 可选执行 `verify_pptx`

### 7.1 Template Resolution Priority

模板路径解析优先级（高 -> 低）：

1. `apply_ops(..., template_pptx=...)`
2. `apply_ops(..., input_pptx=...)`（兼容别名）
3. `ops.template_pptx`

如果三者都缺失，报错：`template_pptx is required`。

## 8. Verification Spec / 校验规格

`verify_pptx(path)` 返回 `VerifyReport`：

- `issues: list[str]`
- `ok: bool` (`issues` 为空时为 `True`)

校验项（v0.1）：

- 文件能被 `python-pptx` 打开
- `presentation.xml` 存在
- 每个 slide 的关系引用不悬空（dangling `r:id`）
- 每个 slide 存在 `slideLayout` 关系
- 每个 layout 存在 `slideMaster` 关系
- 使用到的 master 均已在 presentation 中注册

## 9. API Spec / Python API

## 9.1 `apply_ops`

```python
apply_ops(
    input_pptx: str | Path | None,
    ops: Iterable[Operation] | list[dict] | dict,
    output_pptx: str | Path | None,
    verify: bool = False,
    strict_verify: bool = True,
    template_pptx: str | Path | None = None,
) -> ApplyResult
```

返回值 `ApplyResult`：

- `output_path: Path`
- `operations_applied: int`
- `verify_issues: list[str]`

## 9.2 `generate_pptx`

```python
generate_pptx(
    template_pptx: str | Path,
    ops: Iterable[Operation] | list[dict] | dict,
    output_pptx: str | Path,
    verify: bool = False,
    strict_verify: bool = True,
) -> ApplyResult
```

说明：`generate_pptx` 是面向“模板驱动生成”语义的主入口。

## 9.3 Other Public APIs

- `parse_plan(raw) -> OperationPlan`
- `parse_ops(raw) -> list[Operation]`
- `load_ops_schema(version="v1") -> dict`
- `verify_pptx(path) -> VerifyReport`
- `generate_example_outputs(output_dir) -> list[Path]`

## 10. CLI Spec

命令：

```bash
python -m pptx_ooxml_engine.cli \
  --ops-file ops.json \
  --output out.pptx \
  [--template template.pptx] \
  [--verify] \
  [--no-strict-verify]
```

参数：

| Arg | Required | Description |
|---|---|---|
| `--ops-file` | Yes | operation plan JSON 文件 |
| `--output` | Yes | 输出 pptx 路径 |
| `--template` | No | 覆盖 `ops.template_pptx` |
| `--verify` | No | 执行后运行结构校验 |
| `--no-strict-verify` | No | 即使校验有 issue 也不失败 |
| `--version` | No | 输出版本 |

兼容参数：

- `--input` 为 `--template` 的隐藏别名（兼容老用法）。

## 11. JSON Schema

路径：

- `src/pptx_ooxml_engine/schemas/ops.v1.json`

版本管理建议：

- 不破坏兼容的新增字段：维持 `v1`
- 破坏兼容：新增 `ops.v2.json`

## 12. Error Model / 错误模型

常见异常：

- `ModuleNotFoundError`: 缺少 `pptx-copy-ops`（在 `copy_slide` 场景）
- `ValueError`: 参数不合法（如无模板、替换未命中等）
- `IndexError`: slide/layout/library 索引越界
- `ValueError("verification failed: ...")`: `verify=True` 且 `strict_verify=True` 且校验失败

## 13. Determinism / 确定性

在同样输入文件、同样操作计划、同样依赖版本下，执行顺序与结果应保持稳定。  
引擎不引入随机行为。

## 14. Known Limitations / 已知限制

- `rewrite_text` 以整 text_frame 字符串替换，非富文本级编辑
- 不含页面删除/重排等高级结构操作
- 未内置 Office GUI 实机打开验证
- 不负责自动选择“最佳 layout”与美学排版策略

## 15. Layering Recommendation / 分层建议

- `pptx-ooxml-engine`: 执行层（本库）
- 上层 planner/agent: 决策层（大纲、内容、版式策略、检索与推理）

该分层是长期可维护和可开源协作的推荐架构。
