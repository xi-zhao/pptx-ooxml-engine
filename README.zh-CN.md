# pptx-ooxml-engine

用于生成和改写 `.pptx` 的原生 OOXML 执行引擎。

`pptx-ooxml-engine` 通过结构化操作计划执行 PPT 处理，特点是：

- 原生 OOXML 操作
- 不经过 HTML 转换
- 执行过程可确定、可复现
- 不绑定任何模型框架

完整规格文档：`docs/specification.md`

## 定位

本库是 **执行层**。

它不负责 LLM 规划，只负责执行上层规划器/工作流产出的操作计划。

## 核心概念

- `template_pptx`：母版模板（master/layout），不是复用页
- `reuse_slide_libraries`：复用页库文件列表
- `operations`：按顺序执行的原子操作

## 功能（v0.1）

- `copy_slide`（`part` / `shape`）
- `create_slide_on_layout`
- `rewrite_text`
- `delete_slide`
- `move_slide`
- `set_slide_size`（`16:9` / `4:3` / 自定义）
- `set_slide_layout`
- `set_notes`
- `verify_pptx`
- Python API + CLI
- 可运行示例

## 安装

```bash
pip install pptx-ooxml-engine
```

如果使用 `copy_slide`，运行环境还需要 `pptx-copy-ops`。

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
            {"op": "set_notes", "slide_index": 1, "text": "这里填讲稿备注"},
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "新页面", "body": "正文"},
            {"op": "rewrite_text", "slide_index": 1, "find": "旧词", "replace": "新词"}
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

如果 `ops.json` 内已包含 `template_pptx`，则可不传 `--template`。

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
