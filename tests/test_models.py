from __future__ import annotations


def test_parse_copy_and_create_ops() -> None:
    from pptx_ooxml_engine.models import parse_ops, parse_plan

    plan = parse_plan(
        {
            "template_pptx": "/tmp/template_master.pptx",
            "reuse_slide_libraries": ["/tmp/reuse_library_a.pptx"],
            "operations": [
                {
                    "op": "copy_slide",
                    "reuse_library_index": 0,
                    "source_slide_index": 0,
                    "mode": "part",
                },
                {
                    "op": "delete_slide",
                    "slide_index": 0
                },
                {
                    "op": "create_slide_on_layout",
                    "layout_index": 0,
                    "title": "新页标题",
                    "body": "正文",
                },
                {
                    "op": "move_slide",
                    "from_index": 0,
                    "to_index": 1
                },
                {
                    "op": "set_slide_size",
                    "preset": "4:3"
                },
                {
                    "op": "set_slide_layout",
                    "slide_index": 0,
                    "layout_index": 1
                },
                {
                    "op": "set_notes",
                    "slide_index": 0,
                    "text": "讲稿"
                }
            ],
        }
    )

    ops = parse_ops(plan.model_dump())
    assert plan.template_pptx == "/tmp/template_master.pptx"
    assert plan.reuse_slide_libraries == ["/tmp/reuse_library_a.pptx"]
    assert len(ops) == 7
    assert ops[0].op == "copy_slide"
    assert ops[0].reuse_library_index == 0
    assert ops[1].op == "delete_slide"
    assert ops[2].op == "create_slide_on_layout"
    assert ops[3].op == "move_slide"
    assert ops[4].op == "set_slide_size"
    assert ops[5].op == "set_slide_layout"
    assert ops[6].op == "set_notes"
