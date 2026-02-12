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
                    "op": "create_slide_on_layout",
                    "layout_index": 0,
                    "title": "新页标题",
                    "body": "正文",
                },
            ],
        }
    )

    ops = parse_ops(plan.model_dump())
    assert plan.template_pptx == "/tmp/template_master.pptx"
    assert plan.reuse_slide_libraries == ["/tmp/reuse_library_a.pptx"]
    assert len(ops) == 2
    assert ops[0].op == "copy_slide"
    assert ops[0].reuse_library_index == 0
    assert ops[1].op == "create_slide_on_layout"
