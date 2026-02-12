from __future__ import annotations


def test_load_ops_schema() -> None:
    from pptx_ooxml_engine.schema import load_ops_schema

    schema = load_ops_schema()
    assert schema["type"] == "object"
    assert "operations" in schema["properties"]
    assert "template_pptx" in schema["properties"]
    assert "reuse_slide_libraries" in schema["properties"]
    one_of = schema["properties"]["operations"]["items"]["oneOf"]
    op_consts = {
        item.get("properties", {}).get("op", {}).get("const")
        for item in one_of
    }
    expected = {
        "copy_slide",
        "create_slide_on_layout",
        "rewrite_text",
        "delete_slide",
        "move_slide",
        "set_slide_size",
        "set_slide_layout",
        "set_notes",
        "add_textbox",
        "set_shape_text",
        "add_image",
        "add_shape",
        "add_table",
        "set_slide_background",
        "align_shapes",
        "distribute_shapes",
    }
    assert expected.issubset(op_consts)
