from __future__ import annotations


def test_load_ops_schema() -> None:
    from pptx_ooxml_engine.schema import load_ops_schema

    schema = load_ops_schema()
    assert schema["type"] == "object"
    assert "operations" in schema["properties"]
    assert "template_pptx" in schema["properties"]
    assert "reuse_slide_libraries" in schema["properties"]
