from __future__ import annotations

import pytest


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
                },
                {
                    "op": "add_textbox",
                    "slide_index": 0,
                    "x_inches": 1,
                    "y_inches": 1,
                    "width_inches": 4,
                    "height_inches": 2,
                    "paragraphs": [
                        {"text": "一级", "list_type": "bullet", "level": 0},
                        {"text": "二级", "list_type": "number", "level": 1}
                    ]
                },
                {
                    "op": "set_shape_text",
                    "slide_index": 0,
                    "shape_name": "content_box",
                    "paragraphs": [
                        {"text": "更新正文", "font_size_pt": 18, "alignment": "center"}
                    ]
                },
                {
                    "op": "add_shape",
                    "slide_index": 0,
                    "shape_type": "rect",
                    "x_inches": 1,
                    "y_inches": 1,
                    "width_inches": 1,
                    "height_inches": 1
                },
                {
                    "op": "add_table",
                    "slide_index": 0,
                    "x_inches": 1,
                    "y_inches": 1,
                    "width_inches": 5,
                    "height_inches": 2,
                    "data": [["A", "B"], ["1", "2"]]
                },
                {
                    "op": "add_image",
                    "slide_index": 0,
                    "image_path": "/tmp/a.png",
                    "x_inches": 1,
                    "y_inches": 1,
                    "width_inches": 2,
                    "height_inches": 1,
                    "fit": "cover"
                },
                {
                    "op": "set_slide_background",
                    "slide_index": 0,
                    "color_hex": "112233"
                },
                {
                    "op": "align_shapes",
                    "slide_index": 0,
                    "shape_names": ["a", "b"],
                    "align": "middle",
                    "reference": "slide"
                },
                {
                    "op": "distribute_shapes",
                    "slide_index": 0,
                    "shape_names": ["a", "b", "c"],
                    "direction": "horizontal"
                },
                {
                    "op": "fill_placeholder",
                    "slide_index": 0,
                    "placeholder_type": "title",
                    "text": "标题"
                },
                {
                    "op": "set_shape_geometry",
                    "slide_index": 0,
                    "shape_name": "a",
                    "x_inches": 1.2,
                    "y_inches": 2.1
                },
                {
                    "op": "set_shape_z_order",
                    "slide_index": 0,
                    "shape_name": "a",
                    "action": "bring_to_front"
                },
                {
                    "op": "add_chart",
                    "slide_index": 0,
                    "chart_type": "column_clustered",
                    "x_inches": 1,
                    "y_inches": 1,
                    "width_inches": 4,
                    "height_inches": 2,
                    "categories": ["Q1", "Q2"],
                    "series": [
                        {"name": "收入", "values": [1, 2]}
                    ]
                },
                {
                    "op": "set_table_cell",
                    "slide_index": 0,
                    "table_name": "tbl",
                    "row": 0,
                    "col": 0,
                    "text": "A1"
                },
                {
                    "op": "merge_table_cells",
                    "slide_index": 0,
                    "table_name": "tbl",
                    "start_row": 0,
                    "start_col": 0,
                    "end_row": 0,
                    "end_col": 1
                },
                {
                    "op": "set_shape_hyperlink",
                    "slide_index": 0,
                    "shape_name": "btn",
                    "url": "https://example.com"
                },
                {
                    "op": "replace_image",
                    "slide_index": 0,
                    "shape_name": "logo",
                    "image_path": "/tmp/b.png"
                },
                {
                    "op": "update_chart_data",
                    "slide_index": 0,
                    "chart_name": "chart1",
                    "categories": ["Q1", "Q2"],
                    "series": [
                        {"name": "s1", "values": [1, 2]}
                    ]
                },
                {
                    "op": "set_table_style",
                    "slide_index": 0,
                    "table_name": "tbl",
                    "header_bold": True
                },
                {
                    "op": "set_table_row_col_size",
                    "slide_index": 0,
                    "table_name": "tbl",
                    "row_index": 0,
                    "row_height_inches": 0.6
                },
                {
                    "op": "set_text_hyperlink",
                    "slide_index": 0,
                    "shape_name": "tb",
                    "url": "https://example.com"
                }
            ],
        }
    )

    ops = parse_ops(plan.model_dump())
    assert plan.template_pptx == "/tmp/template_master.pptx"
    assert plan.reuse_slide_libraries == ["/tmp/reuse_library_a.pptx"]
    assert len(ops) == 27
    assert ops[0].op == "copy_slide"
    assert ops[0].reuse_library_index == 0
    assert ops[1].op == "delete_slide"
    assert ops[2].op == "create_slide_on_layout"
    assert ops[3].op == "move_slide"
    assert ops[4].op == "set_slide_size"
    assert ops[5].op == "set_slide_layout"
    assert ops[6].op == "set_notes"
    assert ops[7].op == "add_textbox"
    assert ops[8].op == "set_shape_text"
    assert ops[9].op == "add_shape"
    assert ops[10].op == "add_table"
    assert ops[11].op == "add_image"
    assert ops[12].op == "set_slide_background"
    assert ops[13].op == "align_shapes"
    assert ops[14].op == "distribute_shapes"
    assert ops[15].op == "fill_placeholder"
    assert ops[16].op == "set_shape_geometry"
    assert ops[17].op == "set_shape_z_order"
    assert ops[18].op == "add_chart"
    assert ops[19].op == "set_table_cell"
    assert ops[20].op == "merge_table_cells"
    assert ops[21].op == "set_shape_hyperlink"
    assert ops[22].op == "replace_image"
    assert ops[23].op == "update_chart_data"
    assert ops[24].op == "set_table_style"
    assert ops[25].op == "set_table_row_col_size"
    assert ops[26].op == "set_text_hyperlink"


def test_set_shape_text_requires_shape_target() -> None:
    from pptx_ooxml_engine.models import parse_plan

    with pytest.raises(ValueError):
        parse_plan(
            {
                "operations": [
                    {
                        "op": "set_shape_text",
                        "slide_index": 0,
                        "text": "missing target",
                    }
                ]
            }
        )


def test_add_chart_series_values_must_match_categories() -> None:
    from pptx_ooxml_engine.models import parse_plan

    with pytest.raises(ValueError):
        parse_plan(
            {
                "operations": [
                    {
                        "op": "add_chart",
                        "slide_index": 0,
                        "chart_type": "line",
                        "x_inches": 1,
                        "y_inches": 1,
                        "width_inches": 3,
                        "height_inches": 2,
                        "categories": ["Q1", "Q2", "Q3"],
                        "series": [
                            {"name": "s1", "values": [1, 2]}
                        ],
                    }
                ]
            }
        )


def test_replace_image_requires_target() -> None:
    from pptx_ooxml_engine.models import parse_plan

    with pytest.raises(ValueError):
        parse_plan(
            {
                "operations": [
                    {
                        "op": "replace_image",
                        "slide_index": 0,
                        "image_path": "/tmp/a.png",
                    }
                ]
            }
        )


def test_update_chart_data_series_values_must_match_categories() -> None:
    from pptx_ooxml_engine.models import parse_plan

    with pytest.raises(ValueError):
        parse_plan(
            {
                "operations": [
                    {
                        "op": "update_chart_data",
                        "slide_index": 0,
                        "chart_name": "c1",
                        "categories": ["Q1", "Q2", "Q3"],
                        "series": [
                            {"name": "s1", "values": [1, 2]}
                        ],
                    }
                ]
            }
        )
