from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches


def _build_source_pptx(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    if slide.shapes.title is not None:
        slide.shapes.title.text = "From Source"
    prs.save(str(path))


def _build_target_pptx(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    if slide.shapes.title is not None:
        slide.shapes.title.text = "Original Title"
    prs.save(str(path))


def _build_three_slide_pptx(path: Path) -> None:
    prs = Presentation()
    titles = ["Slide A", "Slide B", "Slide C"]
    for title in titles:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        if slide.shapes.title is not None:
            slide.shapes.title.text = title
    prs.save(str(path))


def _slide_texts(slide) -> list[str]:
    out: list[str] = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text_frame.text.strip()
            if text:
                out.append(text)
    return out


def test_apply_ops_builds_expected_output(tmp_path: Path) -> None:
    from pptx_ooxml_engine.engine import apply_ops

    source = tmp_path / "source.pptx"
    target = tmp_path / "target.pptx"
    output = tmp_path / "output.pptx"
    _build_source_pptx(source)
    _build_target_pptx(target)

    result = apply_ops(
        input_pptx=target,
        ops=[
            {"op": "rewrite_text", "slide_index": 0, "find": "Original", "replace": "Rewritten"},
            {"op": "copy_slide", "source_path": str(source), "source_slide_index": 0, "mode": "shape"},
            {"op": "create_slide_on_layout", "layout_index": 0, "title": "Created Slide", "body": "Body"},
        ],
        output_pptx=output,
        verify=True,
    )

    assert result.output_path == output.resolve()
    prs = Presentation(str(output))
    assert len(prs.slides) == 3

    s1_texts = _slide_texts(prs.slides[0])
    s2_texts = _slide_texts(prs.slides[1])
    s3_texts = _slide_texts(prs.slides[2])
    assert any("Rewritten Title" == t for t in s1_texts)
    assert any("From Source" == t for t in s2_texts)
    assert any("Created Slide" == t for t in s3_texts)


def test_apply_ops_can_take_template_from_plan_dict(tmp_path: Path) -> None:
    from pptx_ooxml_engine.engine import apply_ops

    template = tmp_path / "template.pptx"
    output = tmp_path / "output_from_plan.pptx"
    _build_target_pptx(template)

    result = apply_ops(
        input_pptx=None,
        ops={
            "template_pptx": str(template),
            "reuse_slide_libraries": ["/tmp/reuse1.pptx"],
            "operations": [
                {"op": "rewrite_text", "slide_index": 0, "find": "Original", "replace": "Plan"},
                {"op": "create_slide_on_layout", "layout_index": 0, "title": "Added By Plan"},
            ],
        },
        output_pptx=output,
        verify=True,
    )

    assert result.output_path == output.resolve()
    prs = Presentation(str(output))
    assert len(prs.slides) == 2
    assert any("Plan Title" == t for t in _slide_texts(prs.slides[0]))


def test_apply_ops_copy_from_reuse_library_index(tmp_path: Path) -> None:
    from pptx_ooxml_engine.engine import apply_ops

    template = tmp_path / "template.pptx"
    reuse_lib = tmp_path / "reuse_lib.pptx"
    output = tmp_path / "output_from_reuse_index.pptx"
    _build_target_pptx(template)
    _build_source_pptx(reuse_lib)

    result = apply_ops(
        input_pptx=None,
        ops={
            "template_pptx": str(template),
            "reuse_slide_libraries": [str(reuse_lib)],
            "operations": [
                {"op": "copy_slide", "reuse_library_index": 0, "source_slide_index": 0, "mode": "shape"},
            ],
        },
        output_pptx=output,
        verify=True,
    )

    assert result.output_path == output.resolve()
    prs = Presentation(str(output))
    assert len(prs.slides) == 2
    assert any("From Source" == t for t in _slide_texts(prs.slides[1]))


def test_apply_ops_delete_slide_then_create(tmp_path: Path) -> None:
    from pptx_ooxml_engine.engine import apply_ops

    template = tmp_path / "template.pptx"
    output = tmp_path / "output_delete_then_create.pptx"
    _build_target_pptx(template)  # contains one initial slide

    result = apply_ops(
        input_pptx=None,
        ops={
            "template_pptx": str(template),
            "operations": [
                {"op": "delete_slide", "slide_index": 0},
                {"op": "create_slide_on_layout", "layout_index": 0, "title": "Only Slide"},
            ],
        },
        output_pptx=output,
        verify=True,
    )

    assert result.output_path == output.resolve()
    prs = Presentation(str(output))
    assert len(prs.slides) == 1
    assert any("Only Slide" == t for t in _slide_texts(prs.slides[0]))


def test_apply_ops_structure_management_ops(tmp_path: Path) -> None:
    from pptx_ooxml_engine.engine import apply_ops

    template = tmp_path / "template_three_slides.pptx"
    output = tmp_path / "output_structure_ops.pptx"
    _build_three_slide_pptx(template)

    result = apply_ops(
        input_pptx=None,
        ops={
            "template_pptx": str(template),
            "operations": [
                {"op": "move_slide", "from_index": 0, "to_index": 2},
                {"op": "set_slide_layout", "slide_index": 0, "layout_index": 1},
                {"op": "set_notes", "slide_index": 0, "text": "Speaker Notes"},
                {"op": "set_slide_size", "preset": "4:3"},
            ],
        },
        output_pptx=output,
        verify=True,
    )

    assert result.output_path == output.resolve()
    prs = Presentation(str(output))
    assert len(prs.slides) == 3
    # move_slide: A B C -> B C A
    assert any("Slide B" == t for t in _slide_texts(prs.slides[0]))
    assert any("Slide C" == t for t in _slide_texts(prs.slides[1]))
    assert any("Slide A" == t for t in _slide_texts(prs.slides[2]))

    # set_slide_size preset 4:3
    assert prs.slide_width == Inches(10)
    assert prs.slide_height == Inches(7.5)

    # set_slide_layout: slide0 layout should be layout 1
    assert str(prs.slides[0].slide_layout.part.partname) == str(prs.slide_layouts[1].part.partname)

    # set_notes
    assert "Speaker Notes" in prs.slides[0].notes_slide.notes_text_frame.text
