from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pptx.slide import Slide
from pptx.util import Inches

from .models import (
    CopyMode,
    CopySlideOp,
    CreateSlideOnLayoutOp,
    DeleteSlideOp,
    MoveSlideOp,
    Operation,
    OperationPlan,
    RewriteTextOp,
    SetNotesOp,
    SetSlideLayoutOp,
    SetSlideSizeOp,
    parse_plan,
    parse_ops,
)
from .verify import verify_pptx

_SLIDE_LAYOUT_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def _import_copy_ops():
    try:
        from pptx_copy_ops import SlideCopier, SlideSpec  # type: ignore

        return SlideCopier, SlideSpec
    except ModuleNotFoundError:
        pass

    workspace_root = Path(__file__).resolve().parents[3]
    fallback_src = workspace_root / "pptx-copy-ops" / "src"
    if fallback_src.exists() and str(fallback_src) not in sys.path:
        sys.path.insert(0, str(fallback_src))
    try:
        from pptx_copy_ops import SlideCopier, SlideSpec  # type: ignore

        return SlideCopier, SlideSpec
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "pptx-copy-ops is required for copy_slide operations. "
            "Install it or provide sibling path `pptx-copy-ops/src`."
        ) from exc


@dataclass
class ApplyResult:
    output_path: Path
    operations_applied: int
    verify_issues: list[str]


def _iter_text_shapes(slide: Slide):
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            yield shape


def _set_slide_body(slide: Slide, body: str) -> None:
    for shape in _iter_text_shapes(slide):
        if getattr(shape, "is_placeholder", False):
            try:
                placeholder_type = shape.placeholder_format.type
            except Exception:
                placeholder_type = None
            # BODY placeholder
            if placeholder_type == 2:
                shape.text_frame.text = body
                return
    for shape in _iter_text_shapes(slide):
        if shape != slide.shapes.title:
            shape.text_frame.text = body
            return


def _apply_rewrite(op: RewriteTextOp, presentation) -> None:
    if op.slide_index >= len(presentation.slides):
        raise IndexError(
            f"rewrite_text slide_index out of range: {op.slide_index}, total={len(presentation.slides)}"
        )
    slide = presentation.slides[op.slide_index]
    replaced = 0
    for shape in _iter_text_shapes(slide):
        if op.shape_name and shape.name != op.shape_name:
            continue
        text = shape.text_frame.text
        if op.find not in text:
            continue
        count = 1 if op.occurrence == "first" else -1
        shape.text_frame.text = text.replace(op.find, op.replace, count)
        replaced += 1
        if op.occurrence == "first":
            break
    if replaced == 0:
        raise ValueError(
            f"rewrite_text cannot find target text on slide {op.slide_index}: {op.find!r}"
        )


def _apply_create(op: CreateSlideOnLayoutOp, presentation) -> None:
    if op.layout_index >= len(presentation.slide_layouts):
        raise IndexError(
            f"layout_index out of range: {op.layout_index}, total={len(presentation.slide_layouts)}"
        )
    slide = presentation.slides.add_slide(presentation.slide_layouts[op.layout_index])
    if op.title and slide.shapes.title is not None:
        slide.shapes.title.text = op.title
    if op.body:
        _set_slide_body(slide, op.body)


def _apply_delete(op: DeleteSlideOp, presentation) -> None:
    if op.slide_index >= len(presentation.slides):
        raise IndexError(
            f"delete_slide slide_index out of range: {op.slide_index}, total={len(presentation.slides)}"
        )
    r_id = presentation.slides._sldIdLst[op.slide_index].rId
    presentation.part.drop_rel(r_id)
    del presentation.slides._sldIdLst[op.slide_index]


def _apply_move(op: MoveSlideOp, presentation) -> None:
    total = len(presentation.slides)
    if op.from_index >= total:
        raise IndexError(f"move_slide from_index out of range: {op.from_index}, total={total}")
    if op.to_index >= total:
        raise IndexError(f"move_slide to_index out of range: {op.to_index}, total={total}")
    if op.from_index == op.to_index:
        return
    slide_id = presentation.slides._sldIdLst[op.from_index]
    presentation.slides._sldIdLst.remove(slide_id)
    presentation.slides._sldIdLst.insert(op.to_index, slide_id)


def _apply_set_slide_size(op: SetSlideSizeOp, presentation) -> None:
    if op.preset == "16:9":
        width_inches, height_inches = 13.333, 7.5
    elif op.preset == "4:3":
        width_inches, height_inches = 10.0, 7.5
    else:
        width_inches = float(op.width_inches)  # validated non-null by model
        height_inches = float(op.height_inches)
    presentation.slide_width = Inches(width_inches)
    presentation.slide_height = Inches(height_inches)


def _apply_set_slide_layout(op: SetSlideLayoutOp, presentation) -> None:
    if op.slide_index >= len(presentation.slides):
        raise IndexError(
            f"set_slide_layout slide_index out of range: {op.slide_index}, total={len(presentation.slides)}"
        )
    if op.layout_index >= len(presentation.slide_layouts):
        raise IndexError(
            f"set_slide_layout layout_index out of range: {op.layout_index}, total={len(presentation.slide_layouts)}"
        )
    slide = presentation.slides[op.slide_index]
    target_layout = presentation.slide_layouts[op.layout_index]
    for rel in list(slide.part.rels.values()):
        if rel.reltype == _SLIDE_LAYOUT_RELTYPE:
            slide.part.drop_rel(rel.rId)
    slide.part.relate_to(target_layout.part, _SLIDE_LAYOUT_RELTYPE)


def _apply_set_notes(op: SetNotesOp, presentation) -> None:
    if op.slide_index >= len(presentation.slides):
        raise IndexError(
            f"set_notes slide_index out of range: {op.slide_index}, total={len(presentation.slides)}"
        )
    slide = presentation.slides[op.slide_index]
    slide.notes_slide.notes_text_frame.text = op.text


def _to_operations(raw_ops: Iterable[Operation] | list[dict] | dict) -> tuple[list[Operation], OperationPlan | None]:
    if isinstance(raw_ops, dict):
        plan = parse_plan(raw_ops)
        return plan.operations, plan
    raw_list = list(raw_ops)  # type: ignore[arg-type]
    if not raw_list:
        return [], None
    if isinstance(raw_list[0], dict):
        return parse_ops(raw_list), None  # type: ignore[arg-type]
    return raw_list, None  # type: ignore[return-value]


def apply_ops(
    input_pptx: str | Path | None,
    ops: Iterable[Operation] | list[dict] | dict,
    output_pptx: str | Path | None,
    verify: bool = False,
    strict_verify: bool = True,
    template_pptx: str | Path | None = None,
) -> ApplyResult:
    SlideCopier, SlideSpec = _import_copy_ops()

    operations, plan = _to_operations(ops)
    template_path_raw = template_pptx if template_pptx is not None else input_pptx
    if template_path_raw is None and plan and plan.template_pptx:
        template_path_raw = plan.template_pptx
    if template_path_raw is None:
        raise ValueError("template_pptx is required")
    if output_pptx is None:
        raise ValueError("output_pptx is required")

    input_path = Path(template_path_raw).expanduser().resolve()
    output_path = Path(output_pptx).expanduser().resolve()

    copier = SlideCopier(target_template=input_path, clear_existing=False)
    presentation = copier.presentation

    for op in operations:
        if isinstance(op, CopySlideOp):
            mode = op.mode.value if isinstance(op.mode, CopyMode) else str(op.mode)
            source_path = op.source_path
            if source_path is None:
                if not plan:
                    raise ValueError("copy_slide with reuse_library_index requires plan context")
                if op.reuse_library_index is None:
                    raise ValueError("copy_slide missing source_path and reuse_library_index")
                if op.reuse_library_index >= len(plan.reuse_slide_libraries):
                    raise IndexError(
                        f"reuse_library_index out of range: {op.reuse_library_index}, "
                        f"total={len(plan.reuse_slide_libraries)}"
                    )
                source_path = plan.reuse_slide_libraries[op.reuse_library_index]
            copier.copy_slide(
                SlideSpec(source_path=source_path, slide_index=op.source_slide_index),
                mode=mode,
            )
            continue
        if isinstance(op, CreateSlideOnLayoutOp):
            _apply_create(op, presentation)
            continue
        if isinstance(op, RewriteTextOp):
            _apply_rewrite(op, presentation)
            continue
        if isinstance(op, DeleteSlideOp):
            _apply_delete(op, presentation)
            continue
        if isinstance(op, MoveSlideOp):
            _apply_move(op, presentation)
            continue
        if isinstance(op, SetSlideSizeOp):
            _apply_set_slide_size(op, presentation)
            continue
        if isinstance(op, SetSlideLayoutOp):
            _apply_set_slide_layout(op, presentation)
            continue
        if isinstance(op, SetNotesOp):
            _apply_set_notes(op, presentation)
            continue
        raise ValueError(f"Unsupported operation type: {type(op)!r}")

    saved_path = copier.save(output_path)
    issues: list[str] = []
    if verify:
        report = verify_pptx(saved_path)
        issues = report.issues
        if issues and strict_verify:
            raise ValueError("verification failed: " + "; ".join(issues))

    return ApplyResult(
        output_path=saved_path.resolve(),
        operations_applied=len(operations),
        verify_issues=issues,
    )


def generate_pptx(
    template_pptx: str | Path,
    ops: Iterable[Operation] | list[dict] | dict,
    output_pptx: str | Path,
    verify: bool = False,
    strict_verify: bool = True,
) -> ApplyResult:
    """Generate PPTX from a master/layout template and operation list."""
    return apply_ops(
        input_pptx=None,
        template_pptx=template_pptx,
        ops=ops,
        output_pptx=output_pptx,
        verify=verify,
        strict_verify=strict_verify,
    )
