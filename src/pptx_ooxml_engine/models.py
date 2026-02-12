from __future__ import annotations

from enum import Enum
from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field, model_validator

HEX_COLOR_PATTERN = r"^#?[0-9A-Fa-f]{6}$"


class CopyMode(str, Enum):
    SHAPE = "shape"
    PART = "part"


class CopySlideOp(BaseModel):
    op: Literal["copy_slide"]
    source_path: str | None = None
    reuse_library_index: int | None = Field(default=None, ge=0)
    source_slide_index: int = Field(ge=0)
    mode: CopyMode = CopyMode.PART

    @model_validator(mode="after")
    def _check_source(self) -> "CopySlideOp":
        if self.source_path is None and self.reuse_library_index is None:
            raise ValueError("copy_slide requires source_path or reuse_library_index")
        return self


class CreateSlideOnLayoutOp(BaseModel):
    op: Literal["create_slide_on_layout"]
    layout_index: int = Field(default=0, ge=0)
    title: str | None = None
    body: str | None = None


class RewriteTextOp(BaseModel):
    op: Literal["rewrite_text"]
    slide_index: int = Field(ge=0)
    find: str
    replace: str
    shape_name: str | None = None
    occurrence: Literal["first", "all"] = "all"


class DeleteSlideOp(BaseModel):
    op: Literal["delete_slide"]
    slide_index: int = Field(ge=0)


class MoveSlideOp(BaseModel):
    op: Literal["move_slide"]
    from_index: int = Field(ge=0)
    to_index: int = Field(ge=0)


class SetSlideSizeOp(BaseModel):
    op: Literal["set_slide_size"]
    preset: Literal["16:9", "4:3"] | None = None
    width_inches: float | None = Field(default=None, gt=0)
    height_inches: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def _check_size_input(self) -> "SetSlideSizeOp":
        if self.preset is not None and (self.width_inches is not None or self.height_inches is not None):
            raise ValueError("set_slide_size: do not provide width/height when preset is used")
        if self.preset is None and (self.width_inches is None or self.height_inches is None):
            raise ValueError("set_slide_size requires preset or both width_inches and height_inches")
        return self


class SetSlideLayoutOp(BaseModel):
    op: Literal["set_slide_layout"]
    slide_index: int = Field(ge=0)
    layout_index: int = Field(ge=0)


class SetNotesOp(BaseModel):
    op: Literal["set_notes"]
    slide_index: int = Field(ge=0)
    text: str


class ParagraphSpec(BaseModel):
    text: str
    level: int = Field(default=0, ge=0, le=8)
    list_type: Literal["none", "bullet", "number"] = "none"
    font_size_pt: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    color_hex: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    alignment: Literal["left", "center", "right", "justify"] | None = None
    line_spacing: float | None = Field(default=None, gt=0)
    space_before_pt: float | None = Field(default=None, ge=0)
    space_after_pt: float | None = Field(default=None, ge=0)


class AddTextBoxOp(BaseModel):
    op: Literal["add_textbox"]
    slide_index: int = Field(ge=0)
    x_inches: float = Field(ge=0)
    y_inches: float = Field(ge=0)
    width_inches: float = Field(gt=0)
    height_inches: float = Field(gt=0)
    name: str | None = None
    text: str | None = None
    paragraphs: list[ParagraphSpec] = Field(default_factory=list)
    vertical_anchor: Literal["top", "middle", "bottom"] | None = None
    word_wrap: bool | None = None

    @model_validator(mode="after")
    def _check_text_input(self) -> "AddTextBoxOp":
        if self.text is None and not self.paragraphs:
            raise ValueError("add_textbox requires text or paragraphs")
        return self


class SetShapeTextOp(BaseModel):
    op: Literal["set_shape_text"]
    slide_index: int = Field(ge=0)
    shape_name: str | None = None
    shape_index: int | None = Field(default=None, ge=0)
    text: str | None = None
    paragraphs: list[ParagraphSpec] = Field(default_factory=list)
    vertical_anchor: Literal["top", "middle", "bottom"] | None = None
    word_wrap: bool | None = None

    @model_validator(mode="after")
    def _check_shape_and_text(self) -> "SetShapeTextOp":
        if self.shape_name is None and self.shape_index is None:
            raise ValueError("set_shape_text requires shape_name or shape_index")
        if self.text is None and not self.paragraphs:
            raise ValueError("set_shape_text requires text or paragraphs")
        return self


class AddImageOp(BaseModel):
    op: Literal["add_image"]
    slide_index: int = Field(ge=0)
    image_path: str = Field(min_length=1)
    x_inches: float = Field(ge=0)
    y_inches: float = Field(ge=0)
    width_inches: float = Field(gt=0)
    height_inches: float = Field(gt=0)
    fit: Literal["stretch", "contain", "cover"] = "stretch"
    name: str | None = None


class AddShapeOp(BaseModel):
    op: Literal["add_shape"]
    slide_index: int = Field(ge=0)
    shape_type: Literal["rect", "round_rect", "ellipse", "right_arrow", "line"]
    x_inches: float = Field(ge=0)
    y_inches: float = Field(ge=0)
    width_inches: float = Field(gt=0)
    height_inches: float = Field(gt=0)
    name: str | None = None
    text: str | None = None
    fill_color_hex: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    line_color_hex: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    line_width_pt: float | None = Field(default=None, gt=0)
    text_color_hex: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    font_size_pt: float | None = Field(default=None, gt=0)


class AddTableOp(BaseModel):
    op: Literal["add_table"]
    slide_index: int = Field(ge=0)
    x_inches: float = Field(ge=0)
    y_inches: float = Field(ge=0)
    width_inches: float = Field(gt=0)
    height_inches: float = Field(gt=0)
    data: list[list[str]]
    header: bool = False
    name: str | None = None
    font_size_pt: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def _check_data(self) -> "AddTableOp":
        if not self.data or not self.data[0]:
            raise ValueError("add_table requires non-empty 2D data")
        return self


class SetSlideBackgroundOp(BaseModel):
    op: Literal["set_slide_background"]
    slide_index: int = Field(ge=0)
    color_hex: str = Field(pattern=HEX_COLOR_PATTERN)


class AlignShapesOp(BaseModel):
    op: Literal["align_shapes"]
    slide_index: int = Field(ge=0)
    shape_names: list[str] = Field(min_length=2)
    align: Literal["left", "center", "right", "top", "middle", "bottom"]
    reference: Literal["first", "slide"] = "first"


class DistributeShapesOp(BaseModel):
    op: Literal["distribute_shapes"]
    slide_index: int = Field(ge=0)
    shape_names: list[str] = Field(min_length=3)
    direction: Literal["horizontal", "vertical"]


Operation = Annotated[
    Union[
        CopySlideOp,
        CreateSlideOnLayoutOp,
        RewriteTextOp,
        DeleteSlideOp,
        MoveSlideOp,
        SetSlideSizeOp,
        SetSlideLayoutOp,
        SetNotesOp,
        AddTextBoxOp,
        SetShapeTextOp,
        AddImageOp,
        AddShapeOp,
        AddTableOp,
        SetSlideBackgroundOp,
        AlignShapesOp,
        DistributeShapesOp,
    ],
    Field(discriminator="op"),
]


class OperationPlan(BaseModel):
    template_pptx: str | None = None
    reuse_slide_libraries: list[str] = Field(default_factory=list)
    operations: list[Operation]


def parse_ops(raw_ops: list[dict] | dict) -> list[Operation]:
    plan = parse_plan(raw_ops)
    return plan.operations


def parse_plan(raw: list[dict] | dict) -> OperationPlan:
    if isinstance(raw, dict):
        return OperationPlan.model_validate(raw)
    plan = OperationPlan.model_validate({"operations": raw})
    return plan
