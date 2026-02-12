from __future__ import annotations

from enum import Enum
from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field, model_validator


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


Operation = Annotated[
    Union[CopySlideOp, CreateSlideOnLayoutOp, RewriteTextOp],
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
