from __future__ import annotations

from typing import Annotated, Literal
from pydantic import BaseModel, ConfigDict, Field, StrictBool, StrictInt, StrictStr, FiniteFloat, model_validator

Scalar = StrictStr | StrictBool | StrictInt | FiniteFloat | None


class Model(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Target(Model):
    name: str | None = None
    scope: str | None = None
    sheet: str | None = None
    range: str | None = None

    @model_validator(mode="after")
    def kind(self):
        if self.name is not None:
            if self.sheet is not None or self.range is not None:
                raise ValueError("命名区域目标不得同时指定 sheet/range；局部名称使用 scope")
        elif self.sheet is None or self.range is None or self.scope is not None:
            raise ValueError("目标需指定 name（可选 scope），或 sheet 和 range")
        return self


class WriteValue(Model):
    op: Literal["write"]
    target: Target
    value: Scalar | list[list[Scalar]]


class WriteFormula(Model):
    op: Literal["formula"]
    target: Target
    formula: str


class ClearValue(Model):
    op: Literal["clear"]
    target: Target


class CopyStyle(Model):
    op: Literal["copy_style"]
    source: Target
    target: Target


class AddImage(Model):
    op: Literal["add_image"]
    target: Target
    image_path: str


class ReplaceImage(Model):
    op: Literal["replace_image"]
    image_id: str
    image_path: str


Operation = Annotated[WriteValue | WriteFormula | ClearValue | CopyStyle | AddImage | ReplaceImage,
                      Field(discriminator="op")]


class InspectResult(Model):
    path: str
    fingerprint: str
    sheets: list[dict]
    defined_names: list[dict]
    targets: list[dict]
    images: list[dict]
    charts: list[dict]
    capabilities: dict


class ApplyResult(Model):
    output_path: str
    report_path: str
    calculation: dict
    integrity: dict
    changes: list[dict]


class IntegrityResult(Model):
    valid: bool
    input_fingerprint: str
    output_fingerprint: str
    unchanged_parts: int
    changed_parts: list[str]
    added_parts: list[str]
    removed_parts: list[str]
    checks: dict
    errors: list[str]

