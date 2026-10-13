"""Explicit, reusable mappings. Template annotations are data, never instructions."""
from typing import Literal
from pydantic import Field
from .models import Model, Target, Scalar


class OrderLayout(Model):
    fields: dict[str, Target]
    sample_rows: dict[str, Target]
    sample_numbers: Target | None = None
    example_marker: str = "例子"
    checkbox_sheet: str | None = None


class TableSlot(Model):
    target: Target
    header_style: Target | None = None
    data_style: Target | None = None


class CheckboxBinding(Model):
    source_label: str
    target_label: str
    target_sheet: str


class TextRunBinding(Model):
    field: str
    target: Target
    run_index: int = Field(ge=0)
    prefix: str = ""


class ReportProfile(Model):
    schema_version: Literal[1] = 1
    profile_id: str
    data_adapter: Literal["electrode-resistance-csv-v1"] = "electrode-resistance-csv-v1"
    template_fingerprint: str = Field(pattern=r"^[0-9a-f]{64}$")
    order: OrderLayout
    fields: dict[str, list[Target]]
    sample_rows: dict[str, Target] = Field(default_factory=dict)
    tables: dict[str, TableSlot] = Field(default_factory=dict)
    figures: dict[str, Target] = Field(default_factory=dict)
    checkboxes: list[CheckboxBinding] = Field(default_factory=list)
    text_runs: list[TextRunBinding] = Field(default_factory=list)
    defaults: dict[str, Scalar] = Field(default_factory=dict)
    # Physical character-width calibration is part of the reviewed template mapping.
    maximum_digit_width_px: float = Field(default=7, ge=1, le=30)


class ReportInspection(Model):
    template_path: str
    template_fingerprint: str
    fields: dict
    samples: list[dict]
    raw_files: list[dict]
    statistics: list[dict]
    missing_fields: list[str]
    warnings: list[dict]
    suggested_profile: dict | None
    mapping_status: str
    available_targets: list[dict]
    field_catalog: dict[str, str]


class ReportResult(Model):
    output_directory: str
    output_path: str
    integrity_report_path: str
    data_report_path: str
    image_paths: list[str]
    missing_fields: list[str]
    warnings: list[dict]
    integrity: dict
    computation: dict
