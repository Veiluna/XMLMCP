from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from pydantic import TypeAdapter

from .errors import XMLMCPError, require
from .models import Operation, Target, InspectResult, ApplyResult
from .namespaces import S
from .package import Package
from .workbook import Workbook, public_dict
from .cells import preflight_cells, apply_cell_changes, UNSET
from .images import preflight_images, apply_image_changes
from .calculation import (ExcelCalculator, preflight_calculation, harvest_formula_caches,
                          harvest_chart_caches, harvest_pivot_caches, update_calculation_metadata)
from .integrity import verify_packages, validate_caches
from .xmlbytes import digest
from .runtime import capabilities as runtime_capabilities, require_excel_backend

OPERATIONS = TypeAdapter(list[Operation])


def inspect_template(path: str, targets: list[Target | dict] | None = None) -> InspectResult:
    package = Package(path)
    package.validate()
    book = Workbook(package)
    resolved = []
    for raw in targets or []:
        target = Target.model_validate(raw)
        sheet, rect = book.resolve(target)
        cells = book.cell_map(sheet)
        values = []
        for address in rect.cells():
            cell = cells.get(address)
            info = {"address": address, "value": book.value(cell), **book.style_info(cell)}
            formula = cell.child("f", S) if cell else None
            if formula:
                info["formula"] = formula.text
                info["formula_attributes"] = formula.attrs
            try:
                rpr = book.uniform_rpr(cell)
                info["character_format"] = "uniform_run_properties" if rpr else "cell_style"
            except XMLMCPError as exc:
                if exc.code not in ("MIXED_RICH_TEXT", "PHONETIC_TEXT_UNSUPPORTED"):
                    raise
                info["character_format"] = "mixed_rich_text" if exc.code == "MIXED_RICH_TEXT" else "phonetic_text"
            values.append(info)
        resolved.append({"target": target.model_dump(exclude_none=True), "sheet": sheet["name"], "range": rect.address, "cells": values})
    pivot_info = [public_dict(p) for p in book.pivots()]
    calculation_supported = True
    limitation = None
    try:
        preflight_calculation(book, True)
    except XMLMCPError as exc:
        calculation_supported = False
        limitation = exc.as_dict()
    runtime = runtime_capabilities()
    return InspectResult(path=str(package.path), fingerprint=package.fingerprint,
                         sheets=book.sheets, defined_names=book.names, targets=resolved,
                         images=[public_dict(i) for i in book.images()], charts=[public_dict(c) for c in book.charts()],
                         capabilities={"lossless": "unchanged decompressed part bytes; changed parts use recorded spans",
                                       "write": ["text", "number", "boolean", "ordinary_formula", "matrix", "clear"],
                                       "images": ["PNG", "JPEG"], "style": "copy_existing",
                                       "calculation_supported": calculation_supported, "calculation_limitation": limitation,
                                       "runtime": runtime,
                                       "calculation_available": calculation_supported and runtime["calculation_available"],
                                       "calculation_requires": "Windows Microsoft Excel", "external_refresh": False,
                                       "unsupported_chart_parts": book.unsupported_chart_parts(),
                                       "pivots": pivot_info})


def _publish_noclobber(source: Path, destination: Path):
    if os.name == "nt":
        os.rename(source, destination)  # Windows rename fails atomically if destination exists.
    else:
        os.link(source, destination)
        source.unlink()


def apply_changes(input_path: str, output_path: str, operations: list[Operation | dict], *, calculator=None) -> ApplyResult:
    ops = OPERATIONS.validate_python(operations)
    require(ops, "NO_OPERATIONS", "操作列表不能为空")
    package = Package(input_path)
    package.validate()
    output = Path(output_path)
    require(output.is_absolute(), "ABSOLUTE_PATH_REQUIRED", "输出必须使用绝对路径")
    output = output.resolve()
    require(output.suffix.lower() == ".xlsx", "FILE_TYPE", "输出扩展名必须为 .xlsx")
    require(output != package.path, "INPUT_OVERWRITE", "不得覆盖输入模板")
    report_path = output.with_suffix(".integrity.json")
    require(not output.exists() and not report_path.exists(), "OUTPUT_EXISTS", "输出或报告已存在，不会覆盖")
    require(output.parent.is_dir(), "OUTPUT_DIRECTORY", "输出目录必须已经存在")
    book = Workbook(package)
    changes = preflight_cells(book, ops)
    images = preflight_images(book, ops)
    # Include all known dependent cached objects because Excel recalculates at workbook scope.
    has_formulas = any(cell.child("f", S) for sheet in book.sheets if sheet["kind"] == "worksheet" for cell in book.cell_map(sheet).values())
    has_new_formulas = any(c.formula is not None for cells in changes.values() for c in cells.values())
    calculate = bool(changes) and bool(has_formulas or has_new_formulas or book.charts() or book.pivots() or book.unsupported_chart_parts())
    preflight_calculation(book, calculate)
    if calculator is None and (calculate or images):
        require_excel_backend()
    apply_cell_changes(book, changes)
    # Check newly submitted formulas too before any Excel process starts.
    preflight_calculation(book, calculate)
    calculation = {"calculated": False, "engine": None, "external_refresh": False,
                   "reason": "No dependent formulas/charts/pivots were changed"}
    owned_output = owned_report = False
    try:
        with tempfile.TemporaryDirectory(prefix=".xmlmcp-", dir=output.parent) as temporary:
            directory = Path(temporary)
            if calculate or images:
                calculation, computed_package = (calculator or ExcelCalculator()).run(book, directory, calculate, images)
                if calculate:
                    require(computed_package, "CALCULATION_RESULT", "计算引擎没有返回缓存副本")
                    computed = Workbook(computed_package)
                    for sheet_name, cells in changes.items():
                        expected_cells = book.cell_map(book.sheet(sheet_name))
                        actual_cells = computed.cell_map(computed.sheet(sheet_name))
                        for address, change in cells.items():
                            if change.value is UNSET:
                                continue
                            expected_value, actual_value = book.value(expected_cells.get(address)), computed.value(actual_cells.get(address))
                            same_type = isinstance(expected_value, bool) == isinstance(actual_value, bool)
                            require(same_type and expected_value == actual_value, "INPUT_REINTERPRETED",
                                    "Excel 改变了提交的值；可能是控件关联、查询或派生区域，已取消整批", sheet=sheet_name, cell=address)
                    calculation["formula_caches"] = harvest_formula_caches(book, computed)
                    calculation["pivot_caches"] = harvest_pivot_caches(book, computed)
                    calculation["chart_caches"] = harvest_chart_caches(book, computed, calculation)
                    update_calculation_metadata(book)
                apply_image_changes(book, images, calculation.get("geometry", {}))
            package.validate()
            validate_caches(Workbook(package), calculate)
            require(digest(package.path.read_bytes()) == package.fingerprint, "INPUT_CHANGED", "处理过程中模板发生变化，已取消输出")
            staging = directory / "result.xlsx"
            package.write(staging)
            report = package.manifest()
            report.update({"output_fingerprint": digest(staging.read_bytes()), "output_path": str(output),
                           "operations": [op.model_dump(mode="json", exclude_none=True) for op in ops],
                           "calculation": calculation})
            original = Package(package.path)
            result = Package(staging)
            integrity = verify_packages(original, result, report)
            require(integrity.valid, "INTEGRITY_FAILED", "输出未通过完整性验证", errors=integrity.errors)
            report["verification"] = integrity.model_dump()
            staging_report = directory / "report.json"
            staging_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            require(digest(package.path.read_bytes()) == package.fingerprint, "INPUT_CHANGED", "发布前模板发生变化，已取消输出")
            # Publish the report first; rollback only artifacts successfully created by this call.
            _publish_noclobber(staging_report, report_path)
            owned_report = True
            _publish_noclobber(staging, output)
            owned_output = True
        return ApplyResult(output_path=str(output), report_path=str(report_path), calculation=calculation,
                           integrity=integrity.model_dump(),
                           changes=[{"part": part, "patch_count": sum(len(s["patches"]) for s in package.stages if s["part"] == part)}
                                    for part in integrity.changed_parts] + [{"part": p, "added": True} for p in integrity.added_parts]
                           + [{"part": p, "removed": True} for p in integrity.removed_parts])
    except Exception:
        if owned_output:
            output.unlink(missing_ok=True)
        if owned_report:
            report_path.unlink(missing_ok=True)
        raise

