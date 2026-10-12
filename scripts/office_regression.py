"""Exercise formula types, fixed arrays and transaction failures with native Office."""
from __future__ import annotations

import argparse
import json
import uuid
from pathlib import Path

from office_smoke import office
from xmlmcp.engine import apply_changes
from xmlmcp.errors import XMLMCPError
from xmlmcp.integrity import verify_integrity
from xmlmcp.package import Package
from xmlmcp.workbook import Workbook


def check_values(path: Path, expected: dict, wps: bool) -> list[dict]:
    checks = []
    for engine in (["Excel.Application", "KET.Application"] if wps else ["Excel.Application"]):
        with office(engine) as app:
            version = str(app.Version)
            try:
                app.Calculation = -4135
            except Exception:
                pass
            book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
            try:
                sheet = book.Worksheets.Item("报告")
                actual = {address: sheet.Range(address).Value2 for address in expected}
                assert actual == expected, (engine, actual, expected)
                checks.append({"engine": engine, "version": version, "values": actual,
                               "manual_recalculation": False, "saved": False})
                sheet = None
            finally:
                book.Close(False)
                book = None
    return checks


def assert_rejected(template: Path, output: Path, operations: list, expected_codes: set[str]) -> dict:
    before = Package(template).fingerprint
    try:
        apply_changes(str(template), str(output), operations)
    except XMLMCPError as exc:
        assert exc.code in expected_codes, (exc.code, exc.details)
        assert not output.exists() and not output.with_suffix(".integrity.json").exists()
        assert not list(output.parent.glob(".xmlmcp-*"))
        assert Package(template).fingerprint == before
        return {"code": exc.code, "details": exc.details, "no_output": True, "input_unchanged": True}
    raise AssertionError("The batch was expected to fail")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("fixtures", type=Path, help="A successful office_smoke output directory")
    parser.add_argument("--wps", action="store_true")
    args = parser.parse_args()
    fixtures = args.fixtures.resolve()
    directory = fixtures.parent / ("regression-" + uuid.uuid4().hex[:10])
    directory.mkdir()
    template = fixtures / "basic-template.xlsx"
    results = {}
    output = directory / "formulas.xlsx"
    operations = [
        {"op": "write", "target": {"name": "Grid", "scope": "workbook"}, "value": [["甲", 100], ["乙", 200]]},
        {"op": "formula", "target": {"sheet": "报告", "range": "D5"}, "formula": "=SUM(C5:C6)*2"},
        {"op": "formula", "target": {"sheet": "报告", "range": "E5"}, "formula": "=MAX(C5:C6)"},
        {"op": "formula", "target": {"sheet": "报告", "range": "E8"}, "formula": "=C5<C6"},
        {"op": "formula", "target": {"sheet": "报告", "range": "E9"}, "formula": '="金额"&C5'},
        {"op": "write", "target": {"sheet": "报告", "range": "A4"}, "value": "=SUM(1,2)"},
        {"op": "write", "target": {"sheet": "报告", "range": "A5"}, "value": " literal _x0041_ & < > "},
        {"op": "write", "target": {"sheet": "报告", "range": "A6"}, "value": "甲\r\n乙\t末"},
    ]
    applied = apply_changes(str(template), str(output), operations)
    expected = {"D5": 600, "E5": 200, "E8": True, "E9": "金额100", "A4": "=SUM(1,2)",
                "A5": " literal _x0041_ & < > ", "A6": "甲\r\n乙\t末"}
    raw_book = Workbook(Package(output))
    raw_cells = raw_book.cell_map(raw_book.sheet("报告"))
    assert {a: raw_book.value(raw_cells[a]) for a in expected} == expected
    results["formula_types_and_new_cells"] = {
        "integrity": verify_integrity(str(template), str(output), applied.report_path).model_dump(),
        "first_open": check_values(output, expected, args.wps),
    }
    results["formula_error_rolls_back"] = assert_rejected(
        template, directory / "formula-error.xlsx",
        [{"op": "formula", "target": {"sheet": "报告", "range": "D5"}, "formula": "=1/0"}], {"FORMULA_ERROR"})
    results["spill_rejected"] = assert_rejected(
        template, directory / "spill.xlsx",
        [{"op": "formula", "target": {"sheet": "报告", "range": "E5"}, "formula": "=SEQUENCE(2)"}],
        {"FORMULA_MAPPING", "FORMULA_ERROR"})
    array_template = directory / "array-template.xlsx"
    with office() as app:
        book = app.Workbooks.Open(str(template), UpdateLinks=0)
        try:
            sheet = book.Worksheets.Item("报告")
            sheet.Range("E5:E6").FormulaArray = "=C5:C6*2"
            app.CalculateFullRebuild()
            book.SaveAs(str(array_template), FileFormat=51)
            sheet = None
        finally:
            book.Close(False)
            book = None
    output = directory / "array-result.xlsx"
    applied = apply_changes(str(array_template), str(output), [
        {"op": "write", "target": {"name": "Grid", "scope": "workbook"}, "value": [["甲", 100], ["乙", 200]]}])
    results["fixed_array_cache"] = {
        "integrity": verify_integrity(str(array_template), str(output), applied.report_path).model_dump(),
        "first_open": check_values(output, {"D5": 300, "E5": 200, "E6": 400}, args.wps),
    }
    results["pivot_layout_change_rolls_back"] = assert_rejected(
        fixtures / "pivot-template.xlsx", directory / "pivot-layout.xlsx",
        [{"op": "write", "target": {"sheet": "报告", "range": "B13:B15"}, "value": [["甲"], ["甲"], ["甲"]]}],
        {"PIVOT_LAYOUT_CHANGED"})
    results["pivot_output_write_rolls_back"] = assert_rejected(
        fixtures / "pivot-template.xlsx", directory / "pivot-output.xlsx",
        [{"op": "write", "target": {"sheet": "报告", "range": "K13"}, "value": 999}], {"INPUT_REINTERPRETED"})
    path = directory / "acceptance.json"
    path.write_text(json.dumps({"passed": True, "checks": results}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
