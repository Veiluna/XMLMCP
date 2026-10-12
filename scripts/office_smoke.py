r"""Create real Office fixtures, apply XMLMCP, then open results without F9/save.

Run from the project: .venv\Scripts\python.exe scripts\office_smoke.py
Only newly created test workbooks and independently owned Office instances are used.
"""
from __future__ import annotations

import argparse
import gc
import json
import sys
import uuid
from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client
import win32process
from PIL import Image

from xmlmcp.engine import apply_changes, inspect_template
from xmlmcp.integrity import verify_integrity
from xmlmcp.package import Package
from xmlmcp.workbook import Workbook


@contextmanager
def office(progid="Excel.Application"):
    pythoncom.CoInitialize()
    previous = set(win32process.EnumProcesses())
    app = win32com.client.DispatchEx(progid)
    _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    if pid in previous:
        app = None
        pythoncom.CoUninitialize()
        raise RuntimeError("Office did not create an independent process")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AskToUpdateLinks = False
            app.EnableEvents = False
            app.AutomationSecurity = 3
        except Exception:
            pass
        yield app
    finally:
        try:
            try:
                app.Quit()
            except Exception:
                pass  # An Office process that already exited has no live COM endpoint.
        finally:
            app = None
            gc.collect()
            pythoncom.CoUninitialize()


def make_fixture(directory: Path, kind: str) -> Path:
    template = directory / f"{kind}-template.xlsx"
    image = directory / "original.png"
    if not image.exists():
        Image.new("RGB", (100, 50), "navy").save(image)
    with office() as app:
        book = app.Workbooks.Add()
        external = None
        try:
            while book.Worksheets.Count > 1:
                book.Worksheets.Item(book.Worksheets.Count).Delete()
            sheet = book.Worksheets.Item(1)
            sheet.Name = "报告"
            sheet.Range("A1:O40").RowHeight = 18
            sheet.Columns("A:O").ColumnWidth = 12
            sheet.Range("A1").Value2 = "Old title"
            sheet.Range("A1").Font.Name = "Arial"
            sheet.Range("A1").Font.Size = 9
            sheet.Range("A2").Value2 = "Tiny English"
            sheet.Range("A2").Font.Name = "Arial"
            sheet.Range("A2").Font.Size = 7
            sheet.Range("B5:C6").Value2 = (("甲", 10), ("乙", 20))
            sheet.Range("D5").Formula = "=SUM(C5:C6)"
            sheet.Range("H1").AddComment("KEEP COMMENT")
            checkbox = sheet.CheckBoxes().Add(10, 75, 100, 20)
            checkbox.Name = "Keep Checkbox"
            checkbox.Caption = "KEEP CHECKBOX"
            checkbox.Value = 1
            text = sheet.Shapes.AddTextbox(1, 150, 75, 150, 20)
            text.Name = "Keep Caption"
            text.TextFrame2.TextRange.Text = "KEEP QR CAPTION"
            pic = sheet.Shapes.AddPicture(str(image), False, True, 400, 100, 100, 50)
            pic.Name = "Keep Picture"
            book.Names.Add(Name="Title", RefersTo="='报告'!$A$1")
            book.Names.Add(Name="Tiny", RefersTo="='报告'!$A$2")
            book.Names.Add(Name="Grid", RefersTo="='报告'!$B$5:$C$6")
            book.Names.Add(Name="Photo", RefersTo="='报告'!$F$5:$H$9")
            book.SaveAs(str(template), FileFormat=51)
            obj = sheet.ChartObjects().Add(10, 380, 360, 240)
            obj.Name = "Data Chart"
            chart = obj.Chart
            chart.ChartType = 51
            chart.HasTitle = True
            chart.ChartTitle.Text = kind + " data"
            if kind == "pivot":
                sheet.Range("B12:C15").Value2 = (("项目", "金额"), ("甲", 1), ("乙", 2), ("丙", 3))
                cache = book.PivotCaches().Create(SourceType=1, SourceData="'报告'!R12C2:R15C3")
                pivot = cache.CreatePivotTable(TableDestination=sheet.Range("J12"), TableName="Summary")
                field = pivot.PivotFields("项目")
                field.Orientation = 1
                field.Position = 1
                pivot.AddDataField(pivot.PivotFields("金额"), "金额合计", -4157)
                pivot.RowAxisLayout(1)
                pivot.RefreshTable()
                chart.SetSourceData(pivot.TableRange1)
                field = pivot = cache = None
            else:
                series = chart.SeriesCollection().NewSeries()
                series.Name = "金额"
                series.XValues = "='报告'!$B$5:$B$6"
                if kind == "dynamic":
                    book.Names.Add(Name="DynamicValues", RefersTo="=OFFSET('报告'!$C$5,0,0,COUNTA('报告'!$B$5:$B$6),1)")
                    series.Values = f"='{book.Name}'!DynamicValues"
                elif kind == "external":
                    external = app.Workbooks.Add()
                    external.Worksheets.Item(1).Name = "Data"
                    external.Worksheets.Item(1).Range("B2:B3").Value2 = ((7,), (9,))
                    external.SaveAs(str(directory / "external-source.xlsx"), FileFormat=51)
                    series.Values = f"='[{external.Name}]Data'!$B$2:$B$3"
                else:
                    series.Values = "='报告'!$C$5:$C$6"
                series = None
            app.CalculateFullRebuild()
            chart.Refresh()
            book.Save()
            chart = obj = pic = checkbox = text = sheet = None
        finally:
            book.Close(False)
            book = None
            if external:
                external.Close(False)
                external = None
    return template


def read_first_open(path: Path, directory: Path, engine: str, expected: tuple, with_images: bool):
    progid = "Excel.Application" if engine == "excel" else "KET.Application"
    with office(progid) as app:
        try:
            app.Calculation = -4135
        except Exception:
            pass
        book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            sheet = book.Worksheets.Item("报告")
            values = tuple(sheet.ChartObjects().Item("Data Chart").Chart.SeriesCollection().Item(1).Values)
            formula = float(sheet.Range("D5").Value2)
            assert formula == 300, (engine, formula)
            assert values == expected, (engine, values, expected)
            assert sheet.Range("A1").Value2 == "New title"
            assert float(sheet.Range("A2").Font.Size) == 7
            assert sheet.CheckBoxes().Item("Keep Checkbox").Value == 1
            assert sheet.Range("H1").Comment.Text() == "KEEP COMMENT"
            assert sheet.Shapes.Item("Keep Caption").TextFrame2.TextRange.Text == "KEEP QR CAPTION"
            expected_images = 2 if with_images else 1
            pictures = [sheet.Shapes.Item(i) for i in range(1, sheet.Shapes.Count+1) if int(sheet.Shapes.Item(i).Type) == 13]
            assert len(pictures) == expected_images, (engine, len(pictures))
            if with_images:
                original_picture = sheet.Shapes.Item("Keep Picture")
                assert abs(float(original_picture.Left)-400) < 0.1 and abs(float(original_picture.Top)-100) < 0.1
                assert abs(float(original_picture.Width)-100) < 0.1 and abs(float(original_picture.Height)-50) < 0.1
                inserted = next(p for p in pictures if str(p.Name).startswith("XMLMCP Image"))
                target = sheet.Range("F5:H9")
                assert abs(float(inserted.Left)-float(target.Left)) < 0.1
                assert abs(float(inserted.Top)-float(target.Top)) < 0.1
                assert abs(float(inserted.Width)-float(target.Width)) < 0.1
                assert abs(float(inserted.Height)-float(target.Height)) < 0.1
                original_picture = inserted = target = None
            pictures = None
            chart = sheet.ChartObjects().Item("Data Chart").Chart
            png = directory / f"{path.stem}-{engine}.png"
            exported = bool(chart.Export(str(png), "PNG", False))
            export_valid = png.is_file() and png.stat().st_size > 1000
            if export_valid:
                with Image.open(png) as rendered:
                    rendered.verify()
            result = {"engine": engine, "version": str(app.Version), "formula": formula,
                      "chart_values": list(values), "first_open": True, "manual_recalculation": False,
                      "saved": False, "font_size": 7, "checkbox_retained": True, "comment_retained": True,
                      "caption_retained": True, "image_count": len(Workbook(Package(path)).images()),
                      "chart_export": str(png) if export_valid else None, "export_return_value": exported}
            chart = sheet = None
            return result
        finally:
            book.Close(False)
            book = None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--kind", choices=["basic", "dynamic", "pivot", "external", "all"], default="all")
    parser.add_argument("--wps", action="store_true")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    directory = root / "test-artifacts" / "office" / uuid.uuid4().hex[:10]
    directory.mkdir(parents=True)
    replacement = directory / "replacement.png"
    Image.new("RGB", (240, 60), "crimson").save(replacement)
    results = []
    kinds = ["basic", "dynamic", "pivot", "external"] if args.kind == "all" else [args.kind]
    for kind in kinds:
        entry = {"kind": kind}
        try:
            template = make_fixture(directory, kind)
            entry["template"] = str(template)
            original = Package(template)
            info = inspect_template(str(template))
            entry["charts"] = info.charts
            ops = [{"op": "write", "target": {"name": "Title"}, "value": "New title"},
                   {"op": "write", "target": {"name": "Tiny"}, "value": "Small English preserved"},
                   {"op": "write", "target": {"name": "Grid"}, "value": [["甲", 100], ["乙", 200]]}]
            if kind == "basic":
                ops += [{"op": "add_image", "target": {"name": "Photo"}, "image_path": str(replacement)},
                        {"op": "replace_image", "image_id": info.images[0]["image_id"], "image_path": str(replacement)}]
            if kind == "pivot":
                ops.append({"op": "write", "target": {"sheet": "报告", "range": "C13:C15"}, "value": [[10], [20], [30]]})
            output = directory / f"{kind}-result.xlsx"
            result = apply_changes(str(template), str(output), ops)
            entry["apply"] = result.model_dump()
            entry["integrity"] = verify_integrity(str(template), str(output), result.report_path).model_dump()
            assert entry["integrity"]["valid"]
            produced = Package(output)
            for name in original.parts:
                if "vml" in name.lower() or "ctrlprops" in name.lower() or "comments" in name.lower() or name == "xl/styles.xml":
                    assert original.parts[name] == produced.parts[name], name
            expected = (7.0, 9.0) if kind == "external" else (10.0, 20.0, 30.0) if kind == "pivot" else (100.0, 200.0)
            entry["excel"] = read_first_open(output, directory, "excel", expected, kind == "basic")
            if args.wps:
                entry["wps"] = read_first_open(output, directory, "wps", expected, kind == "basic")
            entry["passed"] = True
        except Exception as exc:
            import traceback
            entry["passed"] = False
            entry["error"] = str(exc)
            entry["error_details"] = getattr(exc, "details", {})
            entry["error_code"] = getattr(exc, "code", None)
            entry["traceback"] = traceback.format_exc()
        results.append(entry)
        print(json.dumps({"kind": kind, "passed": entry["passed"], "error": entry.get("error")}, ensure_ascii=False), flush=True)
        (directory / "acceptance.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(directory / "acceptance.json"))
    if not all(e["passed"] for e in results):
        sys.exit(1)


if __name__ == "__main__":
    main()

