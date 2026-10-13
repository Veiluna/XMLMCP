from __future__ import annotations
import asyncio
import csv
import io
import json
import math
import sys
import zipfile
import warnings
from pathlib import Path
from types import SimpleNamespace

import pytest
from mcp.client import Client
from mcp.client.stdio import StdioServerParameters
from conftest import write_template
from xmlmcp import engine, runtime
from xmlmcp.errors import XMLMCPError
from xmlmcp.package import Package
from xmlmcp.workbook import Workbook
from xmlmcp.report_data import parse_csv, read_raw_zip, summarize
from xmlmcp.report_richtext import apply_text_run_values
from xmlmcp.reports import generate_report, inspect_report_inputs
from xmlmcp.publication import publish_bundle
from xmlmcp.integrity import verify_integrity
from xmlmcp.namespaces import S


def instrument_csv(kind, values=(1, 3), nul=False, encoding="utf-8"):
    rows = [["单点实验报告" if kind == "single_point" else "变压实验报告"],
            ["温度(℃):23", "湿度(%RH):50", "实验面积(mm2):153.94"],
            ["开始时间:2026-09-10 10:00:00", "结束时间:2026-09-10 10:01:00"],
            ["实验相关参数"],
            ["采样数目:2", "保压时间(s):25", "压强(MPa):5"] if kind == "single_point" else
            ["区间[1]", "压强下限(MPa):5.00(MPa)", "压强上限(MPa):10.00(MPa)", "压强步进(MPa):5.00(MPa)", "保压时间:15(s)"],
            ["序号", "压力(kg)", "压强(MPa)", "极片总厚度(μm)", "电阻(Ω)", "电导率(S/cm)", "电阻率(Ω*cm)"]]
    rows += [[i + 1, 78 * (i + 1), 5 if kind == "single_point" else 5 * (i + 1), 70, .01, 1 / v, v] for i, v in enumerate(values)]
    if kind == "single_point":
        rows += [["平均值", 78, 5, 70, .01, 1, sum(values) / len(values)], ["COV%", 0, 0, 0, 0, 0, 99]]
    stream = io.StringIO()
    csv.writer(stream).writerows(rows)
    value = stream.getvalue()
    if nul:
        value = value.replace("\r\n", "\r\n\x00")
    return value.encode(encoding)


@pytest.fixture
def report_inputs(tmp_path):
    template = write_template(tmp_path / "report-template.xlsx", image=True)
    order_template = write_template(tmp_path / "order-original.xlsx")
    order = tmp_path / "order.xlsx"
    engine.apply_changes(str(order_template), str(order), [{"op": "write", "target": {"sheet": "报告", "range": "B5:C5"}, "value": [["Sample-A", "Sample-B"]]}])
    raw = tmp_path / "raw.zip"
    with zipfile.ZipFile(raw, "w") as z:
        for sample, values in (("Sample-A", (1, 3)), ("Sample-B", (5, 7))):
            z.writestr("单点/" + sample + ".csv", instrument_csv("single_point", values, sample == "Sample-B", "gb18030"))
            z.writestr("变压/" + sample + ".csv", instrument_csv("pressure", tuple(reversed(values))))
    profile = {"profile_id": "two-sample-example", "template_fingerprint": Package(template).fingerprint,
               "order": {"fields": {"client_company": {"sheet": "报告", "range": "A1"}},
                         "sample_rows": {"sample_name": {"sheet": "报告", "range": "B5:C5"}}},
               "fields": {"client_company": [{"sheet": "报告", "range": "A1"}], "test_date": [{"sheet": "报告", "range": "D2"}],
                          "operator": [{"sheet": "报告", "range": "D3"}]},
               "sample_rows": {"sample_name": {"sheet": "报告", "range": "E2:G2"}},
               "tables": {"single_summary": {"target": {"sheet": "报告", "range": "A10:H13"}},
                          "pressure_resistivity": {"target": {"sheet": "报告", "range": "A15:C18"}}},
               "figures": {"single_point": {"sheet": "报告", "range": "E15:H20"}}}
    mapping = tmp_path / "profile.json"
    mapping.write_text(json.dumps(profile), encoding="utf-8")
    return dict(template_path=str(template), order_path=str(order), raw_zip_path=str(raw), profile_path=str(mapping)), profile


def test_instrument_encoding_nul_and_independent_statistics():
    item = parse_csv(instrument_csv("single_point", nul=True, encoding="gb18030"), "单点/Sample-A.csv")
    result = summarize(item)
    assert result["mean"]["resistivity_ohm_cm"] == 2
    assert result["sample_std"]["resistivity_ohm_cm"] == pytest.approx(math.sqrt(2))
    assert result["sample_cov_pct"]["resistivity_ohm_cm"] == pytest.approx(100 / math.sqrt(2))
    assert item["removed_line_prefix_nuls"] > 0 and item["encoding"] == "gb18030"
    assert item["instrument_summary"]["COV%"]["resistivity_ohm_cm"] == 99


@pytest.mark.parametrize("kind", ["single_point", "pressure"])
def test_chinese_plots_render_without_system_fonts(tmp_path, monkeypatch, kind):
    from PIL import Image
    from xmlmcp import report_plots
    monkeypatch.setattr(report_plots, "SYSTEM_FONTS", ())
    sample = {"sample_name": "Sample-A"}
    item = parse_csv(instrument_csv(kind), "Sample-A.csv")
    image = tmp_path / (kind + ".png")
    with warnings.catch_warnings():
        warnings.filterwarnings("error", message="Glyph .* missing from font")
        report_plots.make_plot(kind, [sample], {("Sample-A", kind): item}, image)
    with Image.open(image) as rendered:
        assert rendered.format == "PNG" and rendered.width > 1000


def test_mac_report_generation_preserves_objects_and_sources(report_inputs, tmp_path, monkeypatch):
    paths, profile = report_inputs
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setattr(engine, "require_excel_backend", lambda: pytest.fail("Report must not request Excel"))
    hashes = {p: Path(p).read_bytes() for p in paths.values()}
    inspected = inspect_report_inputs(**paths)
    assert len(inspected.samples) == 2 and inspected.missing_fields == ["操作人员"]
    result = generate_report(**paths, output_directory=str(tmp_path / "bundle"))
    assert result.integrity["valid"] and not result.computation["office_required"]
    assert verify_integrity(paths["template_path"], result.output_path, result.integrity_report_path).valid
    before, after = Package(paths["template_path"]), Package(result.output_path)
    assert all(after.parts[p] == before.parts[p] for p in before.parts if p not in result.integrity["changed_parts"])
    assert after.parts["xl/media/image1.png"] == before.parts["xl/media/image1.png"]
    assert after.parts["xl/ctrlProps/ctrlProp1.xml"] == before.parts["xl/ctrlProps/ctrlProp1.xml"]
    assert after.parts["xl/drawings/vmlDrawing1.vml"] == before.parts["xl/drawings/vmlDrawing1.vml"]
    book = Workbook(after)
    cells = book.cell_map(book.sheets[0])
    assert book.value(cells["F12"]) == 2 and book.value(cells["H12"]) == 70.71
    assert book.value(cells["E2"]) == "Sample-A" and book.value(cells["F2"]) == "Sample-B"
    assert book.value(cells["D3"]) == "" and book.value(cells["D2"]) == "2026-09-10"
    assert len(book.images()) == 3 and Path(result.image_paths[0]).is_file()
    assert all(Path(p).read_bytes() == b for p, b in hashes.items())
    assert not list(tmp_path.glob(".xmlmcp-report-*"))


@pytest.mark.parametrize("problem,expected", [("capacity", "REPORT_CAPACITY"), ("fingerprint", "REPORT_TEMPLATE_CHANGED"),
                                                ("samples", "SAMPLE_MISMATCH"), ("formula", "OFFLINE_FORMULA"),
                                                ("local_chart", "OFFLINE_CHART_DEPENDENCY"), ("collision", "OUTPUT_EXISTS")])
def test_report_failures_are_atomic(report_inputs, tmp_path, problem, expected):
    paths, profile = report_inputs
    output = tmp_path / "blocked"
    if problem == "capacity":
        profile["tables"]["single_summary"]["target"]["range"] = "A10:H11"
    elif problem == "fingerprint":
        profile["template_fingerprint"] = "0" * 64
    elif problem == "samples":
        with zipfile.ZipFile(paths["raw_zip_path"], "a") as z:
            z.writestr("单点/Unknown.csv", instrument_csv("single_point"))
    elif problem in ("formula", "local_chart"):
        template = write_template(tmp_path / "other-template.xlsx", formula=problem == "formula", chart=problem == "local_chart")
        paths["template_path"] = str(template)
        profile["template_fingerprint"] = Package(template).fingerprint
    elif problem == "collision":
        output.mkdir()
        (output / "keep.txt").write_text("keep")
    Path(paths["profile_path"]).write_text(json.dumps(profile), encoding="utf-8")
    with pytest.raises(XMLMCPError) as caught:
        generate_report(**paths, output_directory=str(output))
    assert caught.value.code == expected
    assert output.exists() == (problem == "collision")
    if output.exists():
        assert list(output.iterdir()) == [output / "keep.txt"]
    assert not list(tmp_path.glob(".xmlmcp-report-*"))


def test_different_report_layout_uses_explicit_mapping(report_inputs, tmp_path):
    paths, profile = report_inputs
    second = write_template(tmp_path / "new-layout.xlsx", prefix=True)
    paths["template_path"] = str(second)
    profile["template_fingerprint"] = Package(second).fingerprint
    profile["fields"]["client_company"] = [{"sheet": "报告", "range": "F1"}]
    profile["tables"] = {}
    profile["figures"] = {"pressure": {"sheet": "报告", "range": "A11:H20"}}
    Path(paths["profile_path"]).write_text(json.dumps(profile), encoding="utf-8")
    result = generate_report(**paths, output_directory=str(tmp_path / "different-layout"))
    book = Workbook(Package(result.output_path))
    assert book.value(book.cell_map(book.sheets[0])["F1"]) == "Old title"
    assert len(book.images()) == 1 and result.integrity["valid"]


def test_inline_profile_is_saved_for_reuse(report_inputs, tmp_path):
    paths, profile = report_inputs
    paths.pop("profile_path")
    result = generate_report(**paths, output_directory=str(tmp_path / "inline-profile"), profile=profile)
    assert json.loads((Path(result.output_directory) / "template-profile.json").read_text(encoding="utf-8"))["template_fingerprint"] == profile["template_fingerprint"]


def test_bundle_publication_keeps_existing_empty_directory(tmp_path):
    bundle, output = tmp_path / "bundle", tmp_path / "existing"
    bundle.mkdir()
    (bundle / "report.xlsx").write_bytes(b"report")
    output.mkdir()
    with pytest.raises(XMLMCPError) as caught:
        publish_bundle(bundle, output)
    assert caught.value.code == "OUTPUT_EXISTS"
    assert not list(output.iterdir()) and (bundle / "report.xlsx").read_bytes() == b"report"


def test_invalid_instrument_sample_count_is_business_error():
    data = instrument_csv("single_point").decode().replace("采样数目:2", "采样数目:unknown")
    with pytest.raises(XMLMCPError) as caught:
        parse_csv(data.encode(), "Sample-A.csv")
    assert caught.value.code == "RAW_SAMPLE_COUNT"


def test_offline_order_formula_cache_is_not_used(report_inputs, tmp_path):
    paths, profile = report_inputs
    order = write_template(tmp_path / "order-with-formula.xlsx", formula=True)
    paths["order_path"] = str(order)
    profile["order"]["fields"]["client_company"]["range"] = "D5"
    Path(paths["profile_path"]).write_text(json.dumps(profile))
    with pytest.raises(XMLMCPError) as caught:
        inspect_report_inputs(**paths)
    assert caught.value.code == "ORDER_FORMULA_INPUT"


def test_zip_unknown_files_and_unsafe_paths_not_silently_skipped(tmp_path):
    for name, code in (("../Sample-A.csv", "RAW_ZIP_PATH"), ("unknown.xlsx", "RAW_FORMAT")):
        raw = tmp_path / "bad.zip"
        with zipfile.ZipFile(raw, "w") as z:
            z.writestr(name, instrument_csv("single_point"))
        with pytest.raises(XMLMCPError) as caught:
            read_raw_zip(str(raw))
        assert caught.value.code == code


def test_targeted_rich_text_keeps_all_run_properties_and_shared_strings(tmp_path):
    template = write_template(tmp_path / "mixed.xlsx", mixed=True)
    book = Workbook(Package(template))
    shared = book.package.parts["xl/sharedStrings.xml"]
    before = book.string_node(book.cell_map(book.sheets[0])["A2"])
    props = [before.document.raw(r.child("rPr", S)) for r in before.children if r.child("rPr", S)]
    apply_text_run_values(book, [{"target": {"sheet": "报告", "range": "A2"}, "run_index": 1, "value": "\nNew method"}])
    after = book.string_node(book.cell_map(book.sheets[0])["A2"])
    assert book.value(book.cell_map(book.sheets[0])["A2"]) == "Small English\nNew method"
    assert props == [after.document.raw(r.child("rPr", S)) for r in after.children if r.child("rPr", S)]
    assert book.package.parts["xl/sharedStrings.xml"] == shared


@pytest.mark.parametrize("mode", ["legacy", "auto"])
def test_report_tools_in_real_stdio(report_inputs, tmp_path, mode):
    paths, _ = report_inputs
    async def run():
        async with Client(StdioServerParameters(command=sys.executable, args=["-m", "xmlmcp"]), mode=mode, read_timeout_seconds=60) as client:
            inspected = await client.call_tool("inspect_report_inputs", paths)
            assert not inspected.is_error and len(inspected.structured_content["samples"]) == 2
            generated = await client.call_tool("generate_report", paths | {"output_directory": str(tmp_path / ("mcp-" + mode))})
            assert not generated.is_error, generated.content
            result = generated.structured_content
            verified = await client.call_tool("verify_integrity", {"input_path": paths["template_path"], "output_path": result["output_path"], "report_path": result["integrity_report_path"]})
            assert not verified.is_error and verified.structured_content["valid"]
            failed = await client.call_tool("generate_report", paths | {"output_directory": result["output_directory"]})
            assert failed.is_error and failed.structured_content["error"]["code"] == "OUTPUT_EXISTS"
    asyncio.run(run())
