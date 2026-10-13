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
from xmlmcp.namespaces import S, XDR, P, CT, REL_BASE
from xmlmcp.report_objects import drawing_objects, apply_example_removals
from xmlmcp.report_styles import verify_report_styles


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


def test_report_preserves_individual_table_styles_even_with_legacy_style_mapping(report_inputs, tmp_path):
    paths, profile = report_inputs
    original = Package(paths['template_path'])
    sheet = original.parts['xl/worksheets/sheet1.xml']
    rows = ''.join(f'<row r="{r}" ht="22" customHeight="1">' + ''.join(
        f'<c r="{chr(65 + c)}{r}" s="{(r + c) % 2}"/>' for c in range(8)) + '</row>' for r in range(10, 14))
    original.parts['xl/worksheets/sheet1.xml'] = sheet.replace(b'<row r="20"/>', rows.encode() + b'<row r="20"/>')
    template = tmp_path / 'styled-table.xlsx'
    original.write(template)
    paths['template_path'] = str(template)
    profile['template_fingerprint'] = Package(template).fingerprint
    slot = profile['tables']['single_summary']
    slot['header_style'] = {'sheet': '报告', 'range': 'A1'}
    slot['data_style'] = {'sheet': '报告', 'range': 'C5'}
    Path(paths['profile_path']).write_text(json.dumps(profile), encoding='utf-8')
    result = generate_report(**paths, output_directory=str(tmp_path / 'original-styles'))
    before, after = Package(template), Package(result.output_path)
    checks = verify_report_styles(before, after)
    assert checks['destination_styles_identical'] and checks['layout_styles_identical']
    old_book, new_book = Workbook(before), Workbook(after)
    old_cells, new_cells = old_book.cell_map(old_book.sheets[0]), new_book.cell_map(new_book.sheets[0])
    assert {a: c.attrs.get('s') for a, c in old_cells.items()} == {a: new_cells[a].attrs.get('s') for a in old_cells}
    assert new_book.value(new_cells['A10']) == '样品' and new_book.value(new_cells['A12']) == 'Sample-A'
    manifest = json.loads(Path(result.integrity_report_path).read_text(encoding='utf-8'))
    assert not any(op['op'] == 'copy_style' for op in manifest['operations'])
    assert manifest['style_preservation']['style_parts_identical']
    assert verify_integrity(str(template), result.output_path, result.integrity_report_path).checks['destination_styles_identical']
    after.parts['xl/worksheets/sheet1.xml'] = after.parts['xl/worksheets/sheet1.xml'].replace(b'r="A10" s="0"', b'r="A10" s="1"')
    with pytest.raises(XMLMCPError) as caught:
        verify_report_styles(before, after)
    assert caught.value.code == 'REPORT_STYLE_CHANGED'


def test_mapped_chart_example_is_removed_without_recalculation(report_inputs, tmp_path):
    paths, profile = report_inputs
    template = write_template(tmp_path / 'example-chart.xlsx', image=True, chart=True)
    before = Package(template)
    paths['template_path'] = str(template)
    profile['template_fingerprint'] = before.fingerprint
    chart = next(o for o in drawing_objects(Workbook(before)) if o['kind'] == 'chart')
    profile['example_objects'] = [{'object_id': chart['object_id'], 'kind': 'chart'}]
    Path(paths['profile_path']).write_text(json.dumps(profile), encoding='utf-8')
    result = generate_report(**paths, output_directory=str(tmp_path / 'replaced-chart'))
    after = Package(result.output_path)
    assert not Workbook(after).charts() and 'xl/charts/chart1.xml' not in after.parts
    assert len(Workbook(after).images()) == 3
    assert after.parts['xl/media/image1.png'] == before.parts['xl/media/image1.png']
    assert after.parts['xl/drawings/vmlDrawing1.vml'] == before.parts['xl/drawings/vmlDrawing1.vml']
    assert after.parts['xl/comments/comment1.xml'] == before.parts['xl/comments/comment1.xml']
    assert verify_integrity(str(template), result.output_path, result.integrity_report_path).valid


def test_shared_example_media_is_kept_until_last_object_removed(tmp_path):
    template = write_template(tmp_path / 'shared-example.xlsx', image=True)
    book = Workbook(Package(template))
    objects = drawing_objects(book)
    first, second = objects
    result = apply_example_removals(book, [{'object_id': first['object_id'], 'kind': 'image'}])
    assert not result['removed_parts'] and len(book.images()) == 1
    assert 'xl/media/image1.png' in book.package.parts
    result = apply_example_removals(book, [{'object_id': second['object_id'], 'kind': 'image'}])
    assert result['removed_parts'] == ['xl/media/image1.png'] and not book.images()
    book.package.validate()


def test_example_chart_dependencies_keep_shared_parts(tmp_path):
    template = write_template(tmp_path / 'dependencies.xlsx', chart=True, image=True)
    package = Package(template)
    # The chart's style is exclusive, but its image is also used by two pictures.
    package.parts['xl/charts/style1.xml'] = b'<style xmlns="urn:test"/>'
    package.parts['xl/charts/_rels/chart1.xml.rels'] = (
        f'<Relationships xmlns="{P}"><Relationship Id="style" Type="{REL_BASE}chartStyle" Target="style1.xml"/>'
        f'<Relationship Id="image" Type="{REL_BASE}image" Target="../media/image1.png"/></Relationships>').encode()
    source = tmp_path / 'with-dependencies.xlsx'
    package.write(source)
    book = Workbook(Package(source))
    chart = next(o for o in drawing_objects(book) if o['kind'] == 'chart')
    apply_example_removals(book, [{'object_id': chart['object_id'], 'kind': 'chart'}])
    assert 'xl/charts/style1.xml' not in book.package.parts
    assert 'xl/charts/_rels/chart1.xml.rels' not in book.package.parts
    assert 'xl/media/image1.png' in book.package.parts and len(book.images()) == 2
    book.package.validate()


def test_unknown_example_selection_blocks_publication(report_inputs, tmp_path):
    paths, profile = report_inputs
    profile['example_objects'] = [{'object_id': 'unknown|drawing|10', 'kind': 'chart'}]
    Path(paths['profile_path']).write_text(json.dumps(profile), encoding='utf-8')
    output = tmp_path / 'bad-object'
    with pytest.raises(XMLMCPError) as caught:
        generate_report(**paths, output_directory=str(output))
    assert caught.value.code == 'REPORT_OBJECT_MAPPING' and not output.exists()


@pytest.mark.parametrize('kind', ['single_resistance','single_resistivity','single_conductivity','single_thickness',
                                 'pressure_resistivity','pressure_thickness'])
def test_reference_charts_compute_real_statistics_and_keep_transparency(tmp_path, monkeypatch, kind):
    from PIL import Image
    from xmlmcp import report_plots
    monkeypatch.setattr(report_plots,'SYSTEM_FONTS',())
    monkeypatch.setattr(report_plots,'SYSTEM_BOLD_FONTS',())
    mode=report_plots.plot_mode(kind)
    item=parse_csv(instrument_csv(mode),'Sample-A.csv')
    image=tmp_path/(kind+'.png')
    with warnings.catch_warnings():
        warnings.filterwarnings('error',message='Glyph .* missing from font')
        record=report_plots.make_plot(kind,[{'sample_name':'Sample-A'}],{('Sample-A',mode):item},image,style='template-statistics')
    with Image.open(image) as rendered:
        assert rendered.mode=='RGBA' and rendered.getpixel((0,0))[3]==0
        assert json.loads(rendered.info['XMLMCPStyle'])['background']=='transparent'
    data=record['data']['series'][0]
    if kind=='single_resistance':assert data['mean']==10 and data['cov_pct']==0
    if kind=='single_resistivity':
        assert data['mean']==2 and data['cov_pct']==pytest.approx(100/math.sqrt(2))
    if kind=='pressure_thickness':assert data['x']==[5,10] and data['y']==[70,70]
    assert record['style']['font']=='XMLMCPChartLabelsBold.ttf'


def test_reference_image_replacement_cleans_only_exclusive_old_media(report_inputs,tmp_path):
    paths,profile=report_inputs
    before=Workbook(Package(paths['template_path']))
    images=before.images()
    profile['figures']={}
    profile['figure_replacements']={'single_resistivity':images[0]['image_id'],'pressure_thickness':images[1]['image_id']}
    profile['figure_style']='template-statistics'
    Path(paths['profile_path']).write_text(json.dumps(profile),encoding='utf-8')
    result=generate_report(**paths,output_directory=str(tmp_path/'replaced-statistics'))
    after=Workbook(Package(result.output_path))
    assert len(after.images())==2 and 'xl/media/image1.png' not in after.package.parts
    assert [i['image_id'] for i in after.images()]==[i['image_id'] for i in images]
    assert verify_integrity(paths['template_path'],result.output_path,result.integrity_report_path).valid
    for old,new in zip(images,after.images()):
        assert old['_doc'].raw(old['_pic'].child('spPr',XDR))==new['_doc'].raw(new['_pic'].child('spPr',XDR))
        assert old['name']==new['name']


def test_measurement_tables_use_fixed_mean_rows_and_original_styles(report_inputs,tmp_path):
    paths,profile=report_inputs
    profile['tables']={'single_measurements':{'target':{'sheet':'报告','range':'A10:E17'}},
                       'pressure_measurements':{'target':{'sheet':'报告','range':'A20:E25'}}}
    profile['figures']={}
    Path(paths['profile_path']).write_text(json.dumps(profile),encoding='utf-8')
    result=generate_report(**paths,output_directory=str(tmp_path/'point-tables'))
    book=Workbook(Package(result.output_path));cells=book.cell_map(book.sheets[0])
    assert book.value(cells['B11'])=='Sample-A' and book.value(cells['D11'])=='Sample-A'
    assert book.value(cells['B12'])==1 and book.value(cells['B13'])==3
    assert book.value(cells['B14']) is None and book.value(cells['B15']) is None
    assert book.value(cells['A16'])=='平均值' and book.value(cells['B16'])==2
    assert book.value(cells['A17'])=='COV（%）' and book.value(cells['B17'])==pytest.approx(70.7107)
    assert result.integrity['checks']['destination_styles_identical']


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
