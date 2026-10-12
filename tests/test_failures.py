from __future__ import annotations

import json
from pathlib import Path

import pytest

from conftest import write_template
from xmlmcp.engine import apply_changes
from xmlmcp.errors import XMLMCPError
from xmlmcp.package import Package
from xmlmcp.workbook import Workbook
from xmlmcp.calculation import preflight_calculation, harvest_formula_caches, harvest_chart_caches
from xmlmcp.cells import CellChange, apply_cell_changes
from xmlmcp.xmlbytes import XMLDocument, Edit
from xmlmcp.namespaces import S


def test_formula_caches_preserve_formula_and_cell_style(tmp_path):
    source = write_template(tmp_path / "source.xlsx", formula=True)
    original = Workbook(Package(source))
    computed = Workbook(Package(source))
    apply_cell_changes(computed, {"报告": {"D5": CellChange(value="300", cached=True, cache_type="n", preserve_formula=True)}})
    doc = original.sheet_doc(original.sheet("报告"))
    formula = next(doc.root.descendants("f", S))
    raw_formula = doc.raw(formula)
    count = harvest_formula_caches(original, computed)
    assert count == 1
    node = original.cell_map(original.sheet("报告"))["D5"]
    assert node.document.raw(node.child("f", S)) == raw_formula
    assert node.child("v", S).text == "300"


def test_stale_chart_cache_rejected(tmp_path):
    source = write_template(tmp_path / "source.xlsx", chart=True)
    a, b = Workbook(Package(source)), Workbook(Package(source))
    observation = {"charts": [{"sheet": "报告", "name": "Chart 1", "series": [{"name": "", "values": [100, 200], "categories": ["甲", "乙"]}]}]}
    with pytest.raises(XMLMCPError) as caught:
        harvest_chart_caches(a, b, observation)
    assert caught.value.code == "CHART_CACHE_STALE"


def test_unknown_binary_untouched_and_utf16_opaque_xml(template):
    package = Package(template)
    package.parts["customXml/unknown.xml"] = '<?xml version="1.0" encoding="UTF-16"?><opaque>保持</opaque>'.encode("utf-16")
    package.write(template)
    before = Package(template).parts["customXml/unknown.xml"]
    result = apply_changes(str(template), str(template.with_name("result.xlsx")), [{"op": "write", "target": {"sheet": "报告", "range": "C5"}, "value": 99}])
    assert Package(result.output_path).parts["customXml/unknown.xml"] == before


def test_failure_during_second_publish_rolls_back_own_report(template, monkeypatch):
    import xmlmcp.engine as engine
    actual = engine._publish_noclobber
    def race(source, destination):
        if destination.suffix == ".xlsx":
            destination.write_bytes(b"other writer")
        return actual(source, destination)
    monkeypatch.setattr(engine, "_publish_noclobber", race)
    output = template.with_name("race.xlsx")
    with pytest.raises(FileExistsError):
        apply_changes(str(template), str(output), [{"op": "clear", "target": {"sheet": "报告", "range": "C5"}}])
    assert output.read_bytes() == b"other writer"
    assert not output.with_suffix(".integrity.json").exists()


def test_non_network_formula_preflight(tmp_path):
    template = write_template(tmp_path / "template.xlsx", formula=True)
    package = Package(template)
    book = Workbook(package)
    apply_cell_changes(book, {"报告": {"D5": CellChange(formula='WEBSERVICE("https://example.com")')}})
    with pytest.raises(XMLMCPError) as exc:
        preflight_calculation(book, True)
    assert exc.value.code == "OFFLINE_FORMULA_UNSUPPORTED"


def test_clear_null_boolean_and_style_only(template):
    result = apply_changes(str(template), str(template.with_name("result.xlsx")), [
        {"op": "clear", "target": {"sheet": "报告", "range": "C5"}},
        {"op": "write", "target": {"sheet": "报告", "range": "C6"}, "value": False},
        {"op": "copy_style", "source": {"sheet": "报告", "range": "A1"}, "target": {"sheet": "报告", "range": "B5"}}])
    book = Workbook(Package(result.output_path))
    cells = book.cell_map(book.sheet("报告"))
    assert book.value(cells["C5"]) is None
    assert cells["C5"].attrs["s"] == "1"
    assert book.value(cells["C6"]) is False
    assert cells["B5"].attrs["s"] == "1" and book.value(cells["B5"]) == "甲"


def test_special_array_formula_and_report_corruption(tmp_path):
    template = write_template(tmp_path / "template.xlsx", formula=True)
    package = Package(template)
    part = "xl/worksheets/sheet1.xml"
    package.parts[part] = package.parts[part].replace(b'<f>SUM(C5:C6)</f>', b'<f t="array" ref="D5:E6">SUM(C5:C6)</f>')
    package.write(template)
    with pytest.raises(XMLMCPError) as caught:
        apply_changes(str(template), str(tmp_path / "result.xlsx"), [{"op": "write", "target": {"sheet": "报告", "range": "E6"}, "value": 3}])
    assert caught.value.code == "SPECIAL_FORMULA_TARGET"


def test_chart_cache_extensions_and_point_metadata_preserved(tmp_path):
    from xmlmcp.calculation import chart_cache_edits
    from xmlmcp.xmlbytes import apply_edits
    from xmlmcp.namespaces import C
    before = f'<c:numCache xmlns:c="{C}" xmlns:k="urn:keep"><c:formatCode>General</c:formatCode><c:ptCount val="2" k:keep="yes"/><c:pt idx="0" k:keep="point"><c:v>10</c:v><k:metadata/></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:extLst><c:ext uri="keep"><k:opaque a="2"/></c:ext></c:extLst></c:numCache>'.encode()
    after = f'<c:numCache xmlns:c="{C}"><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>200</c:v></c:pt></c:numCache>'.encode()
    a, b = XMLDocument(before), XMLDocument(after)
    edits = chart_cache_edits(a, a.root, b, b.root, "test")
    output = apply_edits(before, edits)[0]
    assert output == before.replace(b'<c:v>10</c:v>', b'<c:v>100</c:v>').replace(b'<c:v>20</c:v>', b'<c:v>200</c:v>')


def test_empty_chart_cache_can_receive_headers_and_points():
    from xmlmcp.calculation import chart_cache_edits
    from xmlmcp.xmlbytes import apply_edits
    from xmlmcp.namespaces import C
    old = XMLDocument(f'<c:numCache xmlns:c="{C}"></c:numCache>'.encode())
    new = XMLDocument(f'<c:numCache xmlns:c="{C}"><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>100</c:v></c:pt></c:numCache>'.encode())
    output = XMLDocument(apply_edits(old.data, chart_cache_edits(old, old.root, new, new.root, "test"))[0])
    assert [n.local for n in output.root.children] == ["formatCode", "ptCount", "pt"]


def test_chart_literal_series_name_is_verified():
    from xmlmcp.calculation import verify_chart_values
    from xmlmcp.namespaces import C
    doc = XMLDocument(f'<c:chart xmlns:c="{C}"><c:ser><c:tx><c:v>固定标题</c:v></c:tx><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser></c:chart>'.encode())
    chart = {"name": "Chart", "_doc": doc}
    observed = {"series": [{"name": "固定标题", "values": [30], "categories": [1]}]}
    assert verify_chart_values(chart, observed) == 2
    observed["series"][0]["name"] = "错误标题"
    with pytest.raises(XMLMCPError) as caught:
        verify_chart_values(chart, observed)
    assert caught.value.code == "CHART_CACHE_STALE"


def test_duplicate_picture_identifiers_rejected(tmp_path):
    from xmlmcp.images import preflight_images
    from xmlmcp.models import ReplaceImage
    path = write_template(tmp_path / "duplicate.xlsx", image=True)
    package = Package(path)
    part = "xl/drawings/drawing1.xml"
    package.parts[part] = package.parts[part].replace(b'id="2" name="Picture 2"', b'id="1" name="Picture 2"')
    book = Workbook(package)
    image_id = book.images()[0]["image_id"]
    with pytest.raises(XMLMCPError) as caught:
        preflight_images(book, [ReplaceImage(op="replace_image", image_id=image_id, image_path=str(tmp_path / "original.png"))])
    assert caught.value.code == "IMAGE_MAPPING"


def test_escaped_strings_and_excel_text_limit(template):
    from xmlmcp.strings import decode_xstring
    text = " _x0041_ _x005F_ 甲\r\n乙\t末 😀 "
    result = apply_changes(str(template), str(template.with_name("text.xlsx")), [
        {"op": "write", "target": {"sheet": "报告", "range": "B5"}, "value": text}])
    book = Workbook(Package(result.output_path))
    assert book.value(book.cell_map(book.sheet("报告"))["B5"]) == text
    assert decode_xstring("_xD83D__xDE00_") == "😀"
    with pytest.raises(XMLMCPError) as caught:
        apply_changes(str(template), str(template.with_name("oversized.xlsx")), [
            {"op": "write", "target": {"sheet": "报告", "range": "B5"}, "value": "a" * 32768}])
    assert caught.value.code == "TEXT_TOO_LONG"


def test_phonetic_text_is_read_correctly_and_write_rejected(template):
    package = Package(template)
    part = "xl/worksheets/sheet1.xml"
    package.parts[part] = package.parts[part].replace('<t>甲</t>'.encode(), '<t>甲</t><rPh sb="0" eb="1"><t>jia</t></rPh>'.encode())
    package.write(template)
    book = Workbook(Package(template))
    assert book.value(book.cell_map(book.sheet("报告"))["B5"]) == "甲"
    with pytest.raises(XMLMCPError) as caught:
        apply_changes(str(template), str(template.with_name("phonetic.xlsx")), [
            {"op": "write", "target": {"sheet": "报告", "range": "B5"}, "value": "乙"}])
    assert caught.value.code == "PHONETIC_TEXT_UNSUPPORTED"

