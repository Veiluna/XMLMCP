from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from PIL import Image

from xmlmcp.engine import apply_changes, inspect_template
from xmlmcp.integrity import verify_integrity
from xmlmcp.package import Package
from xmlmcp.workbook import Workbook
from xmlmcp.xmlbytes import XMLDocument, Edit, apply_edits
from xmlmcp.errors import XMLMCPError
from xmlmcp.namespaces import S, A, XDR
from conftest import write_template, GeometryCalculator


def write(target, value):
    return {"op": "write", "target": target, "value": value}


def test_single_part_change_and_uniform_font(template):
    result = apply_changes(str(template), str(template.with_name("result.xlsx")), [write({"name": "SmallFont"}, " New English & 中文 ")])
    assert result.integrity["changed_parts"] == ["xl/worksheets/sheet1.xml"]
    a, b = Package(template), Package(result.output_path)
    for part in a.parts.keys() - {"xl/worksheets/sheet1.xml"}:
        assert a.parts[part] == b.parts[part]
    doc = b.doc("xl/worksheets/sheet1.xml")
    cell = next(n for n in doc.root.descendants("c", S) if n.attrs["r"] == "A2")
    assert cell.attrs["t"] == "inlineStr"
    assert next(cell.descendants("sz", S)).attrs["val"] == "7"
    assert Workbook(b).value(cell) == " New English & 中文 "
    assert verify_integrity(str(template), result.output_path, result.report_path).valid


@pytest.mark.parametrize("prefix", [False, True])
def test_values_matrix_new_cells_rows_and_styles(tmp_path, prefix):
    template = write_template(tmp_path / "template.xlsx", prefix=prefix)
    operations = [write({"name": "Field", "scope": "workbook"}, "=literal"),
                  write({"name": "Grid"}, [["新的甲", 100], ["新的乙", True], [None, 9.25]]),
                  write({"sheet": "报告", "range": "B19"}, "new row"),
                  {"op": "copy_style", "source": {"sheet": "报告", "range": "C5"}, "target": {"sheet": "报告", "range": "C7:C8"}}]
    result = apply_changes(str(template), str(tmp_path / "result.xlsx"), operations)
    book = Workbook(Package(result.output_path))
    cells = book.cell_map(book.sheet("报告"))
    assert book.value(cells["A1"]) == "=literal"
    assert cells["A1"].attrs["s"] == "1"
    assert book.value(cells["C6"]) is True
    assert book.value(cells["C7"]) == 9.25
    assert book.value(cells["B19"]) == "new row"
    assert cells["C7"].attrs["s"] == cells["C8"].attrs["s"] == "1"
    assert verify_integrity(str(template), result.output_path, result.report_path).valid


@pytest.mark.parametrize("operation,code", [
    (write({"name": "Field"}, "new"), "AMBIGUOUS_NAME"),
    (write({"name": "Dynamic"}, 3), "UNSUPPORTED_NAME"),
    (write({"sheet": "报告", "range": "XFE1"}, 3), "INVALID_RANGE"),
    (write({"name": "Grid"}, [[1, 2, 3]]), "MATRIX_OUT_OF_RANGE"),
    (write({"name": "Grid"}, [[1], [1, 2]]), "INVALID_MATRIX"),
    (write({"sheet": "报告", "range": "C5"}, 1234567890123456), "NUMBER_PRECISION"),
    (write({"sheet": "报告", "range": "C5"}, "bad\x00text"), "INVALID_XML_TEXT"),
])
def test_preflight_failures_no_outputs(template, operation, code):
    output = template.with_name("failed.xlsx")
    fingerprint = Package(template).fingerprint
    with pytest.raises(XMLMCPError) as error:
        apply_changes(str(template), str(output), [write({"sheet": "报告", "range": "B5"}, "valid first"), operation])
    assert error.value.code == code
    assert not output.exists() and not output.with_suffix(".integrity.json").exists()
    assert Package(template).fingerprint == fingerprint


def test_mixed_rich_text_rejected(tmp_path):
    template = write_template(tmp_path / "template.xlsx", mixed=True)
    inspected = inspect_template(str(template), [{"name": "SmallFont"}])
    assert inspected.targets[0]["cells"][0]["character_format"] == "mixed_rich_text"
    with pytest.raises(XMLMCPError, match="混合字符格式"):
        apply_changes(str(template), str(tmp_path / "result.xlsx"), [write({"name": "SmallFont"}, "new")])


def test_merged_anchor_and_non_anchor(tmp_path):
    template = write_template(tmp_path / "template.xlsx", merged=True)
    result = apply_changes(str(template), str(tmp_path / "result.xlsx"), [write({"sheet": "报告", "range": "F2:G3"}, "merged")])
    book = Workbook(Package(result.output_path))
    assert book.value(book.cell_map(book.sheet("报告"))["F2"]) == "merged"
    with pytest.raises(XMLMCPError, match="左上角"):
        apply_changes(str(template), str(tmp_path / "bad.xlsx"), [write({"sheet": "报告", "range": "G3"}, "bad")])


def test_shared_image_replacement_and_new_image(tmp_path):
    template = write_template(tmp_path / "template.xlsx", image=True)
    new_image = tmp_path / "replacement.png"
    Image.new("RGB", (200, 50), "red").save(new_image)
    before = Workbook(Package(template))
    original_images = before.images()
    other_pic = original_images[1]["_doc"].raw(original_images[1]["_pic"])
    anchor = original_images[0]["_pic"].parent
    from_bytes = anchor.document.raw(anchor.child("from", XDR))
    ops = [{"op": "replace_image", "image_id": original_images[0]["image_id"], "image_path": str(new_image)},
           {"op": "add_image", "target": {"name": "Photo"}, "image_path": str(new_image)}]
    result = apply_changes(str(template), str(tmp_path / "result.xlsx"), ops, calculator=GeometryCalculator())
    after = Workbook(Package(result.output_path))
    images = after.images()
    assert len(images) == 3
    assert images[1]["_doc"].raw(images[1]["_pic"]) == other_pic
    assert images[0]["_pic"].parent.document.raw(images[0]["_pic"].parent.child("from", XDR)) == from_bytes
    assert after.package.parts["xl/media/image1.png"] == before.package.parts["xl/media/image1.png"]
    fill = next(images[0]["_pic"].descendants("fillRect", A))
    assert fill.attrs == {"l": "0", "r": "0", "t": "37500", "b": "37500"}
    crop = next(images[0]["_pic"].descendants("srcRect", A))
    assert all(crop.attrs[k] == "0" for k in ("l", "r", "t", "b"))
    assert verify_integrity(str(template), result.output_path, result.report_path).valid


def test_insert_drawing_without_existing_drawing(template):
    image = template.with_name("new.png")
    Image.new("RGB", (40, 80), "green").save(image)
    result = apply_changes(str(template), str(template.with_name("result.xlsx")),
                           [{"op": "add_image", "target": {"name": "Photo"}, "image_path": str(image)}], calculator=GeometryCalculator())
    assert len(Workbook(Package(result.output_path)).images()) == 1
    assert verify_integrity(str(template), result.output_path, result.report_path).valid


def test_outside_patch_tampering_detected(template):
    result = apply_changes(str(template), str(template.with_name("result.xlsx")), [write({"sheet": "报告", "range": "C5"}, 999)])
    package = Package(result.output_path)
    package.parts["customXml/unknown.xml"] = package.parts["customXml/unknown.xml"].replace(b' b="2"', b' b="3"')
    package.write(result.output_path)
    verified = verify_integrity(str(template), result.output_path, result.report_path)
    assert not verified.valid


def test_no_overwrite_and_conflict(template):
    output = template.with_name("existing.xlsx")
    output.write_bytes(b"keep")
    with pytest.raises(XMLMCPError, match="已存在"):
        apply_changes(str(template), str(output), [write({"sheet": "报告", "range": "C5"}, 3)])
    assert output.read_bytes() == b"keep"
    with pytest.raises(XMLMCPError, match="重复写入"):
        apply_changes(str(template), str(template.with_name("conflict.xlsx")),
                      [write({"sheet": "报告", "range": "C5"}, 1), write({"sheet": "报告", "range": "C5"}, 2)])


def test_byte_spans_quoted_angle_and_non_ascii():
    data = '<s:r xmlns:s="urn:test" a="&gt;">汉<!-->--><s:v>旧</s:v><s:x/></s:r>'.encode()
    doc = XMLDocument(data)
    v = doc.root.child("v")
    new = '<s:v>新</s:v>'.encode()
    output, records = apply_edits(data, [Edit(v.start, v.end, new, "cell:value", "test")])
    assert output == data.replace('<s:v>旧</s:v>'.encode(), new)
    assert len(records) == 1


def test_calculation_failure_atomic(tmp_path):
    template = write_template(tmp_path / "template.xlsx", formula=True)
    class Fail:
        def run(self, *args):
            raise XMLMCPError("CALCULATION_TIMEOUT", "timeout")
    with pytest.raises(XMLMCPError):
        apply_changes(str(template), str(tmp_path / "result.xlsx"), [write({"sheet": "报告", "range": "C5"}, 999)], calculator=Fail())
    assert not (tmp_path / "result.xlsx").exists()
    assert not list(tmp_path.glob(".xmlmcp-*"))

