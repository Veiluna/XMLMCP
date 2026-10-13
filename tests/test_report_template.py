"""Template authoring is explicit; filled reports then preserve its styles."""
import hashlib
import re

import pytest

from conftest import write_template
from xmlmcp.errors import XMLMCPError
from xmlmcp.integrity import verify_integrity
from xmlmcp.namespaces import S, XDR
from xmlmcp.package import Package
from xmlmcp.report_template import prepare_electrode_template, reference_tables
from xmlmcp.workbook import Workbook
from xmlmcp.xmlbytes import standalone


def mother_fixture(path, *, prefix=False):
    write_template(path, image=True, chart=True, prefix=prefix, merged=True)
    package = Package(path)
    names = ('<definedName name="单点结果区">\'报告\'!$A$74:$K$125</definedName>'
             '<definedName name="变压结果区">\'报告\'!$A$129:$K$163</definedName>').encode()
    package.parts['xl/workbook.xml'] = package.parts['xl/workbook.xml'].replace(b'</definedNames>', names + b'</definedNames>')
    style = package.parts['xl/styles.xml'].replace(b'<fonts ', b'<numFmts count="0"/><fonts ', 1)
    if prefix:
        style = re.sub(br'<(/?)([A-Za-z]+)', br'<\1s:\2', style).replace(
            b'xmlns="' + S.encode() + b'"', b'xmlns:s="' + S.encode() + b'"')
    package.parts['xl/styles.xml'] = style
    # Put a local sample chart in the appendix so the authoring step removes it.
    package.parts['xl/drawings/drawing1.xml'] = package.parts['xl/drawings/drawing1.xml'].replace(
        b'<xdr:row>10</xdr:row>', b'<xdr:row>200</xdr:row>').replace(b'<xdr:row>20</xdr:row>', b'<xdr:row>210</xdr:row>')
    package.write(path.with_name('mother.xlsx'))
    return path.with_name('mother.xlsx')


def reference_fixture(path):
    write_template(path, image=True)
    package = Package(path)
    image = Workbook(package).images()[0]
    raw = standalone(image['_doc'], image['_pic'].parent)
    anchors = b''.join(raw.replace(b'id="1"', f'id="{identifier}"'.encode(), 1)
                       for identifier in (21, 22, 31, 32, 42, 43))
    package.parts['xl/drawings/drawing1.xml'] = (
        b'<xdr:wsDr xmlns:xdr="' + XDR.encode() + b'">' + anchors + b'</xdr:wsDr>')
    output = path.with_name('reference.xlsx')
    package.write(output)
    return output


@pytest.mark.parametrize('prefix', [False, True])
def test_explicit_mother_authoring_creates_native_tables_and_preserves_other_parts(tmp_path, prefix):
    base = mother_fixture(tmp_path / 'base.xlsx', prefix=prefix)
    reference = reference_fixture(tmp_path / 'ref.xlsx')
    inputs = {p: hashlib.sha256(p.read_bytes()).digest() for p in (base, reference)}
    before = Package(base)
    result = prepare_electrode_template(str(base), str(reference), str(tmp_path / 'prepared'))
    after = Package(result['output_path'])
    book = Workbook(after)
    cells = book.cell_map(book.sheets[0])
    assert book.value(cells['B106']) == '电阻率（Ω·cm）'
    assert book.value(cells['G146']) == '极片总厚度（μm）'
    assert book.value(cells['A118']) == '平均值' and book.value(cells['A119']) == 'COV（%）'
    assert book.value(cells['A147']) == '设定压强（MPa）'
    styles = after.doc('xl/styles.xml')
    for name in ('fonts', 'fills', 'borders', 'cellXfs'):
        original = before.doc('xl/styles.xml').root.child(name, S)
        collection = styles.root.child(name, S)
        assert int(collection.attrs['count']) == len(collection.children)
        assert all(n.ns == S for n in collection.children)
        assert [styles.raw(n) for n in collection.children[:len(original.children)]] == [before.doc('xl/styles.xml').raw(n) for n in original.children]
    xfs = styles.root.child('cellXfs', S).children
    fills = styles.root.child('fills', S).children
    for address, color in (('B107', 'FF00B050'), ('B118', 'FFFFFF00'), ('B147', 'FFB8CCE4')):
        fill = fills[int(xfs[int(cells[address].attrs['s'])].attrs['fillId'])]
        assert fill.child('patternFill', S).child('fgColor', S).attrs['rgb'] == color
    assert len([i for i in book.images() if i['name'].startswith('XMLMCP Reference ')]) == 6
    assert 'xl/charts/chart1.xml' not in after.parts
    for part in ('xl/comments/comment1.xml', 'xl/ctrlProps/ctrlProp1.xml', 'xl/drawings/vmlDrawing1.vml',
                 'xl/sharedStrings.xml', 'customXml/unknown.xml', 'customXml/opaque.bin', 'xl/media/image1.png'):
        assert before.parts[part] == after.parts[part]
    assert verify_integrity(str(base), result['output_path'], result['report_path']).valid
    assert all(hashlib.sha256(p.read_bytes()).digest() == sha for p, sha in inputs.items())
    with pytest.raises(XMLMCPError) as caught:
        prepare_electrode_template(str(base), str(reference), str(tmp_path / 'prepared'))
    assert caught.value.code == 'OUTPUT_EXISTS'


def test_missing_reference_chart_does_not_publish_partial_mother(tmp_path):
    base = mother_fixture(tmp_path / 'base.xlsx')
    reference = write_template(tmp_path / 'reference.xlsx', image=True)
    with pytest.raises(XMLMCPError) as caught:
        prepare_electrode_template(str(base), str(reference), str(tmp_path / 'blocked'))
    assert caught.value.code == 'TEMPLATE_REFERENCE'
    assert not (tmp_path / 'blocked').exists() and not list(tmp_path.glob('.xmlmcp-template-*'))


def test_reference_table_container_is_detected_from_bytes():
    class Images:
        parts = {'xl/media/table.emf': b'\xd7\xcd\xc6\x9a' + b'\0' * 50,
                 'xl/media/other.emf': b'\0' * 40 + b' EMF' + b'\0' * 10}
    result = reference_tables(Images())
    assert [r['container'] for r in result] == ['WMF', 'EMF']
    assert all(not r['editable_cells'] for r in result)
