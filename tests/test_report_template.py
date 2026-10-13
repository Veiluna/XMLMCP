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


def mother_fixture(path, *, prefix=False, chart=True):
    write_template(path, image=True, chart=chart, prefix=prefix, merged=True)
    package = Package(path)
    names = ('<definedName name="单点结果区">\'报告\'!$A$74:$K$125</definedName>'
             '<definedName name="变压结果区">\'报告\'!$A$129:$K$163</definedName>').encode()
    package.parts['xl/workbook.xml'] = package.parts['xl/workbook.xml'].replace(b'</definedNames>', names + b'</definedNames>')
    style = package.parts['xl/styles.xml'].replace(b'<fonts ', b'<numFmts count="0"/><fonts ', 1)
    if prefix:
        style = re.sub(br'<(/?)([A-Za-z]+)', br'<\1s:\2', style).replace(
            b'xmlns="' + S.encode() + b'"', b'xmlns:s="' + S.encode() + b'"')
    package.parts['xl/styles.xml'] = style
    # Put a local chart in the appendix; authoring must now preserve it.
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
    assert before.parts['xl/charts/chart1.xml'] == after.parts['xl/charts/chart1.xml']
    assert before.parts['xl/workbook.xml'] == after.parts['xl/workbook.xml']
    from xmlmcp.report_table_layout import read_layouts
    assert set(read_layouts(book))=={'single_measurements','pressure_measurements'}
    assert after.related(book.part,'customXml')
    assert result['preparation']['preservation']['preserved_original_objects'] == 3
    assert not result['preparation']['removed_examples']['removed_objects']
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


def test_dynamic_table_layout_shrinks_and_expands_without_changing_dimensions(tmp_path):
    from xmlmcp.report_models import ReportProfile
    from xmlmcp.report_table_layout import read_layouts, plan_table_layouts
    from xmlmcp.report_styles import verify_report_styles
    base = mother_fixture(tmp_path / 'base.xlsx')
    result = prepare_electrode_template(str(base), str(reference_fixture(tmp_path / 'ref.xlsx')), str(tmp_path / 'prepared'))
    source = Package(result['output_path'])
    book = Workbook(Package(result['output_path']))
    profile = ReportProfile(profile_id='dynamic', template_fingerprint=source.fingerprint,
                            order={'fields':{}, 'sample_rows':{}}, fields={}, tables=read_layouts(book))
    samples = [{'sample_name': s} for s in ('A','B')]
    files = {(s['sample_name'],m): {'records':[{k:p} for p in (1,2)]}
             for s in samples for m,k in (('single_point','point'),('pressure','nominal_pressure_mpa'))}
    plans = plan_table_layouts(book, profile, samples, files)
    assert [p['active'] for p in plans] == ['A106:E111', 'A146:E149']
    cells = book.cell_map(book.sheets[0])
    assert cells['F107'].attrs.get('s') is None and book.value(cells['A118']) is None
    proof = verify_report_styles(source, book.package, plans)
    assert proof['style_parts_identical'] and proof['outside_table_styles_identical']
    assert source.parts['xl/styles.xml'] == book.package.parts['xl/styles.xml']
    # Reuse the same mother and reviewed profile with more samples and more points.
    expanded = Workbook(Package(result['output_path']))
    profile = ReportProfile.model_validate(profile.model_dump())
    samples = [{'sample_name':str(i)} for i in range(5)]
    files = {(s['sample_name'],m):{'records':[{k:p} for p in range(16)]}
             for s in samples for m,k in (('single_point','point'),('pressure','nominal_pressure_mpa'))}
    plans = plan_table_layouts(expanded, profile, samples, files)
    assert [p['active'] for p in plans] == ['A106:K125','A146:K163']
    assert verify_report_styles(source, expanded.package, plans)['layout_styles_identical']


def test_dynamic_layout_rejects_modified_font_style_and_overflow(tmp_path):
    from xmlmcp.report_models import ReportProfile
    from xmlmcp.report_table_layout import read_layouts, plan_table_layouts
    base = mother_fixture(tmp_path / 'base.xlsx')
    result = prepare_electrode_template(str(base), str(reference_fixture(tmp_path / 'ref.xlsx')), str(tmp_path / 'prepared'))
    book = Workbook(Package(result['output_path']))
    profile = ReportProfile(profile_id='dynamic', template_fingerprint=book.package.fingerprint,
                            order={'fields':{},'sample_rows':{}}, fields={}, tables=read_layouts(book))
    profile.tables['single_measurements'].layout.styles['single_header']='0'
    files={('A','single_point'):{'records':[{'point':1}]}}
    with pytest.raises(XMLMCPError) as error:
        plan_table_layouts(book, profile, [{'sample_name':'A'}], files)
    assert error.value.code=='REPORT_STYLE_CHANGED'
    profile.tables['single_measurements'].layout.styles['single_header']=read_layouts(book)['single_measurements']['layout']['styles']['single_header']
    files[('A','single_point')]['records']=[{'point':i} for i in range(17)]
    with pytest.raises(XMLMCPError) as error:
        plan_table_layouts(book, profile, [{'sample_name':'A'}], files)
    assert error.value.code=='REPORT_CAPACITY'


def test_reference_guard_detects_image_deletion_and_cell_edit(tmp_path):
    from xmlmcp.cells import apply_cell_changes, CellChange
    from xmlmcp.report_preservation import preservation_contract, verify_preservation
    from xmlmcp.report_objects import apply_example_removals
    source_path=write_template(tmp_path/'reference.xlsx', image=True, chart=True)
    book=Workbook(Package(source_path))
    apply_cell_changes(book, {'报告':{'A167':CellChange(value='极片电阻仪参考资料'), 'A168':CellChange(value='原始说明')}})
    source_path=tmp_path/'with-reference.xlsx';book.package.write(source_path)
    source=Package(source_path)
    contract=preservation_contract(Workbook(source))
    assert contract['reference_start_rows']=={'报告':167}
    for edit in ('cell','image'):
        changed=Workbook(Package(source_path))
        if edit=='cell':
            apply_cell_changes(changed, {'报告':{'A168':CellChange(value='错误修改')}})
        else:
            apply_example_removals(changed,[{'object_id':contract['object_ids'][0], 'kind':'image'}])
        with pytest.raises(XMLMCPError) as error:
            verify_preservation(source, changed.package, contract)
        assert error.value.code=='REPORT_REFERENCE_CHANGED'


def test_wmf_selected_fonts_keep_weight_size_and_object_table():
    import struct
    from xmlmcp.report_table_fonts import wmf_text_fonts
    def record(function, payload):
        if len(payload)%2:
            payload+=b'\0'
        return struct.pack('<IH',(len(payload)+6)//2,function)+payload
    def font(height,weight):
        return struct.pack('<5h8B',height,0,0,0,weight,0,0,0,134,0,0,0,0)+'宋体'.encode('gb18030')+b'\0'
    def text(y,value):
        encoded=value.encode('gb18030')
        return struct.pack('<hhHH',y,10,len(encoded),0)+encoded
    data=struct.pack('<IH4hHIH',0x9ac6cdd7,0,0,0,900,450,96,0,0)+b'\0'*18
    data+=record(0x2fa,b'\0'*10) # Pen occupies slot zero.
    data+=record(0x2fb,font(-23,700))+record(0x12d,struct.pack('<H',1))
    data+=record(0xa32,text(27,'标题'))+record(0x1f0,struct.pack('<H',0))
    data+=record(0x2fb,font(-18,0))+record(0x12d,struct.pack('<H',0))+record(0xa32,text(90,'数值'))
    data+=record(0,b'')
    height,texts=wmf_text_fonts(data)
    assert height==450 and [t['font']['weight'] for t in texts]==[700,0]
    assert texts[0]['font']['height']==-23 and texts[1]['font']['face']=='宋体'
    with pytest.raises(XMLMCPError):
        wmf_text_fonts(data[:-3])


def test_reference_font_roles_ignore_symbol_baseline_offset(monkeypatch):
    from types import SimpleNamespace
    from xmlmcp import report_table_fonts
    from xmlmcp.namespaces import A
    from xmlmcp.xmlbytes import XMLDocument
    def font(name,weight=0):
        return {'face':name,'height':-18,'weight':weight,'italic':False,'underline':False}
    records=[{'y':24,'text':'Ω','font':font('Calibri')},
             {'y':25,'text':'电阻率','font':font('宋体')},
             {'y':58,'text':'编号','font':font('宋体',700)},
             {'y':91,'text':'1.23','font':font('宋体')},
             {'y':454,'text':'COV','font':font('Calibri')}]
    monkeypatch.setattr(report_table_fonts,'wmf_text_fonts',lambda data:(450,records))
    doc=XMLDocument(f'<xdr:pic xmlns:xdr="{XDR}" xmlns:a="{A}"><xdr:spPr><a:xfrm><a:ext cy="2857500" cx="1"/></a:xfrm></xdr:spPr></xdr:pic>'.encode())
    book=SimpleNamespace(package=SimpleNamespace(parts={'table.emf':b'test'}),
                         images=lambda:[{'shape_id':'6','media_part':'table.emf','_pic':doc.root}])
    fonts=report_table_fonts.table_fonts(book,6,True)
    assert fonts['title']['name']=='宋体' and fonts['title']['size_pt']==9
    assert fonts['header']['bold'] and not fonts['body']['bold']
    assert fonts['cov_label']['name']=='Calibri'
