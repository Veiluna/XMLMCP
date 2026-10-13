"""Explicit preparation of a reusable, editable electrode-report mother template.

This is a separate user-requested authoring step. Report generation never runs it
implicitly and never creates styles.
"""
import json
import tempfile
from pathlib import Path

from .errors import require
from .package import Package
from .workbook import Workbook, Rectangle, cell_position
from .namespaces import S, XDR, A, R
from .xmlbytes import Edit, append_children, set_attribute, xml_escape, digest
from .cells import CellChange, apply_cell_changes
from .report_preservation import preservation_contract, verify_preservation
from .report_table_fonts import table_fonts
from .report_table_layout import embed_layouts, table_style_map, table_merges
from .publication import publish_bundle
from .integrity import verify_packages
from .models import Target

REFERENCE_NAMES = {21:'single_resistance', 22:'single_resistivity', 31:'single_conductivity',
                   32:'single_thickness', 42:'pressure_resistivity', 43:'pressure_thickness'}


def reference_tables(package):
    """Identify embedded reference images without claiming editable cell styles."""
    result = []
    for part, data in package.parts.items():
        if not part.endswith(('.emf', '.wmf')):
            continue
        kind = ('WMF' if data[:4] == b'\xd7\xcd\xc6\x9a' else
                'EMF' if len(data) >= 44 and data[40:44] == b' EMF' else 'unknown')
        result.append({'part': part, 'container': kind, 'editable_cells': False,
                       'style_source': 'reference appearance; native cell styles authored separately'})
    return result


def _append_styles(package, part, name, fragments):
    import re
    fragments=[re.sub(br'^(<[A-Za-z0-9:]+)',lambda m:m[0]+b' xmlns="'+S.encode()+b'"',f,count=1) for f in fragments]
    doc = package.doc(part)
    node = doc.root.child(name, S)
    require(node is not None, 'TEMPLATE_STYLE', '母版缺少标准样式集合', collection=name)
    first = len(node.children)
    opening = set_attribute(doc.data, node, b'count', str(first + len(fragments)))
    if node.self_closing:
        replacement = opening.rstrip()[:-2] + b'>' + b''.join(fragments) + b'</' + node.raw_name + b'>'
        edits = [Edit(node.start, node.end, replacement, 'cell:template_styles', name)]
    else:
        edits = [Edit(node.start, node.open_end, opening, 'cell:template_style_count', name),
                 append_children(doc.data, node, b''.join(fragments), 'cell:template_styles', name)]
    package.patch(part, edits)
    return list(range(first, first + len(fragments)))


def _table_styles(book, fonts):
    part = book.package.related(book.part, 'styles')[0]['part']
    font_ids = {}
    for mode, roles in fonts.items():
        for role, font in roles.items():
            flags = ('<b/>' if font['bold'] else '') + ('<i/>' if font['italic'] else '') + ('<u/>' if font['underline'] else '')
            fragment = b'<font>' + flags.encode() + b'<sz val="' + str(font['size_pt']).encode() + b'"/><name val="' + xml_escape(font['name'], True) + b'"/><family val="2"/><charset val="134"/></font>'
            font_ids[mode + '_' + role] = _append_styles(book.package, part, 'fonts', [fragment])[0]
    fills = _append_styles(book.package, part, 'fills', [
        f'<fill><patternFill patternType="solid"><fgColor rgb="FF{color}"/><bgColor indexed="64"/></patternFill></fill>'.encode()
        for color in ('FFFFFF','00B050','FFFF00','B8CCE4')])
    border = _append_styles(book.package, part, 'borders', [
        ('<border>' + ''.join(f'<{edge} style="thin"><color rgb="FF808080"/></{edge}>' for edge in ('left','right','top','bottom'))
         + '<diagonal/></border>').encode()])[0]
    doc = book.package.doc(part)
    formats = doc.root.child('numFmts', S)
    require(formats is not None, 'TEMPLATE_STYLE', '母版缺少 numFmts 集合')
    next_id = max([163] + [int(n.attrs['numFmtId']) for n in formats.children]) + 1
    codes = ('0.0000','0.0','0.00"%"')
    fragments = [b'<numFmt numFmtId="' + str(next_id+i).encode() + b'" formatCode="' + xml_escape(code, True) + b'"/>' for i,code in enumerate(codes)]
    _append_styles(book.package, part, 'numFmts', fragments)
    definitions = {}
    for mode in ('single', 'pressure'):
        definitions.update({mode+'_title': (fills[0], 0, font_ids[mode+'_title']),
                            mode+'_header': (fills[1 if mode=='single' else 3], 0, font_ids[mode+'_header']),
                            mode+'_index': (fills[0], 0, font_ids[mode+'_body']),
                            mode+'_resistivity': (fills[0], next_id, font_ids[mode+'_body']),
                            mode+'_thickness': (fills[0], next_id+1, font_ids[mode+'_body'])})
    definitions.update({'mean_index':(fills[2],0,font_ids['single_body']), 'mean_resistivity':(fills[2],next_id,font_ids['single_body']),
                        'mean_thickness':(fills[2],next_id+1,font_ids['single_body']), 'cov':(fills[0],next_id+2,font_ids['single_body']),
                        'cov_label':(fills[0],0,font_ids.get('single_cov_label',font_ids['single_body']))})
    indexes = _append_styles(book.package, part, 'cellXfs', [
        f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'.encode()
        for fill,fmt,font in definitions.values()])
    return dict(zip(definitions, indexes))


def _prepare_table(book, sheet, rect, styles, single):
    cells = book.cell_map(sheet)
    require(all(book.value(cells.get(a)) in (None,'',' ') for a in rect.cells()),
            'TEMPLATE_TABLE_OVERLAP', '新原生表格区域包含已有数据，请选择空白母版')
    changes = {}
    capacity = (rect.size[1]-1)//2
    style_map = table_style_map(rect, styles, single)
    for address in rect.cells():
        row,col = cell_position(address)
        values = {(rect.top,rect.left+1):'电阻率（Ω·cm）', (rect.top,rect.left+capacity+1):'极片总厚度（μm）',
                  (rect.top+1,rect.left):'编号' if single else '设定压强（MPa）'}
        if single:
            values[(rect.bottom-1,rect.left)]='平均值'
            values[(rect.bottom,rect.left)]='COV（%）'
        change=CellChange(style=style_map[address])
        if (row,col) in values:change.value=values[(row,col)]
        changes[address] = change
    apply_cell_changes(book, {sheet['name']:changes})
    doc = book.sheet_doc(sheet)
    merges = doc.root.child('mergeCells', S)
    require(merges is not None, 'TEMPLATE_LAYOUT', '母版缺少合并区域集合')
    ranges = table_merges(rect)
    existing=[Rectangle.parse(n.attrs['ref']) for n in merges.children]
    require(not any(not (r.right<rect.left or r.left>rect.right or r.bottom<rect.top or r.top>rect.bottom) for r in existing),
            'TEMPLATE_LAYOUT', '新表格区域与已有合并范围冲突')
    tag = merges.qualify('mergeCell')
    payload = b''.join(b'<'+tag+b' ref="'+r.encode()+b'"/>' for r in ranges)
    book.package.patch(sheet['part'], [
        Edit(merges.start, merges.open_end, set_attribute(doc.data, merges, b'count', str(len(merges.children)+len(ranges))), 'cell:template_merge_count', rect.address),
        append_children(doc.data, merges, payload, 'cell:template_merge', rect.address)])


def prepare_electrode_template(base_path, reference_path, output_directory, *, sample_count=5, single_points=10, pressure_points=12):
    destination = Path(output_directory)
    require(destination.is_absolute(), 'ABSOLUTE_PATH_REQUIRED', '输出目录必须使用绝对路径')
    destination=destination.resolve()
    require(destination.parent.is_dir() and not destination.exists(), 'OUTPUT_EXISTS', '输出目录必须不存在，父目录必须存在')
    package=Package(base_path);package.validate()
    reference=Package(reference_path);reference.validate()
    book,refbook=Workbook(package),Workbook(reference)
    sheet,single=book.resolve(Target(name='单点结果区',scope='workbook'))
    pressure_sheet,pressure=book.resolve(Target(name='变压结果区',scope='workbook'))
    require(single.size==(52,11) and pressure.size==(35,11) and sheet['name']==pressure_sheet['name'],
            'TEMPLATE_LAYOUT', '此准备器只适配已标记极片母版的标准布局')
    contract = preservation_contract(book, authoring=True)
    require(isinstance(sample_count,int) and 1<=sample_count<=5 and isinstance(single_points,int) and 1<=single_points<=16
            and isinstance(pressure_points,int) and 1<=pressure_points<=16, 'REPORT_CAPACITY', '母版空白区最多容纳五个样品、十六个测量点或压强工步')
    fonts={'single':table_fonts(refbook,6,True), 'pressure':table_fonts(refbook,11,False)}
    styles=_table_styles(book, fonts)
    tables={'single_measurements':Rectangle(single.top+32,1,single.top+35+single_points,1+2*sample_count),
            'pressure_measurements':Rectangle(pressure.top+17,1,pressure.top+18+pressure_points,1+2*sample_count)}
    regions={'single_measurements':Rectangle(single.top+32,1,single.bottom,11),
             'pressure_measurements':Rectangle(pressure.top+17,1,pressure.bottom,11)}
    layouts={}
    original_cells=book.cell_map(sheet)
    for kind,rect in regions.items():
        require(all(book.value(original_cells.get(a)) in (None,'',' ') for a in rect.cells()), 'TEMPLATE_TABLE_OVERLAP', '表格扩展区域包含原有数据')
        layouts[kind]={'target':{'sheet':sheet['name'],'range':tables[kind].address},
                       'layout':{'region':{'sheet':sheet['name'],'range':rect.address},
                                 'styles':{k:str(v) for k,v in styles.items()},
                                 'background_styles':{a:original_cells[a].attrs.get('s') if a in original_cells else None for a in rect.cells()}}}
    for kind,rect in tables.items():_prepare_table(book,sheet,rect,styles,kind.startswith('single'))
    embed_layouts(book, layouts)
    drawings=package.related(sheet['part'],'drawing')
    require(len(drawings)==1,'TEMPLATE_LAYOUT','母版必须有唯一 drawing')
    drawing=drawings[0]['part']
    refimages=refbook.images()
    for shape_id,kind in REFERENCE_NAMES.items():
        matches=[i for i in refimages if i['shape_id']==str(shape_id)]
        require(len(matches)==1 and matches[0]['media_part'].endswith('.png'), 'TEMPLATE_REFERENCE', '缺少预期的单点/变压 PNG 示例',kind=kind)
        original=matches[0];doc=original['_doc'];anchor=original['_pic'].parent
        raw=doc.raw(anchor)
        media=package.fresh_part('xl/media','xmlmcpReference','png')
        package.add(media,reference.parts[original['media_part']],'image:template_reference')
        package.add_content_type(media,'image/png','png')
        rid=package.add_relationship(drawing,'image',media)
        from .xmlbytes import XMLDocument
        from .xmlbytes import apply_edits, standalone
        copied_raw=standalone(doc,anchor)
        local=XMLDocument(copied_raw);copied=local.root
        nv=next(copied.descendants('cNvPr',XDR));blip=next(copied.descendants('blip',A))
        raw=apply_edits(local.data,[Edit(nv.start,nv.open_end,set_attribute(local.data,nv,b'name','XMLMCP Reference '+kind),'image:template_name',kind),
                                    Edit(blip.start,blip.open_end,set_attribute(local.data,blip,b'r:embed',rid),'image:template_media',kind)])[0]
        target_doc=package.doc(drawing)
        require(not any(n.attrs.get('id')==str(shape_id) for n in target_doc.root.descendants('cNvPr',XDR)),
                'TEMPLATE_REFERENCE','图形 ID 冲突，请选择原始空白母版')
        package.patch(drawing,[append_children(target_doc.data,target_doc.root,raw,'image:template_anchor',kind)])
    package.validate()
    with tempfile.TemporaryDirectory(prefix='.xmlmcp-template-',dir=destination.parent) as temporary:
        bundle=Path(temporary)/'bundle';bundle.mkdir()
        staging=bundle/'template.xlsx';package.write(staging)
        report=package.manifest()
        audit={'base_sha256':package.fingerprint,'reference_sha256':reference.fingerprint,
               'removed_examples':{'removed_objects':[],'removed_parts':[]},'tables':{k:r.address for k,r in tables.items()},
               'table_style_ids':styles,'table_fonts':fonts,'table_layouts':layouts,'embedded_table_sources':reference_tables(reference),
               'graph_style':'template-statistics-v1','graph_style_status':'reconstructed from reference PNG; no pixel identity claim',
               'purpose':'Explicit template authoring; subsequent reports preserve these destination styles'}
        report.update({'output_fingerprint':digest(staging.read_bytes()),'output_path':str(destination/'template.xlsx'),'preparation':audit})
        report['reference_preservation'] = contract
        audit['preservation'] = verify_preservation(Package(base_path), Package(staging), contract)
        verified=verify_packages(Package(base_path),Package(staging),report)
        require(verified.valid,'INTEGRITY_FAILED','新母版未通过字节验证',errors=verified.errors)
        report['verification']=verified.model_dump()
        (bundle/'template.integrity.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        (bundle/'preparation.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
        require(digest(package.path.read_bytes())==package.fingerprint and digest(reference.path.read_bytes())==reference.fingerprint,
                'INPUT_CHANGED','制作母版过程中来源改变，取消输出')
        publish_bundle(bundle,destination)
    return {'output_path':str(destination/'template.xlsx'),'report_path':str(destination/'template.integrity.json'),'verification':verified.model_dump(),'preparation':audit}
