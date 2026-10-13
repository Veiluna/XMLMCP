"""Protect the supplied report's reference information and existing objects."""
from .errors import require
from .namespaces import S, P
from .package import rels_path
from .report_objects import drawing_objects
from .workbook import Workbook, cell_position


def preservation_contract(book, *, authoring=False):
    rows = {}
    for sheet in book.sheets:
        if sheet['kind'] != 'worksheet':
            continue
        headings = [cell_position(a)[0] for a, c in book.cell_map(sheet).items()
                    if book.value(c) == '极片电阻仪参考资料']
        if headings:
            rows[sheet['name']] = min(headings)
    objects = [o['object_id'] for o in drawing_objects(book)
               if authoring or (o['sheet'] in rows and not o['name'].startswith('XMLMCP Reference '))]
    return {'reference_start_rows': rows, 'object_ids': objects}


def verify_preservation(source, output, contract):
    before, after = Workbook(source), Workbook(output)
    checked_rows = 0
    for name, first in contract['reference_start_rows'].items():
        old_doc, new_doc = before.sheet_doc(before.sheet(name)), after.sheet_doc(after.sheet(name))
        def rows(doc):
            return {n.attrs['r']: doc.raw(n) for n in doc.root.descendants('row', S) if int(n.attrs['r']) >= first}
        require(rows(old_doc) == rows(new_doc), 'REPORT_REFERENCE_CHANGED', '参考资料的单元格或行格式发生改变', sheet=name)
        checked_rows += len(rows(old_doc))
        def merges(doc):
            return [doc.raw(n) for n in doc.root.descendants('mergeCell', S) if cell_position(n.attrs['ref'].split(':')[0])[0] >= first]
        require(merges(old_doc) == merges(new_doc), 'REPORT_REFERENCE_CHANGED', '参考资料的合并单元格发生改变')
    old_objects = {o['object_id']: o for o in drawing_objects(before)}
    new_objects = {o['object_id']: o for o in drawing_objects(after)}
    dependencies = set()
    for identifier in contract['object_ids']:
        require(identifier in old_objects and identifier in new_objects, 'REPORT_REFERENCE_CHANGED', '原有图片或图表被移除', object_id=identifier)
        old, new = old_objects[identifier], new_objects[identifier]
        require(source.doc(old['drawing_part']).raw(old['_anchor']) == output.doc(new['drawing_part']).raw(new['_anchor']),
                'REPORT_REFERENCE_CHANGED', '原有图片或图表的锚点、样式或关系被改变', object_id=identifier)
        for node in old['_anchor'].descendants():
            for attr, rid in node.attrs.items():
                if attr.rsplit('}', 1)[-1] in ('embed', 'link', 'id') and 'relationships}' in attr:
                    path = rels_path(old['drawing_part'])
                    old_doc, new_doc = source.doc(path), output.doc(path)
                    def relation(doc):
                        return [doc.raw(n) for n in doc.root.children if n.ns == P and n.attrs.get('Id') == rid]
                    require(relation(old_doc) == relation(new_doc), 'REPORT_REFERENCE_CHANGED', '参考对象关系改变')
        if old['part']:
            pending = [old['part']]
            while pending:
                part = pending.pop()
                if part in dependencies:
                    continue
                dependencies.add(part)
                path = rels_path(part)
                if path in source.parts:
                    dependencies.add(path)
                pending.extend(r['part'] for r in source.relationships(part) if r['part'])
    for part in dependencies:
        require(source.parts[part] == output.parts.get(part), 'REPORT_REFERENCE_CHANGED', '参考对象的依赖部件改变', part=part)
    return {'reference_information_identical': True, 'preserved_reference_rows': checked_rows,
            'preserved_original_objects': len(contract['object_ids']), 'preserved_object_parts': len(dependencies)}
