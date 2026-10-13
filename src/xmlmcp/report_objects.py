"""Remove only explicitly mapped example drawing objects and exclusive dependencies."""
from collections import defaultdict

from .errors import require
from .namespaces import XDR, C, A, R, P
from .package import rels_path, relationship_source
from .workbook import cell_address
from .xmlbytes import Edit


def drawing_objects(book) -> list[dict]:
    objects = []
    for sheet, part, doc, rels in book.drawings():
        for anchor in doc.root.children:
            nv = list(anchor.descendants('cNvPr', XDR))
            if len(nv) != 1:
                continue
            pic = anchor.child('pic', XDR)
            frame = anchor.child('graphicFrame', XDR)
            chart = next(frame.descendants('chart', C), None) if frame else None
            if pic is None and chart is None:
                continue
            kind = 'image' if pic is not None else 'chart'
            node = next(pic.descendants('blip', A), None) if pic is not None else chart
            rid = node.attrs.get(R + ('}embed' if kind == 'image' else '}id')) if node else None
            relation = rels.get(rid)
            marker = anchor.child('from', XDR)
            start = None
            if marker and marker.child('row', XDR) and marker.child('col', XDR):
                start = cell_address(int(marker.child('row', XDR).text) + 1, int(marker.child('col', XDR).text) + 1)
            objects.append({'object_id': f"{sheet['name']}|{part}|{nv[0].attrs['id']}",
                            'sheet': sheet['name'], 'drawing_part': part, 'shape_id': nv[0].attrs['id'],
                            'name': nv[0].attrs.get('name', ''), 'kind': kind, 'start_cell': start,
                            'part': relation['part'] if relation else None, '_anchor': anchor, '_rid': rid})
    return objects


def apply_example_removals(book, selections) -> dict:
    package = book.package
    inventory = drawing_objects(book)
    selected = []
    ids = set()
    for selection in selections:
        value = selection.model_dump() if hasattr(selection, 'model_dump') else selection
        require(value['object_id'] not in ids, 'REPORT_OBJECT_MAPPING', '示例对象重复指定')
        ids.add(value['object_id'])
        matches = [o for o in inventory if o['object_id'] == value['object_id']]
        require(len(matches) == 1 and matches[0]['kind'] == value['kind'],
                'REPORT_OBJECT_MAPPING', '示例对象不存在、不唯一或类型改变，请重新检查映射', object_id=value['object_id'])
        item = matches[0]
        anchor = item['_anchor']
        require(anchor.ns == XDR and anchor.local in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor')
                and item['part'] is not None, 'REPORT_OBJECT_UNSUPPORTED', '不支持删除分组、外链或无明确关系的示例对象')
        selected.append(item)
    edits = defaultdict(list)
    relation_ids = defaultdict(set)
    for item in selected:
        anchor = item['_anchor']
        edits[item['drawing_part']].append(Edit(anchor.start, anchor.end, b'', 'image:remove_example_anchor', item['object_id']))
        # A chart can also have hyperlinks or other explicit relationship attributes.
        for node in (anchor, *anchor.descendants()):
            for attr, rid in node.attrs.items():
                if attr in (R + '}id', R + '}embed', R + '}link'):
                    relation_ids[item['drawing_part']].add(rid)
    for part, patches in edits.items():
        package.patch(part, patches)
    removed = remove_unused_relations(book, relation_ids)
    return {'removed_objects': [{k: v for k, v in o.items() if not k.startswith('_')} for o in selected],
            'removed_parts': removed}


def remove_unused_relations(book, relation_ids) -> list[str]:
    package = book.package
    roots = set()
    for source, ids in relation_ids.items():
        # Keep a relationship if another remaining object uses it (shared pictures/charts).
        referenced = {v for node in package.doc(source).root.descendants() for a, v in node.attrs.items()
                      if a in (R + '}id', R + '}embed', R + '}link')}
        path = rels_path(source)
        doc = package.doc(path)
        rels = {r['Id']: r for r in package.relationships(source)}
        patches = []
        for node in doc.root.children:
            rid = node.attrs.get('Id')
            if node.name == P + '}Relationship' and rid in ids and rid not in referenced:
                target = rels[rid]['part']
                if target:
                    roots.add(target)
                patches.append(Edit(node.start, node.end, b'', 'image:remove_example_relationship', rid))
        package.patch(path, patches)
    # Delete only the selected objects' dependency closure without any incoming reference
    # from outside that closure. Shared media/styles/embedded workbooks remain untouched.
    candidates = set()
    pending = list(roots)
    while pending:
        part = pending.pop()
        if part in candidates:
            continue
        candidates.add(part)
        pending.extend(r['part'] for r in package.relationships(part) if r['part'] is not None)
    incoming = defaultdict(set)
    for path in package.parts:
        if path.endswith('.rels'):
            source = relationship_source(path)
            for relation in package.relationships(source):
                if relation['part'] is not None:
                    incoming[relation['part']].add(source)
    removable = set(candidates)
    while True:
        shared = {p for p in removable if incoming[p] - removable}
        if not shared:
            break
        removable -= shared
    removed = set()
    for part in sorted(removable):
        for path in (rels_path(part), part):
            if path in package.parts:
                package.remove(path, 'image:example_dependency')
                removed.add(path)
    if removed:
        types = package.doc('[Content_Types].xml')
        package.patch('[Content_Types].xml', [
            Edit(n.start, n.end, b'', 'image:remove_example_content_type', n.attrs['PartName'])
            for n in types.root.children if n.local == 'Override' and n.attrs.get('PartName', '').lstrip('/') in removed])
    return sorted(removed)
