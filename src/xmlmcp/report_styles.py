"""Verify a report keeps its destination formatting, including inherited styles."""
from .errors import require
from .namespaces import S
from .workbook import Workbook, semantic_node


def verify_report_styles(source, output) -> dict:
    before, after = Workbook(source), Workbook(output)
    style_parts={r['part'] for kind in ('styles','theme') for r in source.related(before.part,kind)}
    for part, data in source.parts.items():
        if part in style_parts or part.startswith('xl/theme/'):
            require(output.parts.get(part) == data, 'REPORT_STYLE_CHANGED', '报告改变了模板样式或主题部件', part=part)
    checked = 0
    for sheet in before.sheets:
        if sheet['kind'] != 'worksheet':
            continue
        old_doc, new_doc = before.sheet_doc(sheet), after.sheet_doc(after.sheet(sheet['name']))
        for local in ('cols', 'sheetFormatPr', 'mergeCells', 'conditionalFormatting'):
            old_nodes = [old_doc.raw(n) for n in old_doc.root.children if n.ns == S and n.local == local]
            new_nodes = [new_doc.raw(n) for n in new_doc.root.children if n.ns == S and n.local == local]
            require(old_nodes == new_nodes, 'REPORT_STYLE_CHANGED', '报告改变了列样式、尺寸、合并或条件格式', sheet=sheet['name'], element=local)
        old_rows = {n.attrs['r']: n.attrs for n in old_doc.root.descendants('row', S)}
        new_rows = {n.attrs['r']: n.attrs for n in new_doc.root.descendants('row', S)}
        require(all(new_rows.get(r) == attrs for r, attrs in old_rows.items()),
                'REPORT_STYLE_CHANGED', '报告改变了原行的样式或尺寸', sheet=sheet['name'])
        old_cells, new_cells = before.cell_map(sheet), after.cell_map(after.sheet(sheet['name']))
        for address, old_cell in old_cells.items():
            new_cell = new_cells.get(address)
            require(new_cell is not None and old_cell.attrs.get('s') == new_cell.attrs.get('s'),
                    'REPORT_STYLE_CHANGED', '报告改变了原单元格样式', sheet=sheet['name'], cell=address)
            old_string, new_string = before.string_node(old_cell), after.string_node(new_cell)
            def properties(node):
                return [repr(semantic_node(r)) for r in node.descendants('rPr', S)] if node else []
            old_properties, new_properties = properties(old_string), properties(new_string)
            # Uniform rich text may consolidate runs while keeping the same character properties.
            if len(set(old_properties)) == 1 and old_properties:
                require(new_properties and set(old_properties) == set(new_properties),
                        'REPORT_STYLE_CHANGED', '报告改变了原字符字体格式', sheet=sheet['name'], cell=address)
            else:
                require(old_properties == new_properties, 'REPORT_STYLE_CHANGED', '报告改变了原富文本格式', sheet=sheet['name'], cell=address)
            checked += 1
        require(all('s' not in c.attrs for a, c in new_cells.items() if a not in old_cells),
                'REPORT_STYLE_CHANGED', '报告为新单元格指定了样式，可能覆盖原继承格式', sheet=sheet['name'])
    return {'destination_styles_identical': True, 'rich_text_styles_identical': True,
            'style_parts_identical': True, 'layout_styles_identical': True, 'original_cells_checked': checked}
