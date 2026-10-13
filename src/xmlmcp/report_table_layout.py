"""Resize reviewed native table cells without changing fonts, row heights or columns."""
import json

from .cells import CellChange, apply_cell_changes
from .errors import require
from .namespaces import S
from .models import Target
from .workbook import Rectangle, cell_position
from .xmlbytes import Edit, append_children, set_attribute, xml_escape

LAYOUT_NS = 'urn:xmlmcp:report-table-layouts:1'


def embed_layouts(book, layouts):
    require(not read_layouts(book), 'TEMPLATE_LAYOUT', '母版已包含布局元数据，请使用原始空白母版')
    payload = json.dumps(layouts, separators=(',', ':'), ensure_ascii=False)
    part = book.package.fresh_part('customXml', 'xmlmcpTableLayouts', 'xml')
    book.package.add(part, b'<tableLayouts xmlns="' + LAYOUT_NS.encode() + b'">' + xml_escape(payload) + b'</tableLayouts>', 'cell:template_layout')
    book.package.add_content_type(part, 'application/xml')
    book.package.add_relationship(book.part, 'customXml', part)


def read_layouts(book):
    expressions = [book.package.doc(r['part']).root.text for r in book.package.related(book.part, 'customXml')
                   if r['part'] and r['part'].endswith('.xml') and book.package.doc(r['part']).root.ns == LAYOUT_NS]
    if not expressions:
        return {}
    require(len(expressions) == 1, 'TEMPLATE_LAYOUT', '表格布局元数据不唯一')
    try:
        result = json.loads(expressions[0])
        require(isinstance(result, dict), 'TEMPLATE_LAYOUT', '表格布局元数据不是对象')
        return result
    except (ValueError, UnicodeError) as exc:
        require(False, 'TEMPLATE_LAYOUT', '表格布局元数据无法读取', error=str(exc))


def table_style_map(rect, styles, single):
    capacity = (rect.size[1] - 1) // 2
    result = {}
    for address in rect.cells():
        row, col = cell_position(address)
        if row == rect.top:
            role = 'single_title' if single else 'pressure_title'
        elif row == rect.top + 1:
            role = 'single_header' if single else 'pressure_header'
        elif single and row == rect.bottom - 1:
            role = 'mean_index' if col == rect.left else 'mean_resistivity' if col <= rect.left + capacity else 'mean_thickness'
        elif single and row == rect.bottom:
            role = 'cov_label' if col == rect.left else 'cov'
        else:
            role = 'index' if col == rect.left else 'resistivity' if col <= rect.left + capacity else 'thickness'
        if role in ('index', 'resistivity', 'thickness'):
            role = ('single_' if single else 'pressure_') + role
        result[address] = str(styles[role])
    return result


def table_merges(rect):
    capacity = (rect.size[1] - 1) // 2
    return [r.address for r in (Rectangle(rect.top, rect.left + 1, rect.top, rect.left + capacity),
                                Rectangle(rect.top, rect.left + capacity + 1, rect.top, rect.right)) if r.size[1] > 1]


def apply_table_layouts(book, plans):
    validate_table_plans(book, plans)
    for plan in plans:
        sheet = book.sheet(plan['sheet'])
        region, active = Rectangle.parse(plan['region']), Rectangle.parse(plan['active'])
        require(region.contains(active.top, active.left) and region.contains(active.bottom, active.right),
                'REPORT_CAPACITY', '实际表格超出母版空白区域')
        styles = plan['cell_styles']
        changes = {a: CellChange(value=None, style=s, remove_style=s is None) for a, s in styles.items()}
        apply_cell_changes(book, {sheet['name']: changes})
        doc = book.sheet_doc(sheet)
        merges = doc.root.child('mergeCells', S)
        require(merges is not None, 'TEMPLATE_LAYOUT', '母版缺少合并集合')
        remove = []
        for node in merges.children:
            rect = Rectangle.parse(node.attrs['ref'])
            overlaps = not (rect.right < region.left or rect.left > region.right or rect.bottom < region.top or rect.top > region.bottom)
            if overlaps:
                require(region.contains(rect.top, rect.left) and region.contains(rect.bottom, rect.right),
                        'TEMPLATE_LAYOUT', '表格布局与区域外合并范围冲突')
                remove.append(node)
        additions = table_merges(active)
        tag = merges.qualify('mergeCell')
        payload = b''.join(b'<' + tag + b' ref="' + a.encode() + b'"/>' for a in additions)
        patches = [Edit(n.start, n.end, b'', 'cell:table_merge', n.attrs['ref']) for n in remove]
        patches.append(Edit(merges.start, merges.open_end, set_attribute(doc.data, merges, b'count', str(len(merges.children) - len(remove) + len(additions))), 'cell:table_merge_count', region.address))
        if payload:
            patches.append(append_children(doc.data, merges, payload, 'cell:table_merge', active.address))
        book.package.patch(sheet['part'], patches)


def validate_table_plans(book, plans):
    embedded = read_layouts(book)
    used = set()
    for plan in plans:
        kind = plan['table']
        require(kind not in used and kind in embedded, 'TEMPLATE_LAYOUT', '未知或重复的动态表格')
        used.add(kind)
        recipe = embedded[kind]['layout']
        sheet, region = book.resolve(Target.model_validate(recipe['region']))
        active = Rectangle.parse(plan['active'])
        require(plan['sheet'] == sheet['name'] and plan['region'] == region.address
                and active.top == region.top and active.left == region.left
                and active.bottom <= region.bottom and active.right <= region.right
                and active.size[1] >= 3 and active.size[1] % 2 == 1
                and active.size[0] >= (5 if kind.startswith('single') else 3),
                'TEMPLATE_LAYOUT', '动态表格范围与母版布局不一致')
        expected = dict(recipe['background_styles'])
        expected.update(table_style_map(active, recipe['styles'], kind.startswith('single')))
        require(plan['cell_styles'] == expected and plan['merges'] == table_merges(active),
                'REPORT_STYLE_CHANGED', '动态表格含非母版样式或额外单元格')


def plan_table_layouts(book, profile, samples, files):
    embedded = read_layouts(book)
    plans = []
    for kind, slot in profile.tables.items():
        if slot.layout is None:
            continue
        require(kind in ('single_measurements', 'pressure_measurements') and kind in embedded,
                'TEMPLATE_LAYOUT', '动态表格必须使用母版内已有的布局')
        from .report_models import MeasurementLayout
        require(slot.layout == MeasurementLayout.model_validate(embedded[kind]['layout']),
                'REPORT_STYLE_CHANGED', '已确认映射的表格样式与母版不一致')
        single = kind == 'single_measurements'
        mode = 'single_point' if single else 'pressure'
        key = 'point' if single else 'nominal_pressure_mpa'
        records = [files.get((s['sample_name'], mode)) for s in samples]
        require(all(records), 'RAW_MODE_MISSING', '动态表格所需记录缺失')
        points = [r[key] for r in records[0]['records']]
        require(points and all([r[key] for r in item['records']] == points for item in records),
                'REPORT_PRESSURE_GRID', '样品测量点或压强工步不同，不能按行合并')
        sheet, region = book.resolve(slot.layout.region)
        active = Rectangle(region.top, region.left, region.top + len(points) + (3 if single else 1), region.left + 2 * len(samples))
        require(active.bottom <= region.bottom and active.right <= region.right, 'REPORT_CAPACITY',
                '实际表格超出空白区域；不会改变全页行高列宽或覆盖参考资料', required=list(active.size), capacity=list(region.size))
        cell_styles = dict(slot.layout.background_styles)
        require(set(cell_styles) == set(region.cells()), 'TEMPLATE_LAYOUT', '表格背景样式清单不完整')
        cell_styles.update(table_style_map(active, slot.layout.styles, single))
        plans.append({'table': kind, 'sheet': sheet['name'], 'region': region.address, 'active': active.address,
                      'cell_styles': cell_styles, 'merges': table_merges(active)})
        slot.target = Target(sheet=sheet['name'], range=active.address)
    apply_table_layouts(book, plans)
    return plans
