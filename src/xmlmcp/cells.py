from __future__ import annotations

from dataclasses import dataclass

from .errors import require
from .models import WriteValue, WriteFormula, ClearValue, CopyStyle, Operation
from .namespaces import S
from .strings import encode_xstring
from .workbook import Workbook, Rectangle, cell_position, cell_address
from .xmlbytes import XMLDocument, Edit, apply_edits, append_children, set_attribute, xml_escape, edit_attributes

UNSET = object()


@dataclass
class CellChange:
    value: object = UNSET
    formula: str | None = None
    style: str | None = None
    cached: bool = False
    cache_type: str | None = None
    preserve_formula: bool = False


def preflight_cells(book: Workbook, operations: list[Operation]) -> dict[str, dict[str, CellChange]]:
    changes: dict[str, dict[str, CellChange]] = {}
    maps = {s["name"]: book.cell_map(s) for s in book.sheets if s["kind"] == "worksheet"}
    for op in operations:
        if not isinstance(op, (WriteValue, WriteFormula, ClearValue, CopyStyle)):
            continue
        sheet, rect = book.resolve(op.target)
        scalar = isinstance(op, WriteFormula) or isinstance(op, WriteValue) and not isinstance(op.value, list)
        rect = book.merged_anchor(sheet, rect, scalar)
        matrix = None
        if isinstance(op, WriteValue) and isinstance(op.value, list):
            matrix = op.value
            require(matrix and matrix[0] and all(len(row) == len(matrix[0]) for row in matrix),
                    "INVALID_MATRIX", "矩阵必须非空且各行等长")
            require(len(matrix) <= rect.size[0] and len(matrix[0]) <= rect.size[1], "MATRIX_OUT_OF_RANGE", "矩阵超出目标预留范围")
            rect = Rectangle(rect.top, rect.left, rect.top + len(matrix) - 1, rect.left + len(matrix[0]) - 1)
            book.merged_anchor(sheet, rect)
        if scalar:
            require(rect.size == (1, 1), "SCALAR_TARGET", "标量和普通公式必须写入单格或完整合并单元格")
        style = None
        if isinstance(op, CopyStyle):
            source_sheet, source_rect = book.resolve(op.source)
            source_rect = book.merged_anchor(source_sheet, source_rect, True)
            require(source_rect.size == (1, 1), "STYLE_SOURCE", "样式来源必须是单格")
            source = maps[source_sheet["name"]].get(source_rect.address)
            book.style_info(source)
            style = source.attrs.get("s", "0") if source else "0"
        for address in rect.cells():
            cell = maps[sheet["name"]].get(address)
            for other in maps[sheet["name"]].values():
                f = other.child("f", S)
                if f and f.attrs.get("t") in ("shared", "array", "dataTable"):
                    protected = Rectangle.parse(f.attrs["ref"]) if "ref" in f.attrs else Rectangle.parse(other.attrs["r"])
                    row, col = cell_position(address)
                    require(not protected.contains(row, col), "SPECIAL_FORMULA_TARGET", "不支持修改共享、数组或数据表公式区域", cell=address)
            entry = changes.setdefault(sheet["name"], {}).setdefault(address, CellChange())
            if isinstance(op, CopyStyle):
                require(entry.style is None, "TARGET_CONFLICT", "同一单元格不能重复复制样式", cell=address)
                entry.style = style
                continue
            require(entry.value is UNSET and entry.formula is None, "TARGET_CONFLICT", "同一单元格不能重复写入", cell=address)
            rpr = book.uniform_rpr(cell)
            if isinstance(op, WriteFormula):
                formula = op.formula[1:] if op.formula.startswith("=") else op.formula
                require(formula.strip() and not formula.startswith("{"), "INVALID_FORMULA", "需提供普通公式表达式")
                require(not rpr, "RICH_TEXT_TYPE_CHANGE", "字符格式只存于富文本的单元格不能改为公式；请使用已格式化的数值单元格")
                entry.formula = formula
            elif isinstance(op, ClearValue):
                require(not rpr, "RICH_TEXT_TYPE_CHANGE", "该字体只存于富文本；清为空值会丢失字体，已取消操作")
                entry.value = None
            else:
                row, col = cell_position(address)
                value = matrix[row-rect.top][col-rect.left] if matrix else op.value
                require(not rpr or isinstance(value, str), "RICH_TEXT_TYPE_CHANGE",
                        "字符格式只存于富文本的单元格不能改为数值或布尔值；请使用已格式化的数值单元格")
                entry.value = value
    return changes


def value_fragment(cell_node, value, formula: str | None, rpr: bytes | None, kind: str | None = None) -> tuple[str | None, bytes]:
    qualify = cell_node.qualify
    def element(name, content):
        tag = qualify(name)
        return b"<" + tag + b">" + content + b"</" + tag + b">"
    if formula is not None:
        return None, element("f", xml_escape(formula))
    if value is None:
        return None, b""
    if kind is not None:
        return kind if kind != "n" else None, element("v", xml_escape(value))
    if isinstance(value, str):
        require(len(value.encode("utf-16-le", "surrogatepass")) // 2 <= 32767,
                "TEXT_TOO_LONG", "文字超过 Excel 单元格长度上限 32767，请拆分到预留单元格")
        t = qualify("t")
        text = b"<" + t + b' xml:space="preserve">' + xml_escape(encode_xstring(value)) + b"</" + t + b">"
        if rpr:
            text = element("r", rpr + text)
        return "inlineStr", element("is", text)
    if isinstance(value, bool):
        return "b", element("v", b"1" if value else b"0")
    require(not isinstance(value, int) or abs(value) <= 999999999999999,
            "NUMBER_PRECISION", "超过 Excel 15 位有效整数精度的值请作为文字写入")
    return None, element("v", repr(value).encode("ascii"))


def changed_cell(book: Workbook, doc: XMLDocument, node, change: CellChange) -> bytes:
    raw = doc.raw(node)
    edits = []
    tag = raw[:node.open_end-node.start]
    if change.style is not None:
        tag = _set_open_attribute(tag, b"s", change.style)
    if change.value is not UNSET or change.formula is not None:
        if change.cached:
            kind, fragment = value_fragment(node, change.value, None, None, change.cache_type or "n")
        else:
            kind, fragment = value_fragment(node, change.value, change.formula, book.uniform_rpr(node))
        tag = set_attribute(doc.data, node, b"t", kind)
        if change.style is not None:
            tag = _set_open_attribute(tag, b"s", change.style)
        if node.self_closing:
            return tag[:-2] + b">" + fragment + b"</" + node.raw_name + b">"
        for child in node.children:
            if child.ns == S and child.local in (("v", "is") if change.preserve_formula else ("v", "is", "f")):
                edits.append(Edit(child.start-node.start, child.end-node.start, b"", "cell:value", node.attrs["r"]))
        insert = next((n.start-node.start for n in node.children if n.local == "extLst"), node.close_start-node.start)
        # Put the replacement at the first removed value node, preserving schema order and surrounding bytes.
        if edits:
            first = min(edits, key=lambda e: e.start)
            first.replacement = fragment
        elif fragment:
            edits.append(Edit(insert, insert, fragment, "cell:value", node.attrs["r"]))
    edits.append(Edit(0, node.open_end-node.start, tag, "cell:attributes", node.attrs["r"]))
    return apply_edits(raw, edits)[0]


def _set_open_attribute(tag: bytes, name: bytes, value: str | None) -> bytes:
    # Parse a namespace-neutral clone to locate lexical attributes; its offsets match the original tag.
    import re
    regex = re.compile(br"\s+" + re.escape(name) + br"\s*=\s*(['\"])(.*?)\1", re.S)
    match = regex.search(tag)
    if match:
        return tag[:match.start(2)] + xml_escape(value, True) + tag[match.end(2):] if value is not None else tag[:match.start()] + tag[match.end():]
    if value is None:
        return tag
    cut = len(tag) - (2 if tag.endswith(b"/>") else 1)
    return tag[:cut] + b" " + name + b'="' + xml_escape(value, True) + b'"' + tag[cut:]


def apply_cell_changes(book: Workbook, changes: dict[str, dict[str, CellChange]]) -> None:
    for name, entries in changes.items():
        sheet = book.sheet(name)
        doc = book.sheet_doc(sheet)
        data = doc.root.child("sheetData", S)
        require(data, "INVALID_WORKSHEET", "工作表缺少 sheetData")
        rows = {int(n.attrs["r"]): n for n in data.children if n.name == S + "}row"}
        cells = book.cell_map(sheet)
        edits = []
        missing: dict[int, list[tuple[int, bytes]]] = {}
        for address, change in entries.items():
            if address in cells:
                node = cells[address]
                replacement = changed_cell(book, doc, node, change)
                edits.append(Edit(node.start, node.end, replacement, "cell:cache" if change.cached else "cell:write", f"{name}!{address}"))
            else:
                require(not change.cached, "CALCULATION_LAYOUT_CHANGED", "计算结果新增了公式单元格，无法无损映射", cell=address)
                row, col = cell_position(address)
                tag = data.qualify("c")
                style = b' s="' + change.style.encode() + b'"' if change.style is not None else b""
                virtual = b"<" + tag + b' xmlns="' + S.encode() + b'" r="' + address.encode() + b'"' + style + b"/>"
                if b":" in tag:
                    prefix = tag.split(b":")[0]
                    virtual = virtual.replace(b' xmlns="', b' xmlns:' + prefix + b'="')
                virtual_doc = XMLDocument(virtual)
                kind, fragment = value_fragment(virtual_doc.root, change.value if change.value is not UNSET else None, change.formula, None)
                opening = b"<" + tag + b' r="' + address.encode() + b'"' + style
                if kind:
                    opening += b' t="' + kind.encode() + b'"'
                raw = opening + b">" + fragment + b"</" + tag + b">" if fragment else opening + b"/>"
                missing.setdefault(row, []).append((col, raw))
        inserts: dict[int, list[bytes]] = {}
        for row_num, additions in sorted(missing.items()):
            additions.sort()
            row = rows.get(row_num)
            if row:
                if row.self_closing:
                    fragment = b"".join(raw for col, raw in additions)
                    edits.append(append_children(doc.data, row, fragment, "cell:new", f"{name}!row:{row_num}"))
                else:
                    existing = sorted([(cell_position(n.attrs["r"])[1], n) for n in row.children if n.name == S + "}c"])
                    for col, raw in additions:
                        after = next((n for c, n in existing if c > col), None)
                        pos = after.start if after else next((n.start for n in row.children if n.local == "extLst"), row.close_start)
                        inserts.setdefault(pos, []).append(raw)
            else:
                tag = data.qualify("row")
                raw = b"<" + tag + b' r="' + str(row_num).encode() + b'">' + b"".join(raw for col, raw in additions) + b"</" + tag + b">"
                after = next((rows[n] for n in sorted(rows) if n > row_num), None)
                pos = after.start if after else data.close_start
                inserts.setdefault(pos, []).append(raw)
        if data.self_closing and missing:
            require(not edits, "INVALID_WORKSHEET", "自闭合 sheetData 与已有单元格冲突")
            fragment = b"".join(inserts.pop(data.close_start, []))
            edits.append(append_children(doc.data, data, fragment, "cell:new_rows", name))
        for pos, fragments in inserts.items():
            edits.append(Edit(pos, pos, b"".join(fragments), "cell:new", name))
        if missing:
            dimension = doc.root.child("dimension", S)
            bounds = Rectangle.parse(dimension.attrs["ref"]) if dimension else None
            rows_cols = [cell_position(address) for address in entries]
            top = min([r for r, c in rows_cols] + ([bounds.top] if bounds else []))
            left = min([c for r, c in rows_cols] + ([bounds.left] if bounds else []))
            bottom = max([r for r, c in rows_cols] + ([bounds.bottom] if bounds else []))
            right = max([c for r, c in rows_cols] + ([bounds.right] if bounds else []))
            reference = Rectangle(top, left, bottom, right).address
            if dimension:
                opening = set_attribute(doc.data, dimension, b"ref", reference)
                edits.append(Edit(dimension.start, dimension.open_end, opening, "cell:dimension", name))
            else:
                tag = doc.root.qualify("dimension")
                fragment = b"<" + tag + b' ref="' + reference.encode() + b'"/>'
                after = doc.root.child("sheetPr", S)
                offset = after.end if after else doc.root.open_end
                edits.append(Edit(offset, offset, fragment, "cell:dimension", name))
        book.package.patch(sheet["part"], edits)

