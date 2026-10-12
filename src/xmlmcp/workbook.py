from __future__ import annotations

import re
from dataclasses import dataclass

from .errors import XMLMCPError, require
from .models import Target
from .namespaces import S, R, C, XDR, A
from .package import Package
from .strings import decode_xstring
from .xmlbytes import Node, XMLDocument, digest, standalone

CELL_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)\Z")
STATIC_NAME_RE = re.compile(r"(?P<sheet>'(?:[^']|'')+'|[^'!\[\],()+*/:]+)!(?P<range>\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)\Z")


def cell_position(address: str) -> tuple[int, int]:
    match = CELL_RE.fullmatch(address)
    require(match, "INVALID_RANGE", "单元格坐标无效", address=address)
    col = 0
    for char in match[1].upper():
        col = col * 26 + ord(char) - 64
    row = int(match[2])
    require(col <= 16384 and row <= 1048576, "INVALID_RANGE", "坐标超出 Excel 范围")
    return row, col


def cell_address(row: int, col: int) -> str:
    chars = []
    while col:
        col, rem = divmod(col - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars)) + str(row)


@dataclass(frozen=True)
class Rectangle:
    top: int
    left: int
    bottom: int
    right: int

    @classmethod
    def parse(cls, address: str) -> Rectangle:
        bits = address.split(":")
        require(len(bits) in (1, 2), "INVALID_RANGE", "只支持连续矩形范围")
        top, left = cell_position(bits[0])
        bottom, right = cell_position(bits[-1])
        require(top <= bottom and left <= right, "INVALID_RANGE", "范围起止顺序无效")
        return cls(top, left, bottom, right)

    @property
    def address(self) -> str:
        start, end = cell_address(self.top, self.left), cell_address(self.bottom, self.right)
        return start if start == end else start + ":" + end

    @property
    def size(self) -> tuple[int, int]:
        return self.bottom - self.top + 1, self.right - self.left + 1

    def contains(self, row: int, col: int) -> bool:
        return self.top <= row <= self.bottom and self.left <= col <= self.right

    def cells(self):
        require(self.size[0] * self.size[1] <= 100000, "RANGE_TOO_LARGE", "单次范围最多处理 100000 个单元格")
        for row in range(self.top, self.bottom + 1):
            for col in range(self.left, self.right + 1):
                yield cell_address(row, col)


class Workbook:
    def __init__(self, package: Package):
        self.package = package
        roots = package.related("", "officeDocument")
        require(len(roots) == 1 and roots[0]["part"], "INVALID_XLSX", "无法唯一定位工作簿")
        self.part = roots[0]["part"]
        doc = package.doc(self.part)
        require(doc.root.ns == S, "STRICT_OOXML_UNSUPPORTED", "首版仅修改 Transitional OOXML；其他命名空间保留读取但不修改")
        relationships = {r["Id"]: r for r in package.relationships(self.part)}
        sheets = doc.root.child("sheets", S)
        require(sheets, "INVALID_XLSX", "工作簿缺少 sheets")
        self.sheets = []
        for index, sheet in enumerate(sheets.children):
            rel = relationships[sheet.attrs[R + "}id"]]
            self.sheets.append({"name": sheet.attrs["name"], "sheet_id": sheet.attrs["sheetId"],
                                "index": index, "part": rel["part"],
                                "kind": rel["Type"].rsplit("/", 1)[-1], "state": sheet.attrs.get("state", "visible")})
        self.by_name = {s["name"]: s for s in self.sheets}
        styles = package.related(self.part, "styles")
        self.style_part = styles[0]["part"] if styles else None
        self.style_doc = package.doc(self.style_part) if self.style_part else None
        shared = package.related(self.part, "sharedStrings")
        self.shared_doc = package.doc(shared[0]["part"]) if shared else None
        self.shared_strings = self.shared_doc.root.children if self.shared_doc else []
        names = doc.root.child("definedNames", S)
        self.names = []
        for n in names.children if names else []:
            local = n.attrs.get("localSheetId")
            scope = self.sheets[int(local)]["name"] if local is not None else "workbook"
            match = STATIC_NAME_RE.fullmatch(n.text)
            static = bool(match)
            self.names.append({"name": n.attrs["name"], "scope": scope, "expression": n.text,
                               "writable": static, "hidden": n.attrs.get("hidden") == "1"})

    def sheet(self, name: str) -> dict:
        require(name in self.by_name, "SHEET_NOT_FOUND", "工作表不存在", sheet=name)
        sheet = self.by_name[name]
        require(sheet["kind"] == "worksheet", "UNSUPPORTED_SHEET", "只能向普通工作表写入", sheet=name)
        return sheet

    def resolve(self, target: Target) -> tuple[dict, Rectangle]:
        if target.name is not None:
            matches = [n for n in self.names if n["name"].casefold() == target.name.casefold()
                       and (target.scope is None or n["scope"] == target.scope)]
            require(len(matches) == 1, "AMBIGUOUS_NAME" if matches else "NAME_NOT_FOUND",
                    "命名区域不存在或作用域不唯一；请指定 scope", name=target.name,
                    scopes=[n["scope"] for n in matches])
            expression = matches[0]["expression"]
            match = STATIC_NAME_RE.fullmatch(expression)
            require(match, "UNSUPPORTED_NAME", "动态、多区域、常量及外部命名区域不能作为写入目标", expression=expression)
            name = match["sheet"]
            name = name[1:-1].replace("''", "'") if name.startswith("'") else name
            return self.sheet(name), Rectangle.parse(match["range"])
        return self.sheet(target.sheet), Rectangle.parse(target.range)

    def sheet_doc(self, sheet: dict) -> XMLDocument:
        return self.package.doc(sheet["part"])

    def cell_map(self, sheet: dict) -> dict[str, Node]:
        doc = self.sheet_doc(sheet)
        cells = {}
        for n in doc.root.descendants("c", S):
            row, col = cell_position(n.attrs["r"])
            address = cell_address(row, col)
            require(address not in cells, "DUPLICATE_CELL", "工作表单元格坐标重复", cell=address)
            cells[address] = n
        return cells

    def merged_anchor(self, sheet: dict, rect: Rectangle, scalar: bool = False) -> Rectangle:
        doc = self.sheet_doc(sheet)
        for n in doc.root.descendants("mergeCell", S):
            merged = Rectangle.parse(n.attrs["ref"])
            if scalar and (rect == merged or rect.address == cell_address(merged.top, merged.left)):
                return Rectangle(merged.top, merged.left, merged.top, merged.left)
            intersects = not (rect.right < merged.left or rect.left > merged.right or rect.bottom < merged.top or rect.top > merged.bottom)
            if intersects:
                require(rect.size == (1, 1) and (rect.top, rect.left) == (merged.top, merged.left),
                        "MERGED_CELL_TARGET", "合并单元格只能向左上角写入", merge=merged.address)
        return rect

    def string_node(self, cell: Node | None) -> Node | None:
        if cell is None:
            return None
        kind = cell.attrs.get("t", "n")
        if kind == "inlineStr":
            return cell.child("is", S)
        if kind == "s":
            v = cell.child("v", S)
            if v:
                index = int(v.text)
                require(0 <= index < len(self.shared_strings), "INVALID_SHARED_STRING", "共享字符串索引无效")
                return self.shared_strings[index]
        return None

    def value(self, cell: Node | None):
        if cell is None:
            return None
        string = self.string_node(cell)
        if string:
            values = [n.text if n.local == "t" else n.child("t", S).text
                      for n in string.children if n.ns == S and (n.local == "t" or n.local == "r" and n.child("t", S))]
            return decode_xstring("".join(values))
        v = cell.child("v", S)
        if not v:
            return None
        kind = cell.attrs.get("t", "n")
        if kind == "b":
            return v.text == "1"
        if kind in ("e", "str", "d"):
            return decode_xstring(v.text) if kind == "str" else v.text
        try:
            return int(v.text) if re.fullmatch(r"[-+]?\d+", v.text) else float(v.text)
        except ValueError:
            return v.text

    def uniform_rpr(self, cell: Node | None) -> bytes | None:
        string = self.string_node(cell)
        if not string:
            return None
        require(not any(n.ns == S and n.local in ("rPh", "phoneticPr") for n in string.children),
                "PHONETIC_TEXT_UNSUPPORTED", "该文字含拼音或语音标注，不能证明覆盖后的标注仍正确；已取消修改")
        runs = [n for n in string.children if n.local in ("r", "t") and n.ns == S]
        properties = []
        raw_properties = []
        for run in runs:
            prop = run.child("rPr", S) if run.local == "r" else None
            semantic = semantic_node(prop) if prop else None
            properties.append(semantic)
            if prop:
                doc = prop.document
                raw_properties.append(standalone(doc, prop))
        require(len({repr(p) for p in properties}) <= 1, "MIXED_RICH_TEXT",
                "混合字符格式不能整格覆盖；请改用其他目标", cell=cell.attrs.get("r"))
        return raw_properties[0] if raw_properties else None

    def style_info(self, cell: Node | None) -> dict:
        index = int(cell.attrs.get("s", "0")) if cell else 0
        info = {"style_index": index}
        if self.style_doc:
            xfs = self.style_doc.root.child("cellXfs", S)
            require(xfs and 0 <= index < len(xfs.children), "INVALID_STYLE", "样式索引无效", style=index)
            info["style"] = xfs.children[index].attrs
        return info

    def drawings(self):
        for sheet in self.sheets:
            for relation in self.package.related(sheet["part"], "drawing"):
                part = relation["part"]
                doc = self.package.doc(part)
                rels = {r["Id"]: r for r in self.package.relationships(part)}
                yield sheet, part, doc, rels

    def images(self) -> list[dict]:
        result = []
        for sheet, part, doc, rels in self.drawings():
            for pic in doc.root.descendants("pic", XDR):
                nonvisual = next(pic.descendants("cNvPr", XDR), None)
                blip = next(pic.descendants("blip", A), None)
                if not nonvisual or not blip:
                    continue
                relationship = rels.get(blip.attrs.get(R + "}embed", ""))
                result.append({"image_id": f"{sheet['name']}|{part}|{nonvisual.attrs['id']}",
                               "sheet": sheet["name"], "drawing_part": part, "shape_id": nonvisual.attrs["id"],
                               "name": nonvisual.attrs.get("name", ""), "media_part": relationship["part"] if relationship else None,
                               "anchor_kind": pic.parent.local, "replaceable": pic.parent.ns == XDR and pic.parent.local in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"),
                               "_pic": pic, "_doc": doc})
        return result

    def charts(self) -> list[dict]:
        result = []
        for sheet, part, doc, rels in self.drawings():
            for node in doc.root.descendants("chart", C):
                frame = node
                while frame.parent and frame.local != "graphicFrame":
                    frame = frame.parent
                nonvisual = next(frame.descendants("cNvPr", XDR), None)
                relation = rels.get(node.attrs.get(R + "}id"))
                if not relation or not relation["part"]:
                    continue
                chart_part = relation["part"]
                chart_doc = self.package.doc(chart_part)
                sources = [n.text for n in chart_doc.root.descendants("f", C)]
                kinds = sorted({n.local for n in chart_doc.root.descendants() if n.ns == C and n.local.endswith("Chart")})
                name = nonvisual.attrs.get("name", "") if nonvisual else ""
                result.append({"sheet": sheet["name"], "name": name, "chart_part": chart_part,
                               "drawing_part": part, "kinds": kinds, "sources": sources,
                               "pivot": any(chart_doc.root.descendants("pivotSource", C)),
                               "external_snapshot": any(re.search(r"\[(?!0\])[^]]+\]", s) for s in sources),
                               "_doc": chart_doc, "_frame": frame})
        return result

    def unsupported_chart_parts(self) -> list[str]:
        return [relation["part"] for sheet, part, doc, rels in self.drawings()
                for relation in rels.values() if relation.get("Type", "").rsplit("/", 1)[-1] == "chartEx"]

    def pivots(self) -> list[dict]:
        caches = self.package.doc(self.part).root.child("pivotCaches", S)
        rels = {r["Id"]: r for r in self.package.relationships(self.part)}
        cache_map = {n.attrs["cacheId"]: rels[n.attrs[R + "}id"]]["part"] for n in caches.children} if caches else {}
        result = []
        for sheet in self.sheets:
            for relation in self.package.related(sheet["part"], "pivotTable"):
                doc = self.package.doc(relation["part"])
                definition = cache_map.get(doc.root.attrs["cacheId"])
                require(definition, "INVALID_PIVOT", "透视表无法定位缓存")
                cache_doc = self.package.doc(definition)
                source = cache_doc.root.child("cacheSource", S)
                location = doc.root.child("location", S)
                result.append({"sheet": sheet["name"], "name": doc.root.attrs["name"], "part": relation["part"],
                               "cache_part": definition, "source_type": source.attrs.get("type") if source else None,
                               "range": location.attrs.get("ref") if location else None,
                               "_doc": doc, "_cache_doc": cache_doc})
        return result


def semantic_node(node: Node | None):
    if node is None:
        return None
    return (node.name, tuple(sorted(node.attrs.items())), node.text, tuple(semantic_node(c) for c in node.children))


def public_dict(item: dict) -> dict:
    return {k: v for k, v in item.items() if not k.startswith("_")}

