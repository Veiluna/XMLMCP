from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

from .errors import XMLMCPError, require
from .namespaces import S, C
from .package import Package, rels_path
from .workbook import Workbook, Rectangle, semantic_node, public_dict
from .cells import CellChange, apply_cell_changes, _set_open_attribute
from .images import geometry_requests
from .xmlbytes import Edit, append_children, standalone, digest, xml_escape


def preflight_calculation(book: Workbook, calculate: bool) -> None:
    if not calculate:
        return
    require(not book.unsupported_chart_parts(), "CHART_EXTENSION_UNSUPPORTED", "图表使用未适配的 chartEx 缓存，不能保证首次打开显示正确")
    for sheet in book.sheets:
        if sheet["kind"] != "worksheet":
            continue
        for cell in book.cell_map(sheet).values():
            formula = cell.child("f", S)
            if formula:
                require(not re.search(r"\b(?:WEBSERVICE|IMAGE|RTD|CUBEVALUE|CUBEMEMBER|CUBESET|SQL\.REQUEST)\s*\(", formula.text, re.I),
                        "OFFLINE_FORMULA_UNSUPPORTED", "该公式需要外部服务或数据模型，离线计算无法保证正确", cell=cell.attrs["r"])
    for pivot in book.pivots():
        cache_doc = pivot["_cache_doc"]
        require(not any(cache_doc.root.descendants("fieldGroup", S)), "GROUPED_PIVOT_UNSUPPORTED", "首版不刷新分组透视缓存")
        if pivot["source_type"] == "worksheet":
            fields = cache_doc.root.child("cacheFields", S)
            require(fields and all(n.local == "cacheField" and not n.attrs.get("formula") for n in fields.children),
                    "PIVOT_FIELDS_UNSUPPORTED", "透视字段结构不支持")
        else:
            require(not any("model" in r.get("Type", "").lower() for r in book.package.relationships(pivot["cache_part"]))
                    and not any(part.startswith("xl/model/") for part in book.package.parts),
                    "BINARY_MODEL_UNSUPPORTED", "透视结果依赖未适配的二进制数据模型")
    for chart in book.charts():
        require(chart["name"], "CHART_MAPPING", "图表缺少可稳定映射的对象名称")
        doc = chart["_doc"]
        require(not any(n.ns != C and n.local.endswith("Chart") for n in doc.root.descendants()),
                "CHART_UNSUPPORTED", "图表使用未适配的扩展结构")


def offline_candidate(package: Package, path: Path):
    """Only the disposable package changes connection auto-refresh settings."""
    parts = dict(package.parts)
    try:
        for part in list(parts):
            if not part.endswith(".xml") or not any(token in part for token in ("connections", "queryTables", "pivotCacheDefinition")):
                continue
            doc = package.doc(part)
            edits = []
            for node in [doc.root, *doc.root.descendants()]:
                attrs = {key: "0" for key in ("refreshOnLoad", "backgroundRefresh", "background") if key in node.attrs}
                if not attrs:
                    continue
                tag = doc.data[node.start:node.open_end]
                for attr, value in attrs.items():
                    tag = _set_open_attribute(tag, attr.encode(), value)
                edits.append(Edit(node.start, node.open_end, tag, "temporary:offline", part))
            if edits:
                from .xmlbytes import apply_edits
                package.parts[part] = apply_edits(doc.data, edits)[0]
        package.write(path)
    finally:
        package.parts = parts


def terminate_owned_excel(owner_path: Path) -> None:
    if not owner_path.is_file() or os.name != "nt":
        return
    import win32api
    import win32process
    import win32con
    try:
        owner = json.loads(owner_path.read_text(encoding="utf-8"))
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE, False, owner["pid"])
        try:
            if str(win32process.GetProcessTimes(handle)["CreationTime"]) == owner["created"]:
                win32api.TerminateProcess(handle, 1)
        finally:
            handle.Close()
    except Exception:
        pass  # Already closed, or ownership cannot be established: do not touch it.


class ExcelCalculator:
    def __init__(self, timeout: float = 180):
        self.timeout = timeout

    def run(self, book: Workbook, directory: Path, calculate: bool, images: list) -> tuple[dict, Package | None]:
        require(os.name == "nt", "EXCEL_REQUIRED", "计算公式、缓存或图片尺寸需要 Windows Microsoft Excel")
        candidate = directory / "candidate.xlsx"
        computed = directory / "computed.xlsx"
        offline_candidate(book.package, candidate)
        request_path, response_path = directory / "request.json", directory / "response.json"
        request = {"candidate": str(candidate), "computed": str(computed), "calculate": calculate,
                   "timeout": self.timeout, "geometry": geometry_requests(images),
                   "pivots": [public_dict(p) for p in book.pivots()],
                   "charts": [public_dict(c) for c in book.charts()]}
        request_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        owner = response_path.with_suffix(".owner.json")
        try:
            process = subprocess.run([sys.executable, "-m", "xmlmcp.excel_worker", str(request_path), str(response_path)],
                                     capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=self.timeout + 30,
                                     creationflags=subprocess.CREATE_NO_WINDOW)
            require(response_path.exists(), "EXCEL_FAILED", "Excel 工作进程未返回结果", error=process.stderr[-2000:])
            response = json.loads(response_path.read_text(encoding="utf-8"))
            if not response["ok"]:
                error = response["error"]
                raise XMLMCPError(error["code"], error["message"], **error.get("details", {}))
            require(process.returncode == 0, "EXCEL_FAILED", "Excel 工作进程异常退出")
            return response["result"], Package(computed) if calculate else None
        except subprocess.TimeoutExpired as exc:
            raise XMLMCPError("CALCULATION_TIMEOUT", "Excel 工作进程超时，已取消整批操作") from exc
        finally:
            terminate_owned_excel(owner)


def cached_value(book: Workbook, cell) -> tuple[object, str]:
    kind = cell.attrs.get("t", "n")
    v = cell.child("v", S)
    require(v is not None, "FORMULA_CACHE_MISSING", "Excel 未生成公式缓存", cell=cell.attrs["r"])
    if kind == "s":
        return book.value(cell), "str"
    return v.text, kind


def harvest_formula_caches(book: Workbook, computed: Workbook):
    changes = {}
    count = 0
    for sheet in book.sheets:
        if sheet["kind"] != "worksheet":
            continue
        source_cells = book.cell_map(sheet)
        result_cells = computed.cell_map(computed.sheet(sheet["name"]))
        for address, cell in source_cells.items():
            f = cell.child("f", S)
            if not f:
                continue
            result = result_cells.get(address)
            require(result and result.child("f", S), "FORMULA_MAPPING", "公式单元格在计算副本中无法定位", sheet=sheet["name"], cell=address)
            result_f = result.child("f", S)
            if f.attrs.get("t") not in ("shared", "array", "dataTable"):
                require(result_f.attrs.get("t") not in ("array", "dataTable") and normalized_formula(f.text) == normalized_formula(result_f.text),
                        "FORMULA_MAPPING", "计算改变了普通公式或产生了数组区域，不能安全回填", cell=address,
                        input_formula=f.text, computed_formula=result_f.text, computed_attributes=result_f.attrs)
            if f.attrs.get("t") in ("array", "dataTable"):
                require(f.attrs.get("ref") and f.attrs.get("ref") == result_f.attrs.get("ref"),
                        "CALCULATION_LAYOUT_CHANGED", "数组或数据表计算改变了结果范围")
            value, kind = cached_value(computed, result)
            require(kind in ("n", "str", "b", "e"), "FORMULA_CACHE_TYPE", "不支持的公式缓存类型")
            if kind == "e":
                old = cell.child("v", S)
                require(cell.attrs.get("t") == "e" and old and old.text == value,
                        "FORMULA_ERROR", "计算产生了错误值", sheet=sheet["name"], cell=address, error=value)
            changes.setdefault(sheet["name"], {})[address] = CellChange(value=value, cached=True, cache_type=kind, preserve_formula=True)
            count += 1
            if f.attrs.get("t") in ("array", "dataTable"):
                for member in Rectangle.parse(f.attrs["ref"]).cells():
                    if member == address:
                        continue
                    original_member, result_member = source_cells.get(member), result_cells.get(member)
                    require(original_member and result_member, "CALCULATION_LAYOUT_CHANGED", "数组缓存创建了新单元格，无法保持布局", cell=member)
                    member_value, member_kind = cached_value(computed, result_member)
                    require(member_kind in ("n", "str", "b") or (member_kind == "e" and original_member.attrs.get("t") == "e"
                            and original_member.child("v", S) and original_member.child("v", S).text == member_value),
                            "FORMULA_ERROR", "数组成员产生错误或不支持的缓存类型", cell=member)
                    require(not book.uniform_rpr(original_member), "RICH_TEXT_TYPE_CHANGE", "数组缓存不能覆盖只存于字符运行的字体", cell=member)
                    changes.setdefault(sheet["name"], {})[member] = CellChange(value=member_value, cached=True, cache_type=member_kind, preserve_formula=True)
                    count += 1
    apply_cell_changes(book, changes)
    return count


def normalized_formula(text: str):
    return text.strip().removeprefix("=").replace("$", "").casefold()


def cache_points(cache) -> list:
    count = cache.child("ptCount", C)
    require(count, "CHART_CACHE_MISSING", "图表缓存缺少 ptCount")
    size = int(count.attrs["val"])
    points = [None] * size
    for point in cache.children:
        if point.local != "pt":
            continue
        index = int(point.attrs["idx"])
        require(0 <= index < size, "CHART_CACHE_INVALID", "图表缓存点索引越界")
        value = point.child("v", C)
        points[index] = value.text if value else None
    return points


def verify_chart_values(chart: dict, observed: dict):
    series = list(chart["_doc"].root.descendants("ser", C))
    series.sort(key=lambda node: int(node.child("order", C).attrs["val"]) if node.child("order", C) else 0)
    require(len(series) == len(observed["series"]), "CHART_SERIES_CHANGED", "图表系列数量发生变化")
    checked = 0
    for node, result in zip(series, observed["series"]):
        for role, key in (("val", "values"), ("yVal", "values"), ("xVal", "categories"), ("cat", "categories"), ("tx", "name")):
            parent = node.child(role, C)
            if not parent:
                continue
            literal_name = parent.child("v", C) if role == "tx" else None
            if literal_name is not None:
                require(literal_name.text == str(result[key]), "CHART_CACHE_STALE", "图表系列名称与 Excel 系列结果不一致", chart=chart["name"])
                checked += 1
                continue
            cache = next((n for n in parent.descendants() if n.ns == C and n.local in ("numCache", "strCache")), None)
            if not cache and any(parent.descendants("multiLvlStrCache", C)):
                raise XMLMCPError("CHART_MULTI_LEVEL_UNSUPPORTED", "多层分类缓存尚不能与 Excel 系列值可靠交叉验证")
            if not cache:
                cache = next((n for n in parent.descendants() if n.ns == C and n.local in ("numLit", "strLit")), None)
            require(cache, "CHART_CACHE_MISSING", "Excel 未保存系列缓存", chart=chart["name"], role=role)
            values = result[key]
            if not isinstance(values, list):
                values = [values]
            points = cache_points(cache)
            require(len(points) == len(values), "CHART_CACHE_STALE", "图表缓存与计算后的系列长度不一致", chart=chart["name"])
            for point, value in zip(points, values):
                if point is None:
                    require(value is None or value == "" or isinstance(value, int) and value < -2000000000,
                            "CHART_CACHE_STALE", "图表缓存缺少计算后的数据点", chart=chart["name"])
                elif cache.local in ("numCache", "numLit"):
                    try:
                        equal = abs(float(point) - float(value)) <= max(1e-10, abs(float(value))*1e-12)
                    except (ValueError, TypeError):
                        equal = False
                    require(equal, "CHART_CACHE_STALE", "图表数值缓存与 Excel 系列结果不一致", chart=chart["name"])
                else:
                    require(str(point) == str(value), "CHART_CACHE_STALE", "图表文本缓存与 Excel 系列结果不一致", chart=chart["name"])
            checked += 1
    return checked


def harvest_chart_caches(book: Workbook, computed: Workbook, observation: dict) -> int:
    result_charts = computed.charts()
    observations = {(c["sheet"], c["name"]): c for c in observation.get("charts", [])}
    count = 0
    for chart in book.charts():
        key = (chart["sheet"], chart["name"])
        matches = [c for c in result_charts if (c["sheet"], c["name"]) == key]
        require(len(matches) == 1 and key in observations, "CHART_MAPPING", "图表在计算副本中无法唯一映射", chart=key)
        result = matches[0]
        require(result["kinds"] == chart["kinds"] and [normalized_formula(f) for f in result["sources"]] == [normalized_formula(f) for f in chart["sources"]],
                "CHART_SOURCE_CHANGED", "计算改变了图表类型或来源表达式，无法安全回填", chart=key)
        verify_chart_values(result, observations[key])
        old_doc, new_doc = chart["_doc"], result["_doc"]
        old_refs = [n for n in old_doc.root.descendants() if n.ns == C and n.local in ("numRef", "strRef", "multiLvlStrRef")]
        new_refs = [n for n in new_doc.root.descendants() if n.ns == C and n.local in ("numRef", "strRef", "multiLvlStrRef")]
        require(len(old_refs) == len(new_refs), "CHART_MAPPING", "图表引用数量变化")
        edits = []
        for old, new in zip(old_refs, new_refs):
            old_f, new_f = old.child("f", C), new.child("f", C)
            require(old.local == new.local and old_f and new_f and normalized_formula(old_f.text) == normalized_formula(new_f.text),
                    "CHART_MAPPING", "图表缓存来源无法对应")
            kind = {"numRef": "numCache", "strRef": "strCache", "multiLvlStrRef": "multiLvlStrCache"}[old.local]
            cache = new.child(kind, C)
            require(cache, "CHART_CACHE_MISSING", "计算副本缺少图表引用缓存", chart=key)
            original = old.child(kind, C)
            fragment = standalone(new_doc, cache)
            if original:
                edits.extend(chart_cache_edits(old_doc, original, new_doc, cache, f"{key}:{old_f.text}"))
            else:
                edits.append(append_children(old_doc.data, old, fragment, "chart:cache", f"{key}:{old_f.text}"))
            count += 1
        book.package.patch(chart["chart_part"], edits)
    return count


def chart_cache_edits(old_doc, original, new_doc, cache, selector: str) -> list[Edit]:
    """Preserve cache extensions and point metadata; only cache values/counts change."""
    if original.self_closing:
        return [append_children(old_doc.data, original, b"".join(standalone(new_doc, n) for n in cache.children), "chart:cache", selector)]
    edits = []
    headers = []
    for local in ("formatCode", "ptCount"):
        old, new = original.child(local, C), cache.child(local, C)
        if new is None:
            continue
        if old is None:
            headers.append(standalone(new_doc, new))
        elif local == "ptCount":
            tag = _set_open_attribute(old_doc.data[old.start:old.open_end], b"val", new.attrs["val"])
            edits.append(Edit(old.start, old.open_end, tag, "chart:cache_count", selector))
        elif old.text != new.text:
            require(not old.children, "CHART_CACHE_EXTENSION", "缓存格式代码包含未适配结构")
            if old.self_closing:
                edits.append(append_children(old_doc.data, old, xml_escape(new.text), "chart:cache_format", selector))
            else:
                edits.append(Edit(old.open_end, old.close_start, xml_escape(new.text), "chart:cache_format", selector))
    if headers:
        offset = original.open_end
        edits.append(Edit(offset, offset, b"".join(headers), "chart:cache_headers", selector))
    old_points = [n for n in original.children if n.name == C + "}pt"]
    new_points = [n for n in cache.children if n.name == C + "}pt"]
    old_ids = [n.attrs["idx"] for n in old_points]
    new_ids = [n.attrs["idx"] for n in new_points]
    if old_ids == new_ids:
        for old, new in zip(old_points, new_points):
            a, b = old.child("v", C), new.child("v", C)
            require(a and b and not a.children and not b.children, "CHART_CACHE_EXTENSION", "数据点的值结构无法可靠更新")
            if a.text != b.text:
                if a.self_closing:
                    edits.append(append_children(old_doc.data, a, xml_escape(b.text), "chart:cache_value", selector + ":" + old.attrs["idx"]))
                else:
                    edits.append(Edit(a.open_end, a.close_start, xml_escape(b.text), "chart:cache_value", selector + ":" + old.attrs["idx"]))
    else:
        require(all(set(n.attrs) == {"idx"} and all(c.name == C + "}v" and not c.children for c in n.children) for n in old_points),
                "CHART_CACHE_EXTENSION", "数据点集合变化会影响未适配的点元数据")
        payload = b"".join(standalone(new_doc, n) for n in new_points)
        for i, node in enumerate(old_points):
            edits.append(Edit(node.start, node.end, payload if i == 0 else b"", "chart:cache_points", selector))
        if not old_points and payload:
            offset = next((n.start for n in original.children if n.local == "extLst"), original.close_start)
            edits.append(Edit(offset, offset, payload, "chart:cache_points", selector))
    # Multi-level caches are rejected during COM cross-validation before reaching this adapter.
    inserts = {}
    result = []
    for edit in edits:
        if edit.start == edit.end:
            inserts.setdefault(edit.start, []).append(edit)
        else:
            result.append(edit)
    for offset, group in inserts.items():
        result.append(Edit(offset, offset, b"".join(e.replacement for e in group), "chart:cache", selector))
    return result



def _same_children_except(old, new, allowed: set[str]):
    return [semantic_node(n) for n in old.children if n.local not in allowed] == [semantic_node(n) for n in new.children if n.local not in allowed]


def harvest_pivot_caches(book: Workbook, computed: Workbook):
    computed_pivots = {(p["sheet"], p["name"]): p for p in computed.pivots()}
    processed = set()
    count = 0
    grid_changes = {}
    for pivot in book.pivots():
        if pivot["source_type"] != "worksheet":
            continue
        result = computed_pivots.get((pivot["sheet"], pivot["name"]))
        require(result and result["range"] == pivot["range"], "PIVOT_LAYOUT_CHANGED", "透视结果范围改变")
        old_doc, new_doc = pivot["_doc"], result["_doc"]
        old_location, new_location = old_doc.root.child("location", S), new_doc.root.child("location", S)
        require(semantic_node(old_location) == semantic_node(new_location), "PIVOT_LAYOUT_CHANGED", "透视表布局属性变化")
        edits = []
        for local in ("rowItems", "colItems"):
            old, new = old_doc.root.child(local, S), new_doc.root.child(local, S)
            require(bool(old) == bool(new), "PIVOT_LAYOUT_CHANGED", "透视行列项目结构变化")
            if old and semantic_node(old) != semantic_node(new):
                edits.append(Edit(old.start, old.end, standalone(new_doc, new), "pivot:items", pivot["name"] + ":" + local))
        old_fields, new_fields = old_doc.root.child("pivotFields", S), new_doc.root.child("pivotFields", S)
        require(old_fields and new_fields and len(old_fields.children) == len(new_fields.children), "PIVOT_LAYOUT_CHANGED", "透视字段数量变化")
        for i, (old, new) in enumerate(zip(old_fields.children, new_fields.children)):
            require(_same_children_except(old, new, {"items"}), "PIVOT_EXTENSION_CHANGED", "透视字段的其他配置变化，不能回填")
            old_items, new_items = old.child("items", S), new.child("items", S)
            require(bool(old_items) == bool(new_items), "PIVOT_LAYOUT_CHANGED", "透视字段项目结构变化")
            if old_items and semantic_node(old_items) != semantic_node(new_items):
                require(len(old_items.children) == len(new_items.children), "PIVOT_LAYOUT_CHANGED", "透视项目数量变化")
                for a, b in zip(old_items.children, new_items.children):
                    require({k: v for k, v in a.attrs.items() if k not in ("x", "t")} == {k: v for k, v in b.attrs.items() if k not in ("x", "t")},
                            "PIVOT_FORMAT_CHANGED", "透视项目的显示配置变化")
                edits.append(Edit(old_items.start, old_items.end, standalone(new_doc, new_items), "pivot:field_items", f"{pivot['name']}:{i}"))
        book.package.patch(pivot["part"], edits)
        if pivot["cache_part"] not in processed:
            processed.add(pivot["cache_part"])
            old_cache, new_cache = pivot["_cache_doc"], result["_cache_doc"]
            old_fields, new_fields = old_cache.root.child("cacheFields", S), new_cache.root.child("cacheFields", S)
            require(old_fields and new_fields and len(old_fields.children) == len(new_fields.children), "PIVOT_SCHEMA_CHANGED", "缓存字段结构变化")
            cache_edits = []
            tag = old_cache.data[old_cache.root.start:old_cache.root.open_end]
            if "recordCount" in new_cache.root.attrs:
                tag = _set_open_attribute(tag, b"recordCount", new_cache.root.attrs["recordCount"])
                cache_edits.append(Edit(old_cache.root.start, old_cache.root.open_end, tag, "pivot:record_count", pivot["cache_part"]))
            for i, (old, new) in enumerate(zip(old_fields.children, new_fields.children)):
                require(old.attrs.get("name") == new.attrs.get("name") and _same_children_except(old, new, {"sharedItems"}),
                        "PIVOT_SCHEMA_CHANGED", "缓存字段名称或结构变化")
                a, b = old.child("sharedItems", S), new.child("sharedItems", S)
                require(bool(a) == bool(b), "PIVOT_SCHEMA_CHANGED", "共享项目结构变化")
                if a and semantic_node(a) != semantic_node(b):
                    cache_edits.append(Edit(a.start, a.end, standalone(new_cache, b), "pivot:shared_items", f"{pivot['cache_part']}:{i}"))
            book.package.patch(pivot["cache_part"], cache_edits)
            old_records = book.package.related(pivot["cache_part"], "pivotCacheRecords")
            new_records = computed.package.related(result["cache_part"], "pivotCacheRecords")
            require(len(old_records) == len(new_records) == 1, "PIVOT_RECORDS_UNSUPPORTED", "缓存记录无法映射")
            old_doc = book.package.doc(old_records[0]["part"])
            new_doc = computed.package.doc(new_records[0]["part"])
            require(all(n.name == S + "}r" for n in old_doc.root.children) and all(n.name == S + "}r" for n in new_doc.root.children),
                    "PIVOT_EXTENSION_UNSUPPORTED", "缓存记录包含未适配扩展")
            fragment = b"".join(standalone(new_doc, n) for n in new_doc.root.children)
            opening = _set_open_attribute(old_doc.data[old_doc.root.start:old_doc.root.open_end], b"count", new_doc.root.attrs["count"])
            if old_doc.root.self_closing:
                raw = opening[:-2] + b">" + fragment + b"</" + old_doc.root.raw_name + b">"
                changes = [Edit(old_doc.root.start, old_doc.root.end, raw, "pivot:records", pivot["cache_part"])]
            else:
                changes = [Edit(old_doc.root.start, old_doc.root.open_end, opening, "pivot:record_count", pivot["cache_part"]),
                           Edit(old_doc.root.open_end, old_doc.root.close_start, fragment, "pivot:records", pivot["cache_part"])]
            book.package.patch(old_records[0]["part"], changes)
            count += 1
        source_sheet, result_sheet = book.sheet(pivot["sheet"]), computed.sheet(pivot["sheet"])
        old_cells, new_cells = book.cell_map(source_sheet), computed.cell_map(result_sheet)
        for address in Rectangle.parse(pivot["range"]).cells():
            a, b = old_cells.get(address), new_cells.get(address)
            if (a and a.child("f", S)) or (b and b.child("f", S)):
                continue
            value = computed.value(b)
            if book.value(a) == value:
                continue
            require(a is not None, "PIVOT_LAYOUT_CHANGED", "透视刷新创建了新单元格，无法保留预设格式", cell=address)
            book.uniform_rpr(a)
            grid_changes.setdefault(pivot["sheet"], {})[address] = CellChange(value=value)
    apply_cell_changes(book, grid_changes)
    return count


def update_calculation_metadata(book: Workbook):
    package = book.package
    doc = package.doc(book.part)
    calc = doc.root.child("calcPr", S)
    if calc:
        tag = doc.data[calc.start:calc.open_end]
        for name, value in ((b"calcMode", "auto"), (b"fullCalcOnLoad", "1"), (b"forceFullCalc", "1")):
            tag = _set_open_attribute(tag, name, value)
        package.patch(book.part, [Edit(calc.start, calc.open_end, tag, "calculation:properties", "calcPr")])
    else:
        tag = doc.root.qualify("calcPr")
        fragment = b"<" + tag + b' calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        # calcPr precedes these workbook elements in the SpreadsheetML sequence.
        after = {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
        following = next((n for n in doc.root.children if n.local in after), None)
        offset = following.start if following else doc.root.close_start
        package.patch(book.part, [Edit(offset, offset, fragment, "calculation:properties", "calcPr")])
    chains = package.related(book.part, "calcChain")
    for chain in chains:
        part = chain["part"]
        package.remove(part, "calculation:chain")
        path = rels_path(book.part)
        rel_doc = package.doc(path)
        node = next(n for n in rel_doc.root.children if n.attrs.get("Id") == chain["Id"])
        package.patch(path, [Edit(node.start, node.end, b"", "calculation:chain_relationship", chain["Id"])])
        ct = package.doc("[Content_Types].xml")
        node = next((n for n in ct.root.children if n.attrs.get("PartName") == "/" + part), None)
        if node:
            package.patch("[Content_Types].xml", [Edit(node.start, node.end, b"", "calculation:chain_content_type", part)])

