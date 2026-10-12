from __future__ import annotations

import base64
import json
from pathlib import Path

from .errors import XMLMCPError, require
from .models import IntegrityResult
from .namespaces import C, S, R, XDR
from .package import Package, absolute_file, read_xml
from .workbook import Workbook
from .xmlbytes import Edit, apply_edits, digest


def validate_caches(book: Workbook, calculated: bool = False) -> dict:
    package = book.package
    formulas = 0
    caches = 0
    for sheet in book.sheets:
        if sheet["kind"] != "worksheet":
            continue
        for node in book.cell_map(sheet).values():
            if node.child("f", S):
                formulas += 1
                if calculated:
                    require(node.child("v", S), "FORMULA_CACHE_MISSING", "已计算结果缺少公式缓存", cell=node.attrs["r"])
    for chart in book.charts():
        for cache in chart["_doc"].root.descendants():
            if cache.ns != C or cache.local not in ("numCache", "strCache", "multiLvlStrCache"):
                continue
            count = cache.child("ptCount", C)
            require(count, "CHART_CACHE_INVALID", "图表缓存缺少 ptCount")
            size = int(count.attrs["val"])
            levels = [cache] if cache.local != "multiLvlStrCache" else [n for n in cache.children if n.name == C + "}lvl"]
            for level in levels:
                indices = [int(p.attrs["idx"]) for p in level.children if p.name == C + "}pt"]
                require(len(set(indices)) == len(indices) and all(0 <= i < size for i in indices),
                        "CHART_CACHE_INVALID", "图表缓存点越界或重复")
            caches += 1
    processed = set()
    for pivot in book.pivots():
        if pivot["cache_part"] in processed:
            continue
        processed.add(pivot["cache_part"])
        fields = pivot["_cache_doc"].root.child("cacheFields", S)
        require(fields, "PIVOT_CACHE_INVALID", "透视缓存缺少字段")
        records = package.related(pivot["cache_part"], "pivotCacheRecords")
        for relation in records:
            doc = package.doc(relation["part"])
            rows = [n for n in doc.root.children if n.name == S + "}r"]
            require(int(doc.root.attrs["count"]) == len(rows), "PIVOT_CACHE_INVALID", "透视记录计数不一致")
            for row in rows:
                require(len(row.children) == len(fields.children), "PIVOT_CACHE_INVALID", "透视记录字段数量不一致")
                for field, value in zip(fields.children, row.children):
                    if value.name == S + "}x":
                        shared = field.child("sharedItems", S)
                        require(shared and 0 <= int(value.attrs["v"]) < len(shared.children),
                                "PIVOT_CACHE_INVALID", "透视共享项目索引越界")
    # Check relationship attributes in each XML against its own relationship part.
    for part in package.parts:
        if not part.endswith(".xml"):
            continue
        root = read_xml(package.parts[part])
        ids = {r["Id"] for r in package.relationships(part)}
        for node in root.iter():
            for attr, value in node.attrib.items():
                if attr in ("{" + R + "}id", "{" + R + "}embed", "{" + R + "}link"):
                    require(value in ids, "BROKEN_RELATIONSHIP_ID", "XML 使用了不存在的关系 ID", part=part, relationship=value)
    return {"cache_structure_valid": True, "formula_count": formulas, "chart_cache_count": caches,
            "pivot_cache_count": len(processed), "relationship_ids_valid": True}


def verify_packages(source: Package, output: Package, report: dict) -> IntegrityResult:
    errors = []
    checks = {}
    changed, added, removed = [], [], []
    unchanged = 0
    try:
        require(report.get("schema_version") == 1, "REPORT_VERSION", "完整性报告版本不支持")
        require(report.get("input_fingerprint") == source.fingerprint, "INPUT_FINGERPRINT", "输入文件指纹与报告不一致")
        require(report.get("output_fingerprint") == output.fingerprint, "OUTPUT_FINGERPRINT", "输出文件指纹与报告不一致")
        original_hashes = {n: digest(d) for n, d in source.parts.items()}
        actual_hashes = {n: digest(d) for n, d in output.parts.items()}
        require(report["original_parts"] == original_hashes, "INPUT_PARTS", "输入部件清单与报告不一致")
        require(report["output_parts"] == actual_hashes, "OUTPUT_PARTS", "输出部件清单与报告不一致")
        expected = dict(source.parts)
        for stage in report["stages"]:
            part = stage["part"]
            require(part in expected and digest(expected[part]) == stage["before_sha256"], "PATCH_SEQUENCE", "补丁的前置部件指纹不一致", part=part)
            edits = []
            for patch in stage["patches"]:
                start, end = patch["start"], patch["end"]
                require(isinstance(start, int) and isinstance(end, int) and 0 <= start <= end <= len(expected[part]),
                        "PATCH_RANGE", "报告中的补丁越界")
                require(digest(expected[part][start:end]) == patch["old_sha256"], "PATCH_SOURCE", "补丁原始片段不一致", part=part)
                require(patch["reason"].split(":")[0] in ("cell", "image", "chart", "pivot", "calculation"),
                        "PATCH_REASON", "报告含未识别的修改类别")
                payload = base64.b64decode(patch["replacement_base64"], validate=True)
                edits.append(Edit(start, end, payload, patch["reason"], patch["selector"]))
            expected[part] = apply_edits(expected[part], edits)[0]
            require(digest(expected[part]) == stage["after_sha256"], "PATCH_RESULT", "补丁结果指纹不一致", part=part)
        for part, info in report["removed_parts"].items():
            require(part in expected and digest(expected[part]) == info["sha256"] and info["reason"] == "calculation:chain",
                    "REMOVED_PART", "删除部件不在许可清单中")
            del expected[part]
        for part, info in report["added_parts"].items():
            require(part not in expected and part in output.parts and digest(output.parts[part]) == info["sha256"]
                    and info["reason"].startswith("image:"), "ADDED_PART", "新增部件不在许可清单中")
            expected[part] = output.parts[part]
        require(expected == output.parts, "UNAUTHORIZED_CHANGE", "输出含补丁清单以外的字节变化")
        checks["patch_replay_valid"] = True
        checks["outside_patches_identical"] = True
        checks.update(output.validate())
        checks.update(validate_caches(Workbook(output), bool(report.get("calculation", {}).get("calculated"))))
    except (XMLMCPError, KeyError, ValueError, TypeError) as exc:
        errors.append(str(exc))
        if isinstance(exc, XMLMCPError):
            checks["failure_code"] = exc.code
    for part in source.parts.keys() | output.parts.keys():
        if part not in source.parts:
            added.append(part)
        elif part not in output.parts:
            removed.append(part)
        elif source.parts[part] != output.parts[part]:
            changed.append(part)
        else:
            unchanged += 1
    return IntegrityResult(valid=not errors, input_fingerprint=source.fingerprint, output_fingerprint=output.fingerprint,
                           unchanged_parts=unchanged, changed_parts=sorted(changed), added_parts=sorted(added),
                           removed_parts=sorted(removed), checks=checks, errors=errors)


def verify_integrity(input_path: str, output_path: str, report_path: str) -> IntegrityResult:
    source = Package(input_path)
    output = Package(output_path)
    path = absolute_file(report_path)
    try:
        report = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeError) as exc:
        raise XMLMCPError("INVALID_REPORT", "无法读取完整性报告") from exc
    return verify_packages(source, output, report)

