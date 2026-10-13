"""Restricted Office-independent path; never claims general Excel recalculation."""
from __future__ import annotations
import math
import re
from .errors import require
from .namespaces import S, C, A, XDR
from .workbook import Workbook


def preflight_offline(book: Workbook) -> None:
    require(not any(cell.child("f", S) for sheet in book.sheets if sheet["kind"] == "worksheet"
                    for cell in book.cell_map(sheet).values()),
            "OFFLINE_FORMULA", "本机报告流程不执行工作簿公式；请使用无公式模板或 Windows Excel 计算流程")
    require(not book.pivots() and not book.unsupported_chart_parts(), "OFFLINE_CACHE", "模板含需重算的透视表或未知图表")
    for chart in book.charts():
        for source in chart["_doc"].root.descendants("f", C):
            formula = source.text
            literal = re.fullmatch(r'"(?:[^\"]|\"\")*"', formula)
            external = re.fullmatch(r"(?:'\[\d+\](?:[^']|'')+'|\[\d+\][^'!\[\]]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?", formula)
            require(literal or external, "OFFLINE_CHART_DEPENDENCY",
                    "图表引用本地、动态或无法证明独立的数据；不能留下旧缓存", chart=chart["name"], source=formula)


def image_geometry(book: Workbook, changes: list, digit_width: float = 7) -> dict:
    result = {}
    for change in changes:
        if change.rectangle:
            rect = change.rectangle
            require(rect.bottom <= 10000 and rect.right <= 1000, "IMAGE_GEOMETRY", "离线图片目标超出支持的布局范围")
            sheet = book.sheet(change.sheet)
            doc = book.sheet_doc(sheet)
            settings = doc.root.child("sheetFormatPr", S)
            default_height = float(settings.attrs.get("defaultRowHeight", "15")) if settings else 15
            default_width = float(settings.attrs.get("defaultColWidth", "8.7109375")) if settings else 8.7109375
            widths = [default_width] * rect.right
            hidden = set()
            for col in doc.root.descendants("col", S):
                first, last = int(col.attrs["min"]), int(col.attrs["max"])
                for i in range(max(1, first), min(last, rect.right) + 1):
                    widths[i - 1] = float(col.attrs.get("width", str(default_width)))
                    if col.attrs.get("hidden") in ("1", "true"):
                        hidden.add(i - 1)
            points = [0 if i in hidden else math.trunc((256 * width + math.trunc(128 / digit_width)) / 256 * digit_width) * .75
                      for i, width in enumerate(widths)]
            heights = [default_height] * rect.bottom
            for row in doc.root.descendants("row", S):
                i = int(row.attrs["r"]) - 1
                if i < rect.bottom:
                    heights[i] = 0 if row.attrs.get("hidden") in ("1", "true") else float(row.attrs.get("ht", str(default_height)))
            dimensions = {"left": sum(points[:rect.left - 1]), "top": sum(heights[:rect.top - 1]),
                          "width": sum(points[rect.left - 1:rect.right]), "height": sum(heights[rect.top - 1:rect.bottom]),
                          "basis": "OOXML row heights and configured maximum digit width", "maximum_digit_width_px": digit_width}
        else:
            pic = change.image["_pic"]
            frame = pic.parent
            ext = frame.child("ext", XDR) if frame.local == "oneCellAnchor" else None
            if ext is None:
                transform = pic.child("spPr", XDR).child("xfrm", A)
                ext = transform.child("ext", A) if transform else None
            require(ext is not None, "IMAGE_GEOMETRY", "图片没有明确的原始尺寸")
            dimensions = {"width": int(ext.attrs["cx"]) / 12700, "height": int(ext.attrs["cy"]) / 12700,
                          "basis": "existing OOXML picture extent"}
        require(dimensions["width"] > 0 and dimensions["height"] > 0, "IMAGE_GEOMETRY", "图片目标尺寸无效或已隐藏")
        result[change.geometry_key] = dimensions
    return result
