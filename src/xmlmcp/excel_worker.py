"""One owned, hidden Excel process per request. Never supplies the final package."""
from __future__ import annotations

import gc
import json
import sys
import time
from pathlib import Path

from .errors import XMLMCPError, require


def serializable(value):
    if isinstance(value, (tuple, list)):
        return [serializable(v) for v in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    raise XMLMCPError("EXCEL_VALUE_TYPE", "Excel 返回了不支持的值类型", type=type(value).__name__)


def pivot_layout(pivot) -> dict:
    def fields(collection):
        return [str(collection.Item(i).Name) for i in range(1, collection.Count + 1)]
    return {"range": str(pivot.TableRange2.Address),
            "rows": fields(pivot.RowFields), "columns": fields(pivot.ColumnFields),
            "data": fields(pivot.DataFields), "pages": fields(pivot.PageFields)}


def wait_calculation(app, timeout: float):
    import pythoncom
    deadline = time.monotonic() + timeout
    while int(app.CalculationState) != 0:
        require(time.monotonic() < deadline, "CALCULATION_TIMEOUT", "Excel 计算超时")
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def run(request: dict, owner_path: Path) -> dict:
    import pythoncom
    import win32com.client
    import win32process
    import win32api
    import win32con

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    existing = set(win32process.EnumProcesses())
    app = None
    book = None
    owned = False
    result = {}
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
        require(pid not in existing, "EXCEL_ISOLATION", "Excel 未创建独立进程，不能继续计算")
        owned = True
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
        try:
            created = str(win32process.GetProcessTimes(handle)["CreationTime"])
        finally:
            handle.Close()
        owner_path.write_text(json.dumps({"pid": pid, "created": created}), encoding="utf-8")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        # No workbook is open while changing Application-wide calculation mode.
        try:
            app.Calculation = -4135
        except Exception:
            pass  # Some Excel builds require a workbook before this setter works.
        book = app.Workbooks.Open(str(Path(request["candidate"]).resolve()), UpdateLinks=0,
                                  ReadOnly=False, IgnoreReadOnlyRecommended=True,
                                  Notify=False, AddToMru=False, CorruptLoad=0)
        require(not book.ReadOnly, "EXCEL_READ_ONLY", "Excel 只能只读打开临时计算副本")
        try:
            app.Calculation = -4135
        except Exception:
            pass
        geometry = {}
        for item in request.get("geometry", []):
            sheet = book.Worksheets.Item(item["sheet"])
            obj = sheet.Range(item["range"]) if "range" in item else sheet.Shapes.Item(item["shape_name"])
            geometry[item["key"]] = {"width": float(obj.Width), "height": float(obj.Height), "left": float(obj.Left), "top": float(obj.Top)}
            obj = None
        layouts = {}
        if request.get("calculate"):
            for item in request.get("pivots", []):
                pivot = book.Worksheets.Item(item["sheet"]).PivotTables().Item(item["name"])
                layouts[item["sheet"] + "|" + item["name"]] = pivot_layout(pivot)
                pivot = None
            app.CalculateFullRebuild()
            wait_calculation(app, request["timeout"])
            refreshed = set()
            for item in request.get("pivots", []):
                if item["source_type"] != "worksheet":
                    continue  # Prepared external snapshots remain frozen.
                pivot = book.Worksheets.Item(item["sheet"]).PivotTables().Item(item["name"])
                cache = pivot.PivotCache()
                require(int(cache.SourceType) == 1, "PIVOT_SOURCE_UNSUPPORTED", "透视缓存不是工作表数据源")
                index = int(cache.Index)
                if index not in refreshed:
                    cache.Refresh()
                    refreshed.add(index)
                pivot.RefreshTable()
                require(pivot_layout(pivot) == layouts[item["sheet"] + "|" + item["name"]],
                        "PIVOT_LAYOUT_CHANGED", "透视刷新改变了结果范围或字段布局", pivot=item["name"])
                cache = pivot = None
            app.CalculateFullRebuild()
            wait_calculation(app, request["timeout"])
            charts = []
            for item in request.get("charts", []):
                chart = book.Worksheets.Item(item["sheet"]).ChartObjects().Item(item["name"]).Chart
                chart.Refresh()
                collection = chart.SeriesCollection()
                series = []
                for i in range(1, collection.Count + 1):
                    obj = collection.Item(i)
                    series.append({"name": str(obj.Name), "values": serializable(obj.Values), "categories": serializable(obj.XValues)})
                    obj = None
                charts.append({"sheet": item["sheet"], "name": item["name"], "series": series})
                collection = chart = None
            wait_calculation(app, request["timeout"])
            destination = str(Path(request["computed"]).resolve())
            book.SaveCopyAs(destination)
            require(Path(destination).is_file(), "CALCULATION_SAVE_FAILED", "Excel 未生成缓存副本")
            result["charts"] = charts
            result["pivot_layouts"] = layouts
        result.update({"geometry": geometry, "engine": "Microsoft Excel", "version": str(app.Version),
                       "calculated": bool(request.get("calculate")), "external_refresh": False})
        return result
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
            book = None
        if app is not None and owned:
            try:
                app.Quit()
            except Exception:
                pass
        app = None
        gc.collect()
        pythoncom.CoUninitialize()


def main():
    request_path, response_path = map(Path, sys.argv[1:3])
    request = json.loads(request_path.read_text(encoding="utf-8"))
    try:
        response = {"ok": True, "result": run(request, response_path.with_suffix(".owner.json"))}
    except XMLMCPError as exc:
        response = {"ok": False, "error": exc.as_dict()}
    except Exception as exc:
        import traceback
        response = {"ok": False, "error": {"code": "EXCEL_FAILED", "message": "Excel 自动化失败", "details": {"error": str(exc), "traceback": traceback.format_exc()}}}
    response_path.write_text(json.dumps(response, ensure_ascii=False, indent=2), encoding="utf-8")
    if not response["ok"]:
        sys.exit(1)


if __name__ == "__main__":
    main()

