"""Report platform capabilities without starting Excel or changing user settings."""
from __future__ import annotations

import importlib.util
import sys

from .errors import XMLMCPError


def capabilities() -> dict:
    reason = None
    if sys.platform != "win32":
        reason = "当前平台支持检查和无需重算的 XML 修改；计算缓存和图片修改仍需要 Windows Microsoft Excel"
    elif importlib.util.find_spec("win32com") is None:
        reason = "缺少 Windows Excel 自动化依赖；请重新运行安装脚本"
    else:
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
                pass
        except OSError:
            reason = "未检测到注册的 Microsoft Excel 桌面版；仅安装 WPS 不能提供当前计算后端"
    return {"platform": sys.platform, "core_xml_available": True,
            "calculation_available": reason is None, "image_mutation_available": reason is None,
            "calculation_backend": "windows_excel" if reason is None else None, "limitation": reason}


def require_excel_backend() -> None:
    runtime = capabilities()
    if not runtime["calculation_available"]:
        raise XMLMCPError("EXCEL_REQUIRED", runtime["limitation"], platform=runtime["platform"])
