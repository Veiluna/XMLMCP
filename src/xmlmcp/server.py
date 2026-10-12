from __future__ import annotations

import json
import logging
from functools import partial

import anyio
from mcp.server import MCPServer
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import ValidationError

from . import __version__
from .errors import XMLMCPError
from .models import Target, Operation, InspectResult, ApplyResult, IntegrityResult
from . import engine
from .integrity import verify_integrity as verify

logger = logging.getLogger("xmlmcp")
server = MCPServer("XMLMCP", version=__version__,
                   description="本机 XLSX 模板 XML 无损注入：原部件保留、批量预检、计算缓存回填、独立完整性验证。",
                   instructions="先 inspect_template 确认目标与 runtime 支持情况，再 apply_changes。Windows/macOS 均支持检查及无需重算的 XML 修改；计算缓存和图片修改需要 Windows Microsoft Excel。所有文件使用绝对路径。输出必须是新路径。失败时不要用其他工作簿库重保存来替代。")


async def invoke(fn, *args):
    try:
        return await anyio.to_thread.run_sync(partial(fn, *args))
    except XMLMCPError as exc:
        error = exc.as_dict()
    except ValidationError as exc:
        error = {"code": "INVALID_OPERATIONS", "message": "操作输入校验失败", "details": {"error": str(exc)}}
    except OSError as exc:
        error = {"code": "FILE_IO", "message": "文件读写失败；未交付输出", "details": {"error": str(exc)}}
    except Exception:
        logger.exception("Unhandled XMLMCP failure")
        error = {"code": "INTERNAL_ERROR", "message": "处理失败；未交付输出，请检查 stderr 日志", "details": {}}
    return CallToolResult(content=[TextContent(type="text", text=json.dumps(error, ensure_ascii=False))],
                          structured_content={"error": error}, is_error=True)


@server.tool(annotations=ToolAnnotations(read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=False), structured_output=True)
async def inspect_template(path: str, targets: list[Target] | None = None) -> InspectResult:
    """只读检查当前 XLSX 模板。返回名称及作用域、指定目标的值与样式、图片 ID 和图表来源；请在修改前调用。path 为绝对路径。"""
    return await invoke(engine.inspect_template, path, targets)


@server.tool(annotations=ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False), structured_output=True)
async def apply_changes(input_path: str, output_path: str, operations: list[Operation]) -> ApplyResult:
    """整批执行 write/formula/clear/copy_style/add_image/replace_image。保留模板，只创建不存在的新 XLSX 和完整性报告；任一失败取消整批。无需重算的 XML 修改可在 Windows/macOS 执行。计算缓存和图片修改需要 Windows Microsoft Excel。write 的 = 开头字符串是文字。"""
    return await invoke(engine.apply_changes, input_path, output_path, operations)


@server.tool(annotations=ToolAnnotations(read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=False), structured_output=True)
async def verify_integrity(input_path: str, output_path: str, report_path: str) -> IntegrityResult:
    """只读复核 apply_changes 生成的报告：逐部件哈希、精确补丁重放、片段外字节、XML/关系/缓存结构。返回 valid；不写文件。三个参数均为绝对路径。"""
    return await invoke(verify, input_path, output_path, report_path)

