import asyncio
import sys

import pytest
from mcp.client import Client
from mcp.client.stdio import StdioServerParameters

from xmlmcp.server import server


@pytest.mark.parametrize("mode", ["legacy", "auto"])
def test_in_memory_protocol_tools_and_errors(template, mode):
    async def run():
        async with Client(server, mode=mode) as client:
            listing = await client.list_tools()
            assert {t.name for t in listing.tools} == {"inspect_template", "apply_changes", "verify_integrity"}
            assert all(t.output_schema for t in listing.tools)
            result = await client.call_tool("inspect_template", {"path": str(template)})
            assert not result.is_error and result.structured_content["fingerprint"]
            assert result.content[0].type == "text"
            failed = await client.call_tool("apply_changes", {"input_path": str(template), "output_path": str(template), "operations": [{"op": "clear", "target": {"sheet": "报告", "range": "C5"}}]})
            assert failed.is_error
            assert failed.structured_content["error"]["code"] == "INPUT_OVERWRITE"
            return client.protocol_version
    version = asyncio.run(run())
    assert version


@pytest.mark.parametrize("mode", ["legacy", "auto"])
def test_real_stdio_process(template, mode):
    async def run():
        params = StdioServerParameters(command=sys.executable, args=["-m", "xmlmcp"])
        async with Client(params, mode=mode, read_timeout_seconds=30) as client:
            listing = await client.list_tools()
            assert len(listing.tools) == 3
            output = str(template.with_name("stdio-result.xlsx"))
            result = await client.call_tool("apply_changes", {"input_path": str(template), "output_path": output,
                                             "operations": [{"op": "write", "target": {"name": "SmallFont"}, "value": "stdio English"}]})
            assert not result.is_error, result.content
            report = result.structured_content["report_path"]
            verified = await client.call_tool("verify_integrity", {"input_path": str(template), "output_path": output, "report_path": report})
            assert not verified.is_error and verified.structured_content["valid"]
    asyncio.run(run())

