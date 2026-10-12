"""Check core XML operations and actual stdio tools without starting Office."""
from __future__ import annotations

import asyncio
import json
import sys
import tempfile
from pathlib import Path

from mcp.client import Client
from mcp.client.stdio import StdioServerParameters
from xmlmcp import __version__
from xmlmcp.namespaces import S
from xmlmcp.package import Package
from xmlmcp.runtime import capabilities
from xmlmcp.workbook import Workbook


async def check(root: Path) -> dict:
    template = root / "examples" / "core-template.xlsx"
    before = Package(template)
    source = Workbook(before)
    assert not source.charts() and not source.pivots()
    assert all(not c.child("f", S) for sheet in source.sheets for c in source.cell_map(sheet).values())
    operations = json.loads((root / "examples/core-operations.json").read_text(encoding="utf-8"))
    temp = root / ".tmp"
    temp.mkdir(exist_ok=True)
    checks = []
    with tempfile.TemporaryDirectory(prefix="xmlmcp-install-", dir=temp) as directory:
        for mode in ("legacy", "auto"):
            output = Path(directory) / (mode + ".xlsx")
            parameters = StdioServerParameters(command=sys.executable, args=["-m", "xmlmcp"])
            async with Client(parameters, mode=mode, read_timeout_seconds=30) as client:
                listing = await client.list_tools()
                assert {t.name for t in listing.tools} == {"inspect_template", "apply_changes", "verify_integrity"}
                inspected = await client.call_tool("inspect_template", {"path": str(template)})
                assert not inspected.is_error, inspected.content
                applied = await client.call_tool("apply_changes", {"input_path": str(template), "output_path": str(output), "operations": operations})
                assert not applied.is_error, applied.content
                result = applied.structured_content
                assert not result["calculation"]["calculated"]
                verified = await client.call_tool("verify_integrity", {"input_path": str(template), "output_path": str(output), "report_path": result["report_path"]})
                assert not verified.is_error and verified.structured_content["valid"], verified.content
                produced = Package(output)
                assert set(before.parts) == set(produced.parts)
                changed = [part for part in before.parts if before.parts[part] != produced.parts[part]]
                assert changed == [source.sheets[0]["part"]], changed
                workbook = Workbook(produced)
                cells = workbook.cell_map(workbook.sheets[0])
                assert workbook.value(cells["A1"]) == "XMLMCP core example"
                assert workbook.value(cells["C5"]) == 100 and workbook.value(cells["C6"]) == 200
                checks.append({"protocol": client.protocol_version, "three_tools_passed": True,
                               "changed_parts": changed, "integrity": True})
    return {"version": __version__, "runtime": capabilities(), "checks": checks, "core_smoke": "passed"}


if __name__ == "__main__":
    print(json.dumps(asyncio.run(check(Path(__file__).absolute().parent.parent)), ensure_ascii=False, indent=2))
