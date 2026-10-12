"""Exercise the exact command/args configuration using both SDK protocol eras."""
import asyncio
import argparse
import json
import sys
import uuid
from pathlib import Path

from mcp.client import Client
from mcp.client.stdio import StdioServerParameters


async def run():
    parser = argparse.ArgumentParser()
    parser.add_argument("template", type=Path)
    parser.add_argument("--apply", action="store_true", help="Also generate and verify a native Office fixture through MCP")
    args = parser.parse_args()
    template = str(args.template.resolve())
    directory = args.template.resolve().parent / ("protocol-" + uuid.uuid4().hex[:10]) if args.apply else None
    if directory:
        directory.mkdir()
    results = []
    for mode in ("legacy", "auto"):
        params = StdioServerParameters(command=sys.executable, args=["-m", "xmlmcp"])
        async with Client(params, mode=mode, read_timeout_seconds=120 if args.apply else 30) as client:
            listing = await client.list_tools()
            result = await client.call_tool("inspect_template", {"path": template})
            assert not result.is_error, result.content
            entry = {"mode": mode, "protocol": client.protocol_version,
                     "tools": [t.name for t in listing.tools], "fingerprint": result.structured_content["fingerprint"]}
            assert set(entry["tools"]) == {"inspect_template", "apply_changes", "verify_integrity"}
            assert result.content and result.structured_content
            if directory:
                generated = await client.call_tool("apply_changes", {
                    "input_path": template, "output_path": str(directory / (mode + ".xlsx")),
                    "operations": [{"op": "write", "target": {"name": "Grid", "scope": "workbook"}, "value": [["甲", 100], ["乙", 200]]}],
                })
                assert not generated.is_error, generated.content
                entry["apply"] = generated.structured_content
                verification = await client.call_tool("verify_integrity", {
                    "input_path": template, "output_path": entry["apply"]["output_path"], "report_path": entry["apply"]["report_path"],
                })
                assert not verification.is_error and verification.structured_content["valid"], verification.content
                entry["verify"] = verification.structured_content
            results.append(entry)
            print(json.dumps(entry, ensure_ascii=False), flush=True)
    if directory:
        report = directory / "acceptance.json"
        report.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        print(report)


if __name__ == "__main__":
    asyncio.run(run())

