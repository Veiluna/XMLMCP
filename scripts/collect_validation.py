"""Assemble evidence without rerunning or replacing the native Office artifacts."""
from __future__ import annotations

import argparse
import compileall
import json
import platform
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree

from PIL import Image
from xmlmcp.integrity import verify_integrity
from xmlmcp.xmlbytes import digest


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("office", type=Path)
    parser.add_argument("regression", type=Path)
    parser.add_argument("protocol", type=Path)
    parser.add_argument("--unit-report", type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    output = root / "validation"
    output.mkdir(exist_ok=True)
    unit_report = (args.unit_report or output / "unit-tests.xml").resolve()
    suites = list(ElementTree.parse(unit_report).getroot().iter("testsuite"))
    totals = {key: sum(int(s.get(key, 0)) for s in suites) for key in ("tests", "failures", "errors", "skipped")}
    assert totals["tests"] > 0 and totals["failures"] == totals["errors"] == 0, totals
    dependency_check = subprocess.run([sys.executable, "-m", "pip", "check"], capture_output=True, text=True, check=True)
    assert compileall.compile_dir(root / "src", quiet=2)
    office_entries = load(args.office.resolve())
    assert len(office_entries) == 4 and all(e["passed"] for e in office_entries)
    office_checks = []
    for entry in office_entries:
        applied = entry["apply"]
        check = verify_integrity(entry["template"], applied["output_path"], applied["report_path"])
        assert check.valid
        renders = []
        for engine in ("excel", "wps"):
            path = Path(applied["output_path"]).with_name(f"{Path(applied['output_path']).stem}-{engine}.png")
            with Image.open(path) as image:
                size = image.size
                image.verify()
            renders.append({"engine": engine, "path": str(path), "size": size, "sha256": digest(path.read_bytes())})
        office_checks.append({"kind": entry["kind"], "integrity": check.model_dump(),
                              "excel": entry["excel"], "wps": entry["wps"], "renders": renders})
    regression = load(args.regression.resolve())
    assert regression["passed"]
    protocols = load(args.protocol.resolve())
    assert {r["protocol"] for r in protocols} == {"2025-11-25", "2026-07-28"}
    assert all(r["verify"]["valid"] for r in protocols)
    example = verify_integrity(str(root / "examples/template.xlsx"), str(root / "examples/result.xlsx"),
                               str(root / "examples/result.integrity.json"))
    assert example.valid
    source_hashes = {str(path.relative_to(root)): digest(path.read_bytes()) for path in sorted((root / "src/xmlmcp").glob("*.py"))}
    evidence = {
        "created_utc": datetime.now(timezone.utc).isoformat(), "python": platform.python_version(),
        "source_sha256": source_hashes,
        "unit_tests": {"passed": totals["tests"] - totals["skipped"], "failed": totals["failures"], "errors": totals["errors"],
                       "skipped": totals["skipped"], "junit_report": str(unit_report), "junit_sha256": digest(unit_report.read_bytes())},
        "installation": {"pip_check": dependency_check.stdout.strip(), "compileall": "passed"},
        "office": office_checks, "office_regression": regression, "protocols": protocols, "example_integrity": example.model_dump(),
        "evidence_files": {"office": str(args.office.resolve()), "regression": str(args.regression.resolve()), "protocol": str(args.protocol.resolve())},
        "pending": {
            "doubao_native_client": "SDK command/args stdio passed; native Doubao settings cannot be controlled with available tools",
            "original_anonymized_business_template": "No user-designated original anonymized template was supplied",
            "interactive_office_prompts": "Native API checks used hidden instances; human confirmation of absence of repair prompts remains pending",
        },
    }
    path = output / "acceptance.json"
    path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
