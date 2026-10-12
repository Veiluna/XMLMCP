"""Publish a self-contained, reproducible example from a verified native fixture."""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from xmlmcp.engine import apply_changes, inspect_template
from xmlmcp.xmlbytes import digest


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("fixtures", type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    directory = root / "examples"
    directory.mkdir(exist_ok=True)
    source = args.fixtures.resolve()
    for original, name in [(source / "basic-template.xlsx", "template.xlsx"), (source / "replacement.png", "photo.png")]:
        target = directory / name
        if target.exists():
            assert digest(target.read_bytes()) == digest(original.read_bytes()), "An existing example differs; will not overwrite"
        else:
            shutil.copyfile(original, target)
    template = directory / "template.xlsx"
    info = inspect_template(str(template))
    operations = [
        {"op": "write", "target": {"name": "Title", "scope": "workbook"}, "value": "New title"},
        {"op": "write", "target": {"name": "Tiny", "scope": "workbook"}, "value": "Small English preserved"},
        {"op": "write", "target": {"name": "Grid", "scope": "workbook"}, "value": [["甲", 100], ["乙", 200]]},
        {"op": "add_image", "target": {"name": "Photo", "scope": "workbook"}, "image_path": str(directory / "photo.png")},
        {"op": "replace_image", "image_id": info.images[0]["image_id"], "image_path": str(directory / "photo.png")},
    ]
    operations_path = directory / "operations.json"
    text = json.dumps(operations, ensure_ascii=False, indent=2)
    if operations_path.exists():
        assert operations_path.read_text(encoding="utf-8") == text
    else:
        operations_path.write_text(text, encoding="utf-8")
    applied = apply_changes(str(template), str(directory / "result.xlsx"), operations)
    (directory / "apply-result.json").write_text(applied.model_dump_json(indent=2), encoding="utf-8")
    print(applied.model_dump_json(indent=2))


if __name__ == "__main__":
    main()
