"""Build a source ZIP that installs afresh on Windows or macOS."""
from __future__ import annotations

import hashlib
import json
import tomllib
import zipfile
from pathlib import Path


ROOT_FILES = ("pyproject.toml", "requirements.lock", "README.md", "INSTALL.md", "VALIDATION.md", ".gitignore")
EXAMPLES = ("core-template.xlsx", "core-operations.json", "template.xlsx", "photo.png", "operations.json")


def release_files(root: Path) -> list[Path]:
    # Explicit roots and suffixes keep private workbooks, local config and venvs out.
    files = [root / name for name in ROOT_FILES]
    files += [root / "examples" / name for name in EXAMPLES]
    files += [root / "src/xmlmcp/fonts" / name for name in ("XMLMCPChartLabels.ttf", "XMLMCPChartLabelsBold.ttf", "OFL.txt", "README.md")]
    for directory, suffixes in (("src/xmlmcp", {".py"}), ("scripts", {".py", ".ps1", ".sh"}),
                                ("tests", {".py"}), ("docs", {".md"})):
        files += [p for p in (root / directory).rglob("*")
                  if p.is_file() and p.suffix in suffixes and "__pycache__" not in p.parts]
    for name in ("portable-install.json", "portable-unit-tests.xml", "mac-arm64-deps.json", "mac-x86_64-deps.json", "report-v0.2.0.json", "report-v0.2.2.json"):
        path = root / "validation" / name
        if path.is_file():
            files.append(path)
    missing = [str(p.relative_to(root)) for p in files if not p.is_file()]
    if missing:
        raise FileNotFoundError("Missing release files: " + ", ".join(missing))
    return sorted(set(files))


def build(root: Path) -> dict:
    version = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    files = release_files(root)
    directory = root / "dist"
    directory.mkdir(exist_ok=True)
    output = directory / f"XMLMCP-{version}-portable.zip"
    manifest = {"version": version, "files": {}}
    staging = output.with_suffix(".zip.tmp")
    try:
        with zipfile.ZipFile(staging, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for path in files:
                relative = path.relative_to(root).as_posix()
                data = path.read_bytes()
                if relative == "examples/operations.json":
                    operations = json.loads(data)
                    for operation in operations:
                        if "image_path" in operation:
                            operation["image_path"] = "examples/photo.png"
                    data = (json.dumps(operations, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
                manifest["files"][relative] = hashlib.sha256(data).hexdigest()
                package.writestr("XMLMCP/" + relative, data)
            package.writestr("XMLMCP/RELEASE.json", json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
        staging.replace(output)
    finally:
        staging.unlink(missing_ok=True)
    fingerprint = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix(".zip.sha256").write_text(f"{fingerprint}  {output.name}\n", encoding="ascii")
    return {"path": str(output), "version": version, "sha256": fingerprint,
            "size": output.stat().st_size, "files": len(files) + 1}


if __name__ == "__main__":
    print(json.dumps(build(Path(__file__).absolute().parent.parent), ensure_ascii=False, indent=2))
