from __future__ import annotations

import json
import os
import runpy
from pathlib import Path
from types import SimpleNamespace

import pytest
from packaging.requirements import Requirement

from conftest import write_template
from xmlmcp import engine, runtime
from xmlmcp.errors import XMLMCPError
from xmlmcp.integrity import verify_integrity
from xmlmcp.package import Package

ROOT = Path(__file__).resolve().parents[1]
SETUP = runpy.run_path(str(ROOT / "scripts/setup.py"))


def test_platform_dependency_lock():
    requirements = [Requirement(line) for line in (ROOT / "requirements.lock").read_text().splitlines() if line]
    windows = {r.name for r in requirements if r.marker is None or r.marker.evaluate({"sys_platform": "win32", "platform_machine": "AMD64"})}
    mac = {r.name for r in requirements if r.marker is None or r.marker.evaluate({"sys_platform": "darwin", "platform_machine": "arm64"})}
    assert windows - mac == {"pywin32"}
    assert not mac - windows and {"mcp", "pydantic", "pillow"} <= mac
    for platform, machine, version in (("win32", "AMD64", "50.0.2"), ("darwin", "arm64", "50.0.2"), ("darwin", "x86_64", "48.0.1")):
        selected = [r for r in requirements if r.name == "cryptography" and r.marker.evaluate({"sys_platform": platform, "platform_machine": machine})]
        assert len(selected) == 1 and str(selected[0].specifier) == "==" + version


def test_relocated_install_generates_local_paths(tmp_path, monkeypatch):
    root = tmp_path / "另一台电脑 XMLMCP"
    examples = root / "examples"
    examples.mkdir(parents=True)
    (examples / "operations.json").write_text(json.dumps([
        {"op": "add_image", "image_path": "G:/old-computer/photo.png"}]), encoding="utf-8")
    for windows, suffix in ((True, "Scripts/python.exe"), (False, "bin/python")):
        python = SETUP["environment_python"](root / ".venv", windows)
        assert python == root / ".venv" / suffix
        # POSIX venv executables are symlinks. Resolving them in command loses the venv.
        with monkeypatch.context() as context:
            context.setattr(Path, "resolve", lambda *a, **k: pytest.fail("Must retain the venv interpreter path"))
            configuration = SETUP["write_local_configuration"](root, python)
        server = json.loads(configuration.read_text(encoding="utf-8"))["mcpServers"]["xmlmcp"]
        assert server == {"command": str(python.absolute()), "args": ["-m", "xmlmcp"]}
        local = json.loads((examples / "operations.local.json").read_text(encoding="utf-8"))
        assert local[0]["image_path"] == str(examples / "photo.png")


def test_mac_core_modifications_keep_existing_objects(tmp_path, monkeypatch):
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="darwin"))
    template = write_template(tmp_path / "template.xlsx", image=True)
    before = Package(template)
    inspection = engine.inspect_template(str(template))
    assert inspection.capabilities["runtime"]["core_xml_available"]
    assert not inspection.capabilities["calculation_available"]
    assert not inspection.capabilities["runtime"]["image_mutation_available"]
    result = engine.apply_changes(str(template), str(tmp_path / "result.xlsx"), [
        {"op": "write", "target": {"name": "Grid"}, "value": [["新数据", 100], ["=文字", True]]},
        {"op": "copy_style", "source": {"sheet": "报告", "range": "C5"}, "target": {"sheet": "报告", "range": "C6"}},
        {"op": "clear", "target": {"sheet": "报告", "range": "A3"}}])
    assert not result.calculation["calculated"]
    assert result.integrity["changed_parts"] == ["xl/worksheets/sheet1.xml"]
    after = Package(result.output_path)
    assert all(after.parts[p] == data for p, data in before.parts.items() if p != "xl/worksheets/sheet1.xml")
    assert verify_integrity(str(template), result.output_path, result.report_path).valid


@pytest.mark.parametrize("kind", ["existing_formula", "new_formula", "chart", "add_image", "replace_image"])
def test_mac_required_excel_operations_cancel_batch(tmp_path, monkeypatch, kind):
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="darwin"))
    template = write_template(tmp_path / "template.xlsx", formula=kind == "existing_formula",
                              chart=kind == "chart", image=kind == "replace_image")
    fingerprint = Package(template).fingerprint
    operations = [{"op": "write", "target": {"sheet": "报告", "range": "C5"}, "value": 100}]
    if kind == "new_formula":
        operations += [{"op": "formula", "target": {"sheet": "报告", "range": "D5"}, "formula": "=C5*2"}]
    elif kind in ("add_image", "replace_image"):
        from PIL import Image
        image = tmp_path / "new.png"
        Image.new("RGB", (20, 20)).save(image)
        target = {"target": {"name": "Photo"}} if kind == "add_image" else {
            "image_id": engine.inspect_template(str(template)).images[0]["image_id"]}
        operations += [{"op": kind, **target, "image_path": str(image)}]
    output = tmp_path / "blocked.xlsx"
    with pytest.raises(XMLMCPError) as caught:
        engine.apply_changes(str(template), str(output), operations)
    assert caught.value.code == "EXCEL_REQUIRED"
    assert not output.exists() and not output.with_suffix(".integrity.json").exists()
    assert not list(tmp_path.glob(".xmlmcp-*"))
    assert Package(template).fingerprint == fingerprint


def test_windows_without_excel_keeps_core_available(template, monkeypatch):
    monkeypatch.setattr(runtime, "sys", SimpleNamespace(platform="win32"))
    monkeypatch.setattr(runtime.importlib.util, "find_spec", lambda name: None)
    assert not runtime.capabilities()["calculation_available"]
    result = engine.apply_changes(str(template), str(template.with_name("result.xlsx")), [
        {"op": "write", "target": {"sheet": "报告", "range": "C5"}, "value": 88}])
    assert not result.calculation["calculated"] and result.integrity["valid"]


def test_posix_publish_succeeds_and_refuses_existing_file(tmp_path, monkeypatch):
    monkeypatch.setattr(engine, "os", SimpleNamespace(name="posix", link=os.link))
    source, destination = tmp_path / "new.tmp", tmp_path / "published.xlsx"
    source.write_bytes(b"first")
    engine._publish_noclobber(source, destination)
    assert not source.exists() and destination.read_bytes() == b"first"
    source.write_bytes(b"second")
    with pytest.raises(FileExistsError):
        engine._publish_noclobber(source, destination)
    assert source.read_bytes() == b"second" and destination.read_bytes() == b"first"


def test_portable_release_excludes_machine_and_business_files():
    builder = runpy.run_path(str(ROOT / "scripts/build_portable.py"))
    paths = {str(p.relative_to(ROOT)).replace("\\", "/") for p in builder["release_files"](ROOT)}
    assert {"scripts/setup.py", "scripts/setup.sh", "examples/core-template.xlsx", "src/xmlmcp/runtime.py"} <= paths
    assert not any(p.startswith((".venv/", ".tmp/", "outputs/", "test-artifacts/")) for p in paths)
    assert not any("__pycache__" in p or ".egg-info" in p or p.endswith(".pyc") for p in paths)
    assert "doubao.mcp.json" not in paths and "mcp.local.json" not in paths
    assert "examples/operations.local.json" not in paths
