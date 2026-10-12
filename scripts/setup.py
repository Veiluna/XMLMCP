"""Install XMLMCP into a local venv and generate command/args for this computer."""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path


def environment_python(directory: Path, windows: bool | None = None) -> Path:
    if windows is None:
        windows = os.name == "nt"
    return directory / ("Scripts/python.exe" if windows else "bin/python")


def client_configuration(python: Path) -> dict:
    # Keep the venv interpreter path: resolving a POSIX symlink would lose its venv.
    return {"mcpServers": {"xmlmcp": {"command": str(python.absolute()), "args": ["-m", "xmlmcp"]}}}


def write_local_configuration(root: Path, python: Path) -> Path:
    path = root / "mcp.local.json"
    path.write_text(json.dumps(client_configuration(python), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    operations = root / "examples" / "operations.json"
    if operations.is_file():
        local = json.loads(operations.read_text(encoding="utf-8"))
        for operation in local:
            if "image_path" in operation:
                operation["image_path"] = str(root / "examples" / "photo.png")
        operations.with_name("operations.local.json").write_text(json.dumps(local, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def run(command: list[str], root: Path, env: dict):
    subprocess.run(command, cwd=root, env=env, check=True)


def main():
    parser = argparse.ArgumentParser(description="Install XMLMCP on Windows or macOS; Python 3.12+ required")
    parser.add_argument("--check-only", action="store_true", help="Check an installed environment and regenerate this computer's MCP configuration")
    args = parser.parse_args()
    if sys.version_info < (3, 12):
        parser.error("XMLMCP requires Python 3.12 or newer")
    root = Path(__file__).absolute().parent.parent
    directory = root / ".venv"
    python = environment_python(directory)
    temporary = root / ".tmp"
    temporary.mkdir(exist_ok=True)
    env = dict(os.environ)
    env.update({"TEMP": str(temporary), "TMP": str(temporary), "TMPDIR": str(temporary),
                "PIP_CACHE_DIR": str(temporary / "pip-cache"), "PYTHONUTF8": "1"})
    if directory.exists() and not python.is_file():
        parser.error("The .venv directory belongs to another platform or is incomplete. Move it aside, then run setup again; do not copy virtual environments between computers.")
    if not python.is_file():
        if args.check_only:
            parser.error("No local environment exists; run setup without --check-only first")
        run([sys.executable, "-m", "venv", str(directory)], root, env)
    run([str(python), "-c", "import sys; assert sys.version_info >= (3,12), 'Python 3.12+ required'"], root, env)
    if not args.check_only:
        run([str(python), "-m", "pip", "install", "-r", str(root / "requirements.lock")], root, env)
        run([str(python), "-m", "pip", "install", "--no-deps", "-e", str(root)], root, env)
    run([str(python), "-m", "pip", "check"], root, env)
    run([str(python), str(root / "scripts" / "check_install.py")], root, env)
    config = write_local_configuration(root, python)
    print(f"\nXMLMCP installed. Import this computer's configuration: {config}")
    print(json.dumps(client_configuration(python), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except (subprocess.CalledProcessError, OSError) as exc:
        print(f"Installation failed: {exc}", file=sys.stderr)
        sys.exit(1)
