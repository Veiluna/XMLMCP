from __future__ import annotations

import argparse
import json
import logging
import sys


def main():
    parser = argparse.ArgumentParser(description="XMLMCP local stdio server / diagnostics")
    sub = parser.add_subparsers(dest="command")
    inspect = sub.add_parser("inspect")
    inspect.add_argument("path")
    apply = sub.add_parser("apply")
    apply.add_argument("input_path")
    apply.add_argument("output_path")
    apply.add_argument("operations_json", help="UTF-8 JSON operation list file")
    verify = sub.add_parser("verify")
    verify.add_argument("input_path")
    verify.add_argument("output_path")
    verify.add_argument("report_path")
    parser.add_argument("--version", action="store_true")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, stream=sys.stderr, format="%(levelname)s %(name)s: %(message)s")
    from . import __version__
    if args.version:
        print(__version__)
        return
    if args.command is None:
        from .server import server
        server.run(transport="stdio")
        return
    from .errors import XMLMCPError
    from .engine import inspect_template, apply_changes
    from .integrity import verify_integrity
    from pathlib import Path
    try:
        if args.command == "inspect":
            result = inspect_template(args.path)
        elif args.command == "apply":
            operations = json.loads(Path(args.operations_json).read_text(encoding="utf-8"))
            result = apply_changes(args.input_path, args.output_path, operations)
        else:
            result = verify_integrity(args.input_path, args.output_path, args.report_path)
        print(result.model_dump_json(indent=2))
        if args.command == "verify" and not result.valid:
            sys.exit(1)
    except XMLMCPError as exc:
        print(json.dumps({"error": exc.as_dict()}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

