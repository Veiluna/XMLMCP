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
    report_inspect = sub.add_parser("report-inspect", help="Inspect order + raw ZIP + report template")
    report_inspect.add_argument("template_path")
    report_inspect.add_argument("order_path")
    report_inspect.add_argument("raw_zip_path")
    report_inspect.add_argument("--profile")
    report_generate = sub.add_parser("report-generate", help="Generate a mapped report without Excel")
    report_generate.add_argument("template_path")
    report_generate.add_argument("order_path")
    report_generate.add_argument("raw_zip_path")
    report_generate.add_argument("profile_path")
    report_generate.add_argument("output_directory")
    report_generate.add_argument("--metadata", help="UTF-8 JSON object with supplemental report fields")
    prepare = sub.add_parser('prepare-report-template', help='Explicitly author editable tables in a new electrode mother template')
    prepare.add_argument('base_path')
    prepare.add_argument('reference_path')
    prepare.add_argument('output_directory')
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
        elif args.command == "verify":
            result = verify_integrity(args.input_path, args.output_path, args.report_path)
        elif args.command == "report-inspect":
            from .reports import inspect_report_inputs
            result = inspect_report_inputs(args.template_path, args.order_path, args.raw_zip_path, args.profile)
        elif args.command == 'prepare-report-template':
            from .report_template import prepare_electrode_template
            result = prepare_electrode_template(args.base_path, args.reference_path, args.output_directory)
        else:
            from .reports import generate_report
            metadata = json.loads(Path(args.metadata).read_text(encoding="utf-8-sig")) if args.metadata else None
            result = generate_report(args.template_path, args.order_path, args.raw_zip_path, args.output_directory, args.profile_path, metadata)
        print(result.model_dump_json(indent=2) if hasattr(result, 'model_dump_json') else json.dumps(result,ensure_ascii=False,indent=2))
        if args.command == "verify" and not result.valid:
            sys.exit(1)
    except XMLMCPError as exc:
        print(json.dumps({"error": exc.as_dict()}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

