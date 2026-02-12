from __future__ import annotations

import argparse
import json

from . import __version__
from .engine import apply_ops


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="pptx-ooxml-engine CLI")
    parser.add_argument("--version", action="store_true", help="Print package version")
    parser.add_argument("--template", dest="template_pptx", help="Master template PPTX path")
    parser.add_argument("--input", dest="template_pptx", help=argparse.SUPPRESS)
    parser.add_argument("--ops-file", help="Operations JSON file path")
    parser.add_argument("--output", help="Output PPTX path")
    parser.add_argument("--verify", action="store_true", help="Run OOXML verification after apply")
    parser.add_argument(
        "--no-strict-verify",
        action="store_true",
        help="Do not fail command when verifier reports issues",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.version:
        print(__version__)
        return 0
    if not args.ops_file or not args.output:
        parser.print_help()
        return 0

    with open(args.ops_file, "r", encoding="utf-8") as f:
        raw_ops = json.load(f)

    result = apply_ops(
        template_pptx=args.template_pptx,
        input_pptx=None,
        ops=raw_ops,
        output_pptx=args.output,
        verify=args.verify,
        strict_verify=not args.no_strict_verify,
    )
    print(result.output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
