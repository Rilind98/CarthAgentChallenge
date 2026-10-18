from __future__ import annotations

import argparse
import json
from pathlib import Path

from .pipeline import load_config, run_many


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="carthagent", description="Evidence-first fixed-asset schedule extraction")
    sub = parser.add_subparsers(dest="command", required=True)
    run = sub.add_parser("run", help="process a PDF/XLSX file or input directory")
    run.add_argument("input", metavar="INPUT", type=Path)
    run.add_argument("--output", metavar="OUTPUT", type=Path, required=True)
    run.add_argument("--config", metavar="PATH", type=Path)
    run.add_argument("--template", metavar="PATH", type=Path, help="explicit private workbook template; requires configured destination mapping")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "run":
        config = load_config(args.config, args.template)
        results = run_many(args.input, args.output, config)
        for result in results:
            print(json.dumps({"status": result.status.value, "run_id": result.run_id, "document_id": result.document_id, "source_hash": result.source_hash, "output_dir": str(result.output_dir) if result.output_dir else None, "evidence_path": str(result.evidence_path) if result.evidence_path else None, "review_path": str(result.review_path) if result.review_path else None, "manifest_path": str(result.manifest_path) if result.manifest_path else None, "diagnostics": result.diagnostics, "duration_seconds": round(result.duration_seconds, 3)}, sort_keys=True))
        return 0 if all(r.status.value in {"complete", "partial"} for r in results) else 2
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
