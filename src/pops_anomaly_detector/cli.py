"""Command-line entry point."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .analysis import InputDiscoveryError, analyze_directories
from .config import AnalysisConfig
from .reporting import write_reports


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pops-anomaly",
        description=(
            "Compare paired sent/received POPS .xlsx files and generate auditable "
            "HTML structural-anomaly reports."
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--sent-dir", type=Path, default=Path("data/sent"))
    parser.add_argument("--received-dir", type=Path, default=Path("data/received"))
    parser.add_argument("--output-dir", type=Path, default=Path("reports"))
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search input directories recursively; pairing still uses the filename.",
    )
    parser.add_argument("--max-active-rows", type=_positive_int, default=100_000)
    parser.add_argument("--max-active-columns", type=_positive_int, default=16_384)
    parser.add_argument("--max-cells-per-sheet", type=_positive_int, default=5_000_000)
    parser.add_argument("--alignment-band", type=_positive_int, default=240)
    parser.add_argument(
        "--always-zero",
        action="store_true",
        help="Return exit code 0 after report generation even when files are ERROR.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = AnalysisConfig(
        sent_dir=args.sent_dir,
        received_dir=args.received_dir,
        output_dir=args.output_dir,
        recursive=args.recursive,
        max_active_rows=args.max_active_rows,
        max_active_columns=args.max_active_columns,
        max_cells_per_sheet=args.max_cells_per_sheet,
        alignment_band=args.alignment_band,
    )
    try:
        run = analyze_directories(config)
        output = write_reports(run, config.resolved().output_dir)
    except (InputDiscoveryError, OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    summary = run.summary
    print(f"Global report: {output / 'index.html'}")
    print(
        f"Sent {summary.sent_files} | received {summary.received_files} | "
        f"matched {summary.matched_pairs} | OK {summary.ok} | ERROR {summary.error} | "
        f"HIGH findings {summary.high_findings}"
    )
    if not run.countries:
        print("ERROR: no .xlsx files were found in either input directory.", file=sys.stderr)
        return 2
    if summary.error and not args.always_zero:
        return 1
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

