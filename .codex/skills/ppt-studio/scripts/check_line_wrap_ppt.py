#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from _wrap_checks import analyze_layout_wraps


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check estimated line-wrap overflow and text stack gaps in parsed PPT layout JSON."
    )
    parser.add_argument("json_path", help="Input layout JSON path")
    parser.add_argument(
        "-o",
        "--output",
        help="Output report JSON path; default is <json_stem>_line_wrap_report.json beside input JSON",
        default=None,
    )
    parser.add_argument(
        "--min-stack-gap-in",
        type=float,
        default=0.12,
        help="Minimum vertical gap required between stacked text boxes",
    )
    parser.add_argument(
        "--wrap-tolerance-ratio",
        type=float,
        default=1.05,
        help="Allowed ratio of estimated text height over box height",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    json_path = Path(args.json_path).expanduser().resolve()
    layout = json.loads(json_path.read_text(encoding="utf-8-sig"))
    report = analyze_layout_wraps(
        layout,
        min_stack_gap_in=args.min_stack_gap_in,
        wrap_tolerance_ratio=args.wrap_tolerance_ratio,
    )
    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else json_path.with_name(f"{json_path.stem}_line_wrap_report.json")
    )
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(output_path))
    print(
        "line_wrap_issues={0} wrap_overflow={1} stack_gap={2}".format(
            report.get("line_wrap_issue_count", 0),
            report.get("wrap_overflow_issue_count", 0),
            report.get("stack_gap_issue_count", 0),
        )
    )
    return 0 if int(report.get("line_wrap_issue_count") or 0) == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
