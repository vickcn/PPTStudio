#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from _api_client import post_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check shape overlap in parsed PPT layout JSON through PPTStudio API."
    )
    parser.add_argument("json_path", help="Input layout JSON path")
    parser.add_argument(
        "-o",
        "--output",
        help="Output report JSON path; default is <json_stem>_overlap_report.json beside input JSON",
        default=None,
    )
    parser.add_argument("--only-overlaps", action="store_true", help="Only keep slides with overlaps")
    parser.add_argument("--api-base", help="PPT API base URL", default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    json_path = Path(args.json_path).expanduser().resolve()
    payload = {
        "json_path": str(json_path),
        "output_path": str(Path(args.output).expanduser().resolve()) if args.output else None,
        "only_overlaps": args.only_overlaps,
    }
    result = post_json("/ppt/check_overlap", payload, api_base=args.api_base)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    data = result.get("data") if isinstance(result, dict) else None
    if isinstance(data, dict) and data.get("output_report_path"):
        print(data["output_report_path"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
