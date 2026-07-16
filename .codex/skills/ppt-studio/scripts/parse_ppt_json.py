#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from _api_client import post_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Parse PPTX structure through PPTStudio API.")
    parser.add_argument("pptx_path", help="Input .pptx file path")
    parser.add_argument("-o", "--output", help="Output JSON path; default is sibling .json", default=None)
    parser.add_argument(
        "--export-media-dir",
        help="Directory to export picture assets; default is <json_stem>_assets beside output JSON",
        default=None,
    )
    parser.add_argument("--no-api", action="store_true", help="Skip font/background/theme merge")
    parser.add_argument("--api-base", help="PPT API base URL", default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    pptx_path = Path(args.pptx_path).expanduser().resolve()
    payload = {
        "file_path": str(pptx_path),
        "output_path": str(Path(args.output).expanduser().resolve()) if args.output else None,
        "export_media_dir": str(Path(args.export_media_dir).expanduser().resolve()) if args.export_media_dir else None,
        "no_api": args.no_api,
    }
    result = post_json("/ppt/parse_structure", payload, api_base=args.api_base)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    data = result.get("data") if isinstance(result, dict) else None
    if isinstance(data, dict) and data.get("output_json_path"):
        print(data["output_json_path"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
