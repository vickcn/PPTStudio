#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from _api_client import get_json, post_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render rebuilt PPTX slides to PNG screenshots through PPTStudio API."
    )
    parser.add_argument("pptx_path", help="Input .pptx file path")
    parser.add_argument(
        "-o",
        "--output-dir",
        help="Directory for PNG screenshots; default is <pptx_stem>_audit beside the pptx",
        default=None,
    )
    parser.add_argument("--dpi", type=int, default=150, help="Render DPI (default: 150)")
    parser.add_argument("--grid", action="store_true", help="Also render a multi-slide grid image")
    parser.add_argument("--cols", type=int, default=2, help="Grid columns when --grid is set")
    parser.add_argument("--api-base", help="PPT API base URL", default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    pptx_path = Path(args.pptx_path).expanduser().resolve()
    output_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else pptx_path.with_name(f"{pptx_path.stem}_audit")
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    info = get_json("/ppt/info", {"file_path": str(pptx_path)}, api_base=args.api_base)
    slide_count = 0
    if isinstance(info, dict):
        data = info.get("data") or {}
        nested = data.get("info") or data
        slide_count = int(nested.get("slide_count") or nested.get("slides") or 0)

    rendered = []
    for slide_index in range(slide_count):
        output_path = output_dir / f"slide_{slide_index + 1:02d}.png"
        payload = {
            "file_path": str(pptx_path),
            "slide_index": slide_index,
            "output_path": str(output_path),
            "dpi": args.dpi,
        }
        result = post_json("/ppt/render_slide_to_image", payload, api_base=args.api_base)
        rendered.append(
            {
                "slide_index": slide_index,
                "slide_number": slide_index + 1,
                "output_path": str(output_path),
                "ok": bool(result.get("ok")),
            }
        )

    grid_path = None
    if args.grid and slide_count > 0:
        grid_path = output_dir / f"{pptx_path.stem}_slides_grid.png"
        payload = {
            "file_path": str(pptx_path),
            "slide_indices": list(range(slide_count)),
            "output_path": str(grid_path),
            "cols": args.cols,
            "dpi": args.dpi,
            "add_page_title": True,
        }
        post_json("/ppt/render_slides_to_grid_image", payload, api_base=args.api_base)

    summary = {
        "pptx_path": str(pptx_path),
        "output_dir": str(output_dir),
        "slide_count": slide_count,
        "rendered": rendered,
        "grid_path": str(grid_path) if grid_path else None,
    }
    manifest_path = output_dir / "screenshot_manifest.json"
    manifest_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(str(output_dir))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
