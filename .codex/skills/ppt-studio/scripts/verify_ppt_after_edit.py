#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from _api_client import get_json, post_json
from _wrap_checks import analyze_layout_wraps

SLIDE_WIDTH_EMU = 9144000
SLIDE_HEIGHT_EMU = 6858000


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Post-edit QA for PPTStudio decks: overlap, line-wrap, and screenshot checks."
    )
    parser.add_argument("pptx_path", help="Rebuilt .pptx path")
    parser.add_argument(
        "--json",
        dest="json_path",
        help="Parsed layout JSON path; default is sibling <pptx_stem>.json",
        default=None,
    )
    parser.add_argument(
        "--audit-dir",
        help="Audit output directory; default is <pptx_stem>_audit beside the pptx",
        default=None,
    )
    parser.add_argument(
        "--min-overlap-ratio",
        type=float,
        default=0.01,
        help="Minimum overlap ratio on the smaller shape to flag text/picture issues",
    )
    parser.add_argument("--dpi", type=int, default=150, help="Screenshot render DPI")
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
    parser.add_argument("--skip-screenshots", action="store_true", help="Only run overlap checks")
    parser.add_argument("--api-base", help="PPT API base URL", default=None)
    return parser.parse_args()


def is_text_shape(shape: Dict[str, Any]) -> bool:
    shape_type = str(shape.get("shape_type") or "").lower()
    if shape_type == "picture":
        return False
    text = str(shape.get("text_preview") or shape.get("text") or "").strip()
    if not text:
        return False
    return shape_type == "text_or_auto_shape" or "text" in shape_type


def is_picture_shape(shape: Dict[str, Any]) -> bool:
    shape_type = str(shape.get("shape_type") or "").lower()
    if shape_type == "picture":
        return True
    return str(shape.get("xml_tag") or "").lower() == "pic"


def is_fullslide_picture(shape: Dict[str, Any]) -> bool:
    geometry = shape.get("geometry") or {}
    width = int(geometry.get("width_emu") or 0)
    height = int(geometry.get("height_emu") or 0)
    if width <= 0 or height <= 0:
        return False
    slide_area = SLIDE_WIDTH_EMU * SLIDE_HEIGHT_EMU
    return (width * height) / slide_area >= 0.9


def is_text_picture_overlap(item: Dict[str, Any], min_ratio: float) -> bool:
    shape_a = item.get("shape_a") or {}
    shape_b = item.get("shape_b") or {}
    ratio = float(item.get("overlap_ratio_vs_smaller_shape") or 0.0)
    if ratio < min_ratio:
        return False

    text_picture = (
        is_text_shape(shape_a) and is_picture_shape(shape_b) and not is_fullslide_picture(shape_b)
    ) or (
        is_text_shape(shape_b) and is_picture_shape(shape_a) and not is_fullslide_picture(shape_a)
    )
    return text_picture


def collect_text_picture_issues(report: Dict[str, Any], min_ratio: float) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    for slide in report.get("slides", []):
        for item in slide.get("overlaps", []):
            if not is_text_picture_overlap(item, min_ratio):
                continue
            issues.append(
                {
                    "slide_index": slide.get("slide_index"),
                    "slide_number": slide.get("slide_number", slide.get("slide_index")),
                    "overlap_ratio_vs_smaller_shape": item.get("overlap_ratio_vs_smaller_shape"),
                    "text_shape": shape_brief(item, is_text_shape),
                    "picture_shape": shape_brief(item, is_picture_shape),
                }
            )
    return issues


def shape_brief(item: Dict[str, Any], matcher) -> Optional[Dict[str, Any]]:
    for key in ("shape_a", "shape_b"):
        shape = item.get(key) or {}
        if matcher(shape):
            return {
                "name": shape.get("name"),
                "shape_id": shape.get("shape_id"),
                "text_preview": shape.get("text_preview"),
                "geometry": shape.get("geometry"),
            }
    return None


def main() -> int:
    args = parse_args()
    pptx_path = Path(args.pptx_path).expanduser().resolve()
    json_path = (
        Path(args.json_path).expanduser().resolve()
        if args.json_path
        else pptx_path.with_suffix(".json")
    )
    audit_dir = (
        Path(args.audit_dir).expanduser().resolve()
        if args.audit_dir
        else pptx_path.with_name(f"{pptx_path.stem}_audit")
    )
    audit_dir.mkdir(parents=True, exist_ok=True)

    overlap_result = post_json(
        "/ppt/check_overlap",
        {
            "json_path": str(json_path),
            "output_path": str(audit_dir / "overlap_report.json"),
            "only_overlaps": False,
        },
        api_base=args.api_base,
    )
    overlap_data = overlap_result.get("data") if isinstance(overlap_result, dict) else {}
    overlap_report_path = Path(overlap_data.get("output_report_path") or audit_dir / "overlap_report.json")
    overlap_report = json.loads(overlap_report_path.read_text(encoding="utf-8-sig"))
    layout = json.loads(json_path.read_text(encoding="utf-8-sig"))
    text_picture_issues = collect_text_picture_issues(overlap_report, args.min_overlap_ratio)
    wrap_report = analyze_layout_wraps(
        layout,
        min_stack_gap_in=args.min_stack_gap_in,
        wrap_tolerance_ratio=args.wrap_tolerance_ratio,
    )
    wrap_report_path = audit_dir / "line_wrap_report.json"
    wrap_report_path.write_text(json.dumps(wrap_report, ensure_ascii=False, indent=2), encoding="utf-8")

    screenshot_summary = None
    if not args.skip_screenshots:
        info = get_json("/ppt/info", {"file_path": str(pptx_path)}, api_base=args.api_base)
        slide_count = 0
        if isinstance(info, dict):
            data = info.get("data") or {}
            nested = data.get("info") or data
            slide_count = int(nested.get("slide_count") or nested.get("slides") or 0)

        rendered = []
        for slide_index in range(slide_count):
            output_path = audit_dir / f"slide_{slide_index + 1:02d}.png"
            post_json(
                "/ppt/render_slide_to_image",
                {
                    "file_path": str(pptx_path),
                    "slide_index": slide_index,
                    "output_path": str(output_path),
                    "dpi": args.dpi,
                },
                api_base=args.api_base,
            )
            rendered.append(str(output_path))

        if slide_count > 0:
            grid_path = audit_dir / f"{pptx_path.stem}_slides_grid.png"
            post_json(
                "/ppt/render_slides_to_grid_image",
                {
                    "file_path": str(pptx_path),
                    "slide_indices": list(range(slide_count)),
                    "output_path": str(grid_path),
                    "cols": 2,
                    "dpi": args.dpi,
                    "add_page_title": True,
                },
                api_base=args.api_base,
            )
        else:
            grid_path = None

        screenshot_summary = {
            "output_dir": str(audit_dir),
            "slide_count": slide_count,
            "rendered": rendered,
            "grid_path": str(grid_path) if slide_count > 0 else None,
        }

    text_picture_issue_count = len(text_picture_issues)
    line_wrap_issue_count = int(wrap_report.get("line_wrap_issue_count") or 0)
    passed = text_picture_issue_count == 0 and line_wrap_issue_count == 0
    summary = {
        "ok": passed,
        "pptx_path": str(pptx_path),
        "json_path": str(json_path),
        "audit_dir": str(audit_dir),
        "overlap_report_path": str(overlap_report_path),
        "line_wrap_report_path": str(wrap_report_path),
        "text_picture_issue_count": text_picture_issue_count,
        "text_picture_issues": text_picture_issues,
        "wrap_overflow_issue_count": wrap_report.get("wrap_overflow_issue_count", 0),
        "stack_gap_issue_count": wrap_report.get("stack_gap_issue_count", 0),
        "line_wrap_issue_count": line_wrap_issue_count,
        "wrap_overflow_issues": wrap_report.get("wrap_overflow_issues", []),
        "stack_gap_issues": wrap_report.get("stack_gap_issues", []),
        "line_wrap_issues": wrap_report.get("line_wrap_issues", []),
        "screenshots": screenshot_summary,
        "notes": [
            "Assumes static deck with no slide animations.",
            "Text/picture overlap is evaluated from parsed JSON geometry, not animation states.",
            "Full-slide background pictures are excluded from text/picture overlap checks.",
            "Line-wrap checks estimate rendered text height from font size, box width, and text length.",
            "Stack-gap checks ensure a lower text box starts below the estimated bottom of the upper text box.",
        ],
    }
    summary_path = audit_dir / "verify_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(summary_path))
    print(
        "ok={0} text_picture_issues={1} line_wrap_issues={2}".format(
            passed,
            text_picture_issue_count,
            line_wrap_issue_count,
        )
    )
    return 0 if passed else 2


if __name__ == "__main__":
    raise SystemExit(main())
