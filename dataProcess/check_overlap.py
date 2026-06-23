#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def rect_from_shape(shape: dict[str, Any]) -> dict[str, Any] | None:
    geometry = shape.get("geometry") or {}
    left = geometry.get("left_emu")
    top = geometry.get("top_emu")
    width = geometry.get("width_emu")
    height = geometry.get("height_emu")
    if None in (left, top, width, height):
        return None
    left = int(left)
    top = int(top)
    width = int(width)
    height = int(height)
    return {
        "left": left,
        "top": top,
        "right": left + width,
        "bottom": top + height,
        "width": width,
        "height": height,
    }


def overlap_rect(a: dict[str, Any], b: dict[str, Any]) -> dict[str, int] | None:
    left = max(a["left"], b["left"])
    top = max(a["top"], b["top"])
    right = min(a["right"], b["right"])
    bottom = min(a["bottom"], b["bottom"])
    if left >= right or top >= bottom:
        return None
    return {
        "left_emu": left,
        "top_emu": top,
        "right_emu": right,
        "bottom_emu": bottom,
        "width_emu": right - left,
        "height_emu": bottom - top,
        "area_emu2": (right - left) * (bottom - top),
    }


def shape_label(shape: dict[str, Any]) -> str:
    parts = []
    if shape.get("shape_id") is not None:
        parts.append(f"id={shape['shape_id']}")
    if shape.get("name"):
        parts.append(str(shape["name"]))
    if shape.get("shape_type"):
        parts.append(str(shape["shape_type"]))
    elif shape.get("xml_tag"):
        parts.append(str(shape["xml_tag"]))
    return " | ".join(parts)


def analyze_slide(slide: dict[str, Any]) -> dict[str, Any]:
    candidates = []
    for shape in slide.get("shapes", []):
        rect = rect_from_shape(shape)
        if rect is None:
            continue
        candidates.append(
            {
                "shape_index": shape.get("shape_index"),
                "shape_id": shape.get("shape_id"),
                "name": shape.get("name"),
                "shape_type": shape.get("shape_type"),
                "xml_tag": shape.get("xml_tag"),
                "text_preview": (shape.get("text") or "")[:80],
                "geometry": shape.get("geometry") or {},
                "rect": rect,
                "label": shape_label(shape),
            }
        )

    overlaps = []
    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a = candidates[i]
            b = candidates[j]
            intersection = overlap_rect(a["rect"], b["rect"])
            if intersection is None:
                continue
            a_area = a["rect"]["width"] * a["rect"]["height"]
            b_area = b["rect"]["width"] * b["rect"]["height"]
            overlap_area = intersection["area_emu2"]
            smaller_area = min(a_area, b_area) if min(a_area, b_area) else 1
            overlaps.append(
                {
                    "shape_a": {
                        "shape_index": a["shape_index"],
                        "shape_id": a["shape_id"],
                        "name": a["name"],
                        "shape_type": a["shape_type"],
                        "text_preview": a["text_preview"],
                        "geometry": a["geometry"],
                        "rect": a["rect"],
                        "label": a["label"],
                    },
                    "shape_b": {
                        "shape_index": b["shape_index"],
                        "shape_id": b["shape_id"],
                        "name": b["name"],
                        "shape_type": b["shape_type"],
                        "text_preview": b["text_preview"],
                        "geometry": b["geometry"],
                        "rect": b["rect"],
                        "label": b["label"],
                    },
                    "intersection": intersection,
                    "overlap_ratio_vs_smaller_shape": round(overlap_area / smaller_area, 6),
                }
            )

    return {
        "slide_index": slide.get("slide_index"),
        "slide_number": slide.get("slide_number", slide.get("slide_index")),
        "shape_count_with_geometry": len(candidates),
        "overlap_count": len(overlaps),
        "overlaps": overlaps,
    }


def analyze_layout(layout: dict[str, Any]) -> dict[str, Any]:
    slides = [analyze_slide(slide) for slide in layout.get("slides", [])]
    return {
        "file_path": layout.get("file_path"),
        "slide_count": layout.get("slide_count"),
        "source_json_path": layout.get("file_path"),
        "slides_with_overlaps": sum(1 for slide in slides if slide["overlap_count"] > 0),
        "total_overlap_count": sum(slide["overlap_count"] for slide in slides),
        "slides": slides,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check overlap between shapes in parsed PPT layout JSON.")
    parser.add_argument("json_path", help="Input layout JSON path")
    parser.add_argument(
        "-o",
        "--output",
        help="Output report JSON path; default is <json_stem>_overlap_report.json beside input JSON",
        default=None,
    )
    parser.add_argument(
        "--only-overlaps",
        action="store_true",
        help="Only keep slides that have overlaps in the output report",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    json_path = Path(args.json_path).resolve()
    if not json_path.exists():
        raise SystemExit(f"JSON file not found: {json_path}")

    layout = read_json(json_path)
    report = analyze_layout(layout)
    if args.only_overlaps:
        report["slides"] = [slide for slide in report["slides"] if slide["overlap_count"] > 0]

    output_path = (
        Path(args.output).resolve()
        if args.output
        else json_path.with_name(f"{json_path.stem}_overlap_report.json")
    )
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(output_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
