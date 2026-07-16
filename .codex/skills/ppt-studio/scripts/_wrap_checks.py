# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

LIST_LINE_RE = re.compile(r"^\d+\.\s")


def shape_has_text(shape: Dict[str, Any]) -> bool:
    text = str(shape.get("text") or "").strip()
    return bool(text)


def is_list_style_text(shape: Dict[str, Any]) -> bool:
    text = str(shape.get("text") or "").replace("\x0b", "\n")
    if "\n" in text:
        return True
    paragraphs = shape.get("paragraphs") or []
    if isinstance(paragraphs, list) and len(paragraphs) > 1:
        return True
    for line in normalized_text_lines(text):
        if LIST_LINE_RE.match(line):
            return True
        if line and line[0] in "-•·※●○":
            return True
    return False


def get_font_size_pt(shape: Dict[str, Any]) -> float:
    summary = shape.get("font_summary") or {}
    sizes = summary.get("font_sizes_pt") or []
    if sizes:
        return float(max(sizes))

    max_size = 0.0
    for para in shape.get("paragraphs", []):
        for run in para.get("runs", []):
            size = run.get("font_size_pt") or run.get("effective_font_size_pt")
            if size:
                max_size = max(max_size, float(size))
    if max_size > 0:
        return max_size

    detail = shape.get("font_detail") or {}
    for para in detail.get("paragraphs", []):
        for run in para.get("runs", []):
            size = run.get("font_size_pt") or run.get("effective_font_size_pt")
            if size:
                max_size = max(max_size, float(size))
    return max_size if max_size > 0 else 16.0


def normalized_text_lines(text: str) -> List[str]:
    normalized = text.replace("\x0b", "\n").replace("\r\n", "\n").replace("\r", "\n")
    lines: List[str] = []
    for paragraph in normalized.split("\n"):
        stripped = paragraph.strip()
        if stripped:
            lines.append(stripped)
    return lines


def estimate_wrapped_line_count(text: str, width_in: float, font_pt: float) -> int:
    if width_in <= 0 or font_pt <= 0:
        return 1

    # Mixed Chinese/Latin heuristic for deck titles and body copy.
    chars_per_line = max(4, int(width_in * 72.0 / (font_pt * 0.58)))
    total = 0
    for line in normalized_text_lines(text):
        total += max(1, (len(line) + chars_per_line - 1) // chars_per_line)
    return max(1, total)


def estimate_text_height_in(line_count: int, font_pt: float) -> float:
    line_height_in = font_pt * 1.28 / 72.0
    return line_count * line_height_in


def shape_geometry(shape: Dict[str, Any]) -> Dict[str, float]:
    geometry = shape.get("geometry") or {}
    return {
        "left_in": float(geometry.get("left_in") or 0.0),
        "top_in": float(geometry.get("top_in") or 0.0),
        "width_in": float(geometry.get("width_in") or 0.0),
        "height_in": float(geometry.get("height_in") or 0.0),
    }


def is_page_number_shape(shape: Dict[str, Any]) -> bool:
    name = str(shape.get("name") or "")
    text = str(shape.get("text") or "").strip()
    if name == "TextBox 1" and text.isdigit() and len(text) <= 2:
        return True
    return False


def get_word_wrap_enabled(shape: Dict[str, Any]) -> Optional[bool]:
    text_frame = shape.get("text_frame")
    if not isinstance(text_frame, dict):
        return None
    value = text_frame.get("word_wrap")
    if value is None:
        return None
    return bool(value)


def is_picture_shape(shape: Dict[str, Any]) -> bool:
    shape_type = str(shape.get("shape_type") or "").lower()
    if shape_type == "picture":
        return True
    return str(shape.get("xml_tag") or "").lower() == "pic"


def estimate_line_width_in(text_line: str, font_pt: float) -> float:
    if not text_line:
        return 0.0
    return len(text_line) * font_pt * 0.58 / 72.0


def check_word_wrap_disabled(shape: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not shape_has_text(shape) or is_page_number_shape(shape):
        return None
    if not is_list_style_text(shape):
        return None
    enabled = get_word_wrap_enabled(shape)
    if enabled is True:
        return None
    return {
        "issue_type": "word_wrap_disabled" if enabled is False else "word_wrap_missing",
        "list_style": True,
        "name": shape.get("name"),
        "shape_id": shape.get("shape_id"),
        "text_preview": str(shape.get("text") or "")[:120],
        "geometry": shape_geometry(shape),
    }


def check_unwrapped_horizontal_overflow(
    shape: Dict[str, Any],
    blockers: List[Dict[str, float]],
    margin_in: float = 0.08,
) -> Optional[Dict[str, Any]]:
    if not shape_has_text(shape) or is_page_number_shape(shape):
        return None

    enabled = get_word_wrap_enabled(shape)
    if enabled is True:
        return None
    if not is_list_style_text(shape):
        return None

    geometry = shape_geometry(shape)
    text = str(shape.get("text") or "")
    font_pt = get_font_size_pt(shape)
    max_line_width = 0.0
    widest_line = ""
    for line in normalized_text_lines(text):
        line_width = estimate_line_width_in(line, font_pt)
        if line_width > max_line_width:
            max_line_width = line_width
            widest_line = line

    visual_right = geometry["left_in"] + max(max_line_width, geometry["width_in"])
    for blocker in blockers:
        if visual_right <= blocker["left_in"] - margin_in:
            continue
        if geometry["top_in"] >= blocker["bottom_in"] or (
            geometry["top_in"] + geometry["height_in"]
        ) <= blocker["top_in"]:
            continue
        return {
            "issue_type": "unwrapped_text_picture_overlap",
            "name": shape.get("name"),
            "shape_id": shape.get("shape_id"),
            "text_preview": text[:120],
            "widest_line": widest_line[:80],
            "estimated_visual_right_in": round(visual_right, 3),
            "picture_left_in": round(blocker["left_in"], 3),
            "geometry": geometry,
        }
    return None


def slide_picture_blockers(slide: Dict[str, Any]) -> List[Dict[str, float]]:
    blockers: List[Dict[str, float]] = []
    for shape in slide.get("shapes", []):
        if not is_picture_shape(shape):
            continue
        geometry = shape_geometry(shape)
        if geometry["width_in"] <= 0 or geometry["height_in"] <= 0:
            continue
        blockers.append(
            {
                "left_in": geometry["left_in"],
                "top_in": geometry["top_in"],
                "right_in": geometry["left_in"] + geometry["width_in"],
                "bottom_in": geometry["top_in"] + geometry["height_in"],
            }
        )
    return blockers


def check_shape_wrap_overflow(
    shape: Dict[str, Any],
    tolerance_ratio: float = 1.05,
) -> Optional[Dict[str, Any]]:
    if not shape_has_text(shape) or is_page_number_shape(shape):
        return None
    if not is_list_style_text(shape):
        return None

    geometry = shape_geometry(shape)
    width_in = geometry["width_in"]
    height_in = geometry["height_in"]
    if width_in <= 0 or height_in <= 0:
        return None

    text = str(shape.get("text") or "")
    font_pt = get_font_size_pt(shape)
    line_count = estimate_wrapped_line_count(text, width_in, font_pt)
    estimated_height_in = estimate_text_height_in(line_count, font_pt)
    if estimated_height_in <= height_in * tolerance_ratio:
        return None

    return {
        "issue_type": "wrap_overflow",
        "name": shape.get("name"),
        "shape_id": shape.get("shape_id"),
        "text_preview": text[:120],
        "font_size_pt": font_pt,
        "estimated_lines": line_count,
        "estimated_height_in": round(estimated_height_in, 3),
        "box_height_in": round(height_in, 3),
        "overflow_in": round(estimated_height_in - height_in, 3),
        "geometry": geometry,
    }


def estimated_visual_bottom_in(shape: Dict[str, Any]) -> float:
    geometry = shape_geometry(shape)
    text = str(shape.get("text") or "")
    font_pt = get_font_size_pt(shape)
    line_count = estimate_wrapped_line_count(text, geometry["width_in"], font_pt)
    estimated_height_in = estimate_text_height_in(line_count, font_pt)
    return geometry["top_in"] + max(geometry["height_in"], estimated_height_in)


def check_slide_stack_gaps(
    slide: Dict[str, Any],
    min_gap_in: float = 0.12,
) -> List[Dict[str, Any]]:
    text_shapes = []
    for shape in slide.get("shapes", []):
        if not shape_has_text(shape) or is_page_number_shape(shape):
            continue
        text_shapes.append(shape)

    text_shapes.sort(key=lambda item: shape_geometry(item)["top_in"])
    issues: List[Dict[str, Any]] = []
    for index in range(len(text_shapes) - 1):
        upper = text_shapes[index]
        lower = text_shapes[index + 1]
        upper_bottom = estimated_visual_bottom_in(upper)
        lower_top = shape_geometry(lower)["top_in"]
        gap = lower_top - upper_bottom
        if gap >= min_gap_in - 0.02:
            continue
        issues.append(
            {
                "issue_type": "stack_gap",
                "upper_shape": {
                    "name": upper.get("name"),
                    "shape_id": upper.get("shape_id"),
                    "text_preview": str(upper.get("text") or "")[:80],
                    "estimated_bottom_in": round(upper_bottom, 3),
                },
                "lower_shape": {
                    "name": lower.get("name"),
                    "shape_id": lower.get("shape_id"),
                    "text_preview": str(lower.get("text") or "")[:80],
                    "top_in": round(lower_top, 3),
                },
                "gap_in": round(gap, 3),
                "required_gap_in": min_gap_in,
            }
        )
    return issues


def analyze_layout_wraps(
    layout: Dict[str, Any],
    min_stack_gap_in: float = 0.12,
    wrap_tolerance_ratio: float = 1.05,
) -> Dict[str, Any]:
    wrap_issues: List[Dict[str, Any]] = []
    stack_issues: List[Dict[str, Any]] = []
    word_wrap_disabled_issues: List[Dict[str, Any]] = []
    unwrapped_overlap_issues: List[Dict[str, Any]] = []

    for slide in layout.get("slides", []):
        slide_index = slide.get("slide_index")
        slide_number = slide.get("slide_number", slide_index)
        picture_blockers = slide_picture_blockers(slide)

        for shape in slide.get("shapes", []):
            disabled_issue = check_word_wrap_disabled(shape)
            if disabled_issue is not None:
                disabled_issue["slide_index"] = slide_index
                disabled_issue["slide_number"] = slide_number
                word_wrap_disabled_issues.append(disabled_issue)

            overflow_issue = check_unwrapped_horizontal_overflow(shape, picture_blockers)
            if overflow_issue is not None:
                overflow_issue["slide_index"] = slide_index
                overflow_issue["slide_number"] = slide_number
                unwrapped_overlap_issues.append(overflow_issue)

            issue = check_shape_wrap_overflow(shape, tolerance_ratio=wrap_tolerance_ratio)
            if issue is None:
                continue
            issue["slide_index"] = slide_index
            issue["slide_number"] = slide_number
            wrap_issues.append(issue)

        for issue in check_slide_stack_gaps(slide, min_gap_in=min_stack_gap_in):
            issue["slide_index"] = slide_index
            issue["slide_number"] = slide_number
            stack_issues.append(issue)

    line_wrap_issues = (
        wrap_issues
        + stack_issues
        + word_wrap_disabled_issues
        + unwrapped_overlap_issues
    )
    return {
        "wrap_overflow_issue_count": len(wrap_issues),
        "stack_gap_issue_count": len(stack_issues),
        "word_wrap_disabled_issue_count": len(word_wrap_disabled_issues),
        "unwrapped_text_picture_issue_count": len(unwrapped_overlap_issues),
        "line_wrap_issue_count": len(line_wrap_issues),
        "wrap_overflow_issues": wrap_issues,
        "stack_gap_issues": stack_issues,
        "word_wrap_disabled_issues": word_wrap_disabled_issues,
        "unwrapped_text_picture_issues": unwrapped_overlap_issues,
        "line_wrap_issues": line_wrap_issues,
    }
