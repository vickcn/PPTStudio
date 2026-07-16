# -*- coding: utf-8 -*-
"""列出 JSON 指定頁各 shape 幾何，並報告非預期的跨 shape 矩形重疊。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def rect(sh: dict) -> tuple[float, float, float, float]:
    g = sh.get("geometry") or {}
    l = float(g.get("left_in") or 0)
    t = float(g.get("top_in") or 0)
    w = float(g.get("width_in") or 0)
    h = float(g.get("height_in") or 0)
    return l, t, l + w, t + h


def overlap_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    ol = max(a[0], b[0])
    ot = max(a[1], b[1])
    or_ = min(a[2], b[2])
    ob = min(a[3], b[3])
    if or_ <= ol or ob <= ot:
        return 0.0
    return (or_ - ol) * (ob - ot)


def is_expected_pair(name_a: str, name_b: str) -> bool:
    """文字在卡片/表格內、色條在卡邊等預期重疊。"""
    pairs = (
        ("Bg", "Bar"),
        ("Bg", "L0"),
        ("Bg", "L1"),
        ("Bg", "L2"),
        ("TableBg", "Th_"),
        ("TableBg", "ObjAnchor"),
        ("MainBg", "MainTitle"),
        ("MainBg", "Flow"),
    )
    for left, right in pairs:
        if left in name_a and right in name_b:
            return True
        if left in name_b and right in name_a:
            return True
    if name_a.rsplit("Bg", 1)[0] == name_b.rsplit("Bar", 1)[0] and "Bar" in name_b:
        return True
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect shape overlaps on one slide")
    parser.add_argument("json_path", type=Path)
    parser.add_argument("--slide-index", type=int, required=True, help="JSON slide_index (1-based)")
    parser.add_argument("--min-area", type=float, default=0.05, help="Min overlap area in^2 to report")
    args = parser.parse_args()

    data = json.loads(args.json_path.read_text(encoding="utf-8"))
    slide = next(s for s in data["slides"] if int(s.get("slide_index") or 0) == args.slide_index)
    shapes = slide.get("shapes") or []

    print(f"slide_index={args.slide_index} shape_count={len(shapes)}")
    for sh in shapes:
        l, t, r, b = rect(sh)
        text = (sh.get("text") or "").replace("\n", " ")[:40]
        print(f"  {sh.get('name','?'):18} [{l:.2f},{t:.2f}]-[{r:.2f},{b:.2f}]  {text!r}")

    issues = 0
    print("\n--- unexpected overlaps ---")
    for i, a in enumerate(shapes):
        ra = rect(a)
        na = str(a.get("name") or "")
        for j, b in enumerate(shapes):
            if j <= i:
                continue
            rb = rect(b)
            nb = str(b.get("name") or "")
            area = overlap_area(ra, rb)
            if area < args.min_area:
                continue
            if is_expected_pair(na, nb):
                continue
            issues += 1
            print(f"  {na} x {nb}: area={area:.3f} in^2")

    if issues:
        print(f"\nunexpected_overlap_count={issues}")
        return 2
    print("\nunexpected_overlap_count=0")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
