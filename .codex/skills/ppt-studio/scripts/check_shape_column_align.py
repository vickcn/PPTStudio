#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""檢查單頁 JSON 中，同名 prefix 的 shape 是否 left_in/width_in 一致（欄位對齊驗收）。"""
from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check column alignment for shapes on one slide.")
    parser.add_argument("json_path", help="Layout JSON path")
    parser.add_argument("--slide-index", type=int, default=2, help="JSON slide_index (1-based), default 2")
    parser.add_argument(
        "--prefixes",
        nargs="*",
        default=["AgendaCard", "PageTag", "AccentBar", "Oval"],
        help="Shape name prefixes to compare",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    data = json.loads(Path(args.json_path).read_text(encoding="utf-8"))
    slide = None
    for item in data.get("slides", []):
        if int(item.get("slide_index") or 0) == args.slide_index:
            slide = item
            break
    if slide is None:
        print(f"slide_index={args.slide_index} not found", file=sys.stderr)
        return 1

    groups: dict[str, list[dict]] = defaultdict(list)
    for shape in slide.get("shapes") or []:
        name = str(shape.get("name") or "")
        for prefix in args.prefixes:
            if name.startswith(prefix):
                geom = shape.get("geometry") or {}
                groups[prefix].append(
                    {
                        "name": name,
                        "left_in": round(float(geom.get("left_in") or 0), 3),
                        "width_in": round(float(geom.get("width_in") or 0), 3),
                        "height_in": round(float(geom.get("height_in") or 0), 3),
                    }
                )
                break

    ok = True
    for prefix, items in sorted(groups.items()):
        if not items:
            continue
        lefts = {x["left_in"] for x in items}
        widths = {x["width_in"] for x in items}
        print(f"[{prefix}] count={len(items)} lefts={sorted(lefts)} widths={sorted(widths)}")
        if len(lefts) > 1 or len(widths) > 1:
            ok = False
            for row in items:
                print(f"  mismatch: {row}")

    if ok:
        print("column_align: OK")
        return 0
    print("column_align: FAIL", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
