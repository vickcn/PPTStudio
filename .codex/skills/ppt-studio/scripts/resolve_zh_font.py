#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""List Traditional Chinese font files and family names for chart/image generation."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List

# family_hint is for humans; matplotlib/PIL may report a different internal name.
WINDOWS_ZH_CANDIDATES = [
    ("Microsoft JhengHei", "msjh.ttc"),
    ("Microsoft JhengHei Bold", "msjhbd.ttc"),
    ("Microsoft JhengHei Light", "msjhl.ttc"),
    ("Microsoft YaHei", "msyh.ttc"),
    ("Microsoft YaHei Bold", "msyhbd.ttc"),
    ("MingLiU", "mingliu.ttc"),
    ("MingLiU-ExtB", "mingliub.ttc"),
    ("DFKai-SB", "kaiu.ttf"),
]

PPTSTUDIO_CONFIG = Path(r"C:\ML_HOME\PPTStudio\config.json")


def _windows_fonts_dir() -> Path:
    windir = os.environ.get("WINDIR", r"C:\Windows")
    return Path(windir) / "Fonts"


def _read_pptstudio_default_zh() -> str | None:
    if not PPTSTUDIO_CONFIG.exists():
        return None
    try:
        data = json.loads(PPTSTUDIO_CONFIG.read_text(encoding="utf-8"))
        value = data.get("fonts", {}).get("default_zh")
        return str(value).strip() if value else None
    except Exception:
        return None


def _scan_windows_font_files() -> List[Dict[str, str]]:
    fonts_dir = _windows_fonts_dir()
    found: List[Dict[str, str]] = []
    for family_hint, filename in WINDOWS_ZH_CANDIDATES:
        path = fonts_dir / filename
        if path.exists():
            found.append(
                {
                    "family_hint": family_hint,
                    "filename": filename,
                    "path": str(path),
                }
            )
    return found


def _scan_matplotlib_fonts() -> List[Dict[str, str]]:
    try:
        from matplotlib import font_manager as fm
    except ImportError:
        return []

    keywords = (
        "Jheng",
        "YaHei",
        "Ming",
        "Kai",
        "DFKai",
        "Noto Sans CJK",
        "PingFang",
        "Microsoft",
        "正黑",
        "雅黑",
        "細明",
        "標楷",
    )
    seen: set[tuple[str, str]] = set()
    result: List[Dict[str, str]] = []
    for entry in fm.fontManager.ttflist:
        name = entry.name or ""
        if not any(k in name for k in keywords):
            continue
        key = (name, entry.fname)
        if key in seen:
            continue
        seen.add(key)
        result.append({"name": name, "path": entry.fname})
    return sorted(result, key=lambda item: item["name"])


def resolve_zh_fonts() -> Dict[str, Any]:
    windows_files = _scan_windows_font_files()
    recommended_path = windows_files[0]["path"] if windows_files else None
    return {
        "config_default_zh": _read_pptstudio_default_zh(),
        "windows_fonts_dir": str(_windows_fonts_dir()),
        "windows_font_files": windows_files,
        "recommended_path": recommended_path,
        "matplotlib_fonts": _scan_matplotlib_fonts(),
        "pptstudio_config": str(PPTSTUDIO_CONFIG) if PPTSTUDIO_CONFIG.exists() else None,
    }


def main() -> int:
    print(json.dumps(resolve_zh_fonts(), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
