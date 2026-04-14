# -*- coding: utf-8 -*-
"""
PPTX 核心操作模組
- 建立新簡報
- 開啟既有簡報
- 儲存簡報
- 新增空白頁
- 新增文字框
- 新增圖片
- 新增表格
- 讀取簡報資訊
- 列出投影片資訊

依賴:
    pip install python-pptx pillow
"""

import os
import json
from typing import Any, Dict, List, Optional, Tuple

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    from copy import deepcopy
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    PYTHON_PPTX_AVAILABLE = True
except ImportError:
    PYTHON_PPTX_AVAILABLE = False

try:
    import math
    import shutil
    import subprocess
    import tempfile
    from pathlib import Path

    import fitz  # PyMuPDF
    import matplotlib.pyplot as plt
    import matplotlib.image as mpimg
    PYTHON_PPTX_EXTRA_AVAILABLE = True
except ImportError:
    PYTHON_PPTX_EXTRA_AVAILABLE = False

try:
    from PIL import Image
    from PIL import ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


EMU_PER_INCH = 914400
DEFAULT_DPI = 96
DEFAULT_FONT_ZH = "微軟正黑體"
DEFAULT_FONT_EN = "Consolas"

import dataProcess as dp
PYTHON_PPTX_AVAILABLE = dp.ptp.PYTHON_PPTX_AVAILABLE
Presentation = dp.ptp.Presentation
m_logger = dp.ptp.m_logger
LOGger = dp.ptp.LOGger

_CONFIG_CACHE: Optional[Dict[str, Any]] = None




def _ensure_pptx_available():
    if not PYTHON_PPTX_AVAILABLE:
        raise ImportError("python-pptx 未安裝，請先執行: pip install python-pptx")


def _px_to_emu(px: int, dpi: int = DEFAULT_DPI) -> int:
    return int((px / float(dpi)) * EMU_PER_INCH)


def _normalize_file_path(file_path: str) -> str:
    if not file_path:
        raise ValueError("file_path 不可為空")
    file_path = os.path.abspath(file_path)
    if not file_path.lower().endswith(".pptx"):
        file_path += ".pptx"
    return file_path


def _ensure_parent_dir(file_path: str) -> None:
    parent_dir = os.path.dirname(file_path)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)


def _get_blank_layout(prs: Presentation):
    # python-pptx 內建常見 blank layout 通常是 index 6
    # 若模板不同，退回最後一個 layout
    try:
        return prs.slide_layouts[6]
    except Exception:
        return prs.slide_layouts[len(prs.slide_layouts) - 1]


def _remove_all_slides(prs: Presentation) -> None:
    """
    清空簡報中所有投影片
    """
    slide_id_list = prs.slides._sldIdLst
    while len(slide_id_list) > 0:
        r_id = slide_id_list[0].rId
        prs.part.drop_rel(r_id)
        del slide_id_list[0]


def _validate_slide_index(prs: Presentation, slide_index: int) -> None:
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(f"slide_index 超出範圍: {slide_index}, 總頁數={len(prs.slides)}")


def _rgb_tuple_to_color(rgb: Optional[Tuple[int, int, int]]) -> Optional[RGBColor]:
    if rgb is None:
        return None
    if len(rgb) != 3:
        raise ValueError("rgb 顏色格式必須是 (R, G, B)")
    r, g, b = rgb
    return RGBColor(int(r), int(g), int(b))


def _load_pptstudio_config() -> Dict[str, Any]:
    global _CONFIG_CACHE
    if _CONFIG_CACHE is not None:
        return _CONFIG_CACHE

    config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")
    if not os.path.exists(config_path):
        _CONFIG_CACHE = {}
        return _CONFIG_CACHE

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            _CONFIG_CACHE = json.load(f)
    except Exception:
        _CONFIG_CACHE = {}
    return _CONFIG_CACHE


def _contains_cjk(text: str) -> bool:
    for char in text:
        code = ord(char)
        if 0x4E00 <= code <= 0x9FFF:
            return True
        if 0x3400 <= code <= 0x4DBF:
            return True
        if 0x3040 <= code <= 0x30FF:
            return True
        if 0xAC00 <= code <= 0xD7AF:
            return True
    return False


def _get_default_fonts() -> Tuple[str, str]:
    config = _load_pptstudio_config()
    fonts_config = config.get("fonts", {}) if isinstance(config, dict) else {}
    zh_font = str(fonts_config.get("default_zh", DEFAULT_FONT_ZH)).strip() or DEFAULT_FONT_ZH
    en_font = str(fonts_config.get("default_en", DEFAULT_FONT_EN)).strip() or DEFAULT_FONT_EN
    return zh_font, en_font


def _looks_like_font_path(value: str) -> bool:
    lowered = value.lower()
    if lowered.endswith((".ttf", ".otf", ".ttc", ".otc")):
        return True
    if os.path.sep in value:
        return True
    if os.path.altsep and os.path.altsep in value:
        return True
    return False


def _resolve_font_name(font_name: Optional[str]) -> Optional[str]:
    if font_name is None:
        return None

    candidate = str(font_name).strip()
    if not candidate:
        return None

    if not _looks_like_font_path(candidate):
        return candidate

    if not os.path.exists(candidate):
        raise FileNotFoundError(f"找不到字型檔: {candidate}")

    font_path = os.path.abspath(candidate)

    if PIL_AVAILABLE:
        try:
            font = ImageFont.truetype(font_path, size=12)
            family_name = font.getname()[0]
            if family_name:
                return family_name
        except Exception:
            pass

    return os.path.splitext(os.path.basename(font_path))[0]


class PPTDocument:
    """
    封裝一個 Presentation 物件，提供可操作方法
    """

    def __init__(self, prs: Presentation, file_path: Optional[str] = None):
        self.prs = prs
        self.file_path = file_path

    @property
    def slide_count(self) -> int:
        return len(self.prs.slides)

    def save(self, file_path: Optional[str] = None) -> str:
        save_path = file_path or self.file_path
        if not save_path:
            raise ValueError("沒有可儲存的 file_path，請明確提供")
        save_path = _normalize_file_path(save_path)
        _ensure_parent_dir(save_path)
        self.prs.save(save_path)
        self.file_path = save_path
        return save_path

    def add_blank_slide(self) -> int:
        blank_layout = _get_blank_layout(self.prs)
        self.prs.slides.add_slide(blank_layout)
        return len(self.prs.slides) - 1

    def add_blank_slides(self, page_num: int = 1) -> List[int]:
        if page_num < 1:
            raise ValueError("page_num 至少要 >= 1")
        indices = []
        for _ in range(page_num):
            indices.append(self.add_blank_slide())
        return indices

    def add_textbox(
            self,
            slide_index: int,
            text: str,
            left: int,
            top: int,
            width: int,
            height: int,
            font_size: int = 20,
            bold: bool = False,
            italic: bool = False,
            font_name: Optional[str] = None,
            font_color: Optional[Tuple[int, int, int]] = None,
            align: str = "left",
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)
        slide = self.prs.slides[slide_index]
        resolved_font_name = _resolve_font_name(font_name)
        default_zh_font, default_en_font = _get_default_fonts()
        resolved_default_zh = _resolve_font_name(default_zh_font)
        resolved_default_en = _resolve_font_name(default_en_font)

        shape = slide.shapes.add_textbox(
            Emu(left),
            Emu(top),
            Emu(width),
            Emu(height),
        )

        tf = shape.text_frame
        tf.clear()

        p = tf.paragraphs[0]
        p.text = text or ""

        align_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY,
        }
        p.alignment = align_map.get((align or "left").lower(), PP_ALIGN.LEFT)

        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            if resolved_font_name:
                run.font.name = resolved_font_name
            else:
                default_font = resolved_default_zh if _contains_cjk(run.text or "") else resolved_default_en
                if default_font:
                    run.font.name = default_font
            color = _rgb_tuple_to_color(font_color)
            if color:
                run.font.color.rgb = color

        return {
            "slide_index": slide_index,
            "shape_id": shape.shape_id,
            "text": text,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }

    def add_image(
            self,
            slide_index: int,
            image_path: str,
            left: int,
            top: int,
            width: Optional[int] = None,
            height: Optional[int] = None,
            keep_aspect_ratio: bool = True,
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)

        if not image_path:
            raise ValueError("image_path 不可為空")
        image_path = os.path.abspath(image_path)
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"找不到圖片: {image_path}")

        slide = self.prs.slides[slide_index]

        img_width = width
        img_height = height

        if keep_aspect_ratio and PIL_AVAILABLE and (width is None or height is None):
            with Image.open(image_path) as img:
                px_w, px_h = img.size
            ratio = px_w / px_h if px_h else 1.0

            if width is not None and height is None:
                img_width = width
                img_height = int(width / ratio)
            elif height is not None and width is None:
                img_height = height
                img_width = int(height * ratio)

        kwargs = {
            "image_file": image_path,
            "left": Emu(left),
            "top": Emu(top),
        }
        if img_width is not None:
            kwargs["width"] = Emu(img_width)
        if img_height is not None:
            kwargs["height"] = Emu(img_height)

        picture = slide.shapes.add_picture(**kwargs)

        return {
            "slide_index": slide_index,
            "shape_id": picture.shape_id,
            "image_path": image_path,
            "left": left,
            "top": top,
            "width": picture.width,
            "height": picture.height,
        }

    def add_table(
            self,
            slide_index: int,
            rows: int,
            cols: int,
            left: int,
            top: int,
            width: int,
            height: int,
            data: Optional[List[List[Any]]] = None,
            first_row_as_header: bool = False,
            font_size: int = 14,
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)

        if rows < 1 or cols < 1:
            raise ValueError("rows 與 cols 都必須 >= 1")

        slide = self.prs.slides[slide_index]
        graphic_frame = slide.shapes.add_table(
            rows=rows,
            cols=cols,
            left=Emu(left),
            top=Emu(top),
            width=Emu(width),
            height=Emu(height),
        )
        table = graphic_frame.table

        if data:
            for r in range(min(rows, len(data))):
                row_data = data[r]
                for c in range(min(cols, len(row_data))):
                    cell = table.cell(r, c)
                    cell.text = "" if row_data[c] is None else str(row_data[c])

                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size)
                            if first_row_as_header and r == 0:
                                run.font.bold = True

        return {
            "slide_index": slide_index,
            "shape_id": graphic_frame.shape_id,
            "rows": rows,
            "cols": cols,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }

    def add_shape(
            self,
            slide_index: int,
            shape_type: str,
            left: int,
            top: int,
            width: int,
            height: int,
            text: str = "",
            fill_color: Optional[Tuple[int, int, int]] = None,
            line_color: Optional[Tuple[int, int, int]] = None,
            line_width: Optional[int] = None,
            font_size: int = 18,
            bold: bool = False,
            font_name: Optional[str] = None,
            font_color: Optional[Tuple[int, int, int]] = None,
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)
        slide = self.prs.slides[slide_index]

        shape_map = {
            "rectangle": MSO_SHAPE.RECTANGLE,
            "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
            "oval": MSO_SHAPE.OVAL,
            "diamond": MSO_SHAPE.DIAMOND,
            "hexagon": MSO_SHAPE.HEXAGON,
            "parallelogram": MSO_SHAPE.PARALLELOGRAM,
            "trapezoid": MSO_SHAPE.TRAPEZOID,
            "chevron": MSO_SHAPE.CHEVRON,
            "right_arrow": MSO_SHAPE.RIGHT_ARROW,
            "left_arrow": MSO_SHAPE.LEFT_ARROW,
            "up_arrow": MSO_SHAPE.UP_ARROW,
            "down_arrow": MSO_SHAPE.DOWN_ARROW,
            "cloud": MSO_SHAPE.CLOUD,
        }

        if shape_type not in shape_map:
            raise ValueError(f"不支援的 shape_type: {shape_type}")

        shape = slide.shapes.add_shape(
            shape_map[shape_type],
            Emu(left),
            Emu(top),
            Emu(width),
            Emu(height),
        )

        # fill
        if fill_color is not None:
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb_tuple_to_color(fill_color)

        # line
        if line_color is not None:
            shape.line.color.rgb = _rgb_tuple_to_color(line_color)

        if line_width is not None:
            shape.line.width = Emu(line_width)

        # text
        if text:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = text

            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.bold = bold
                if font_name:
                    run.font.name = font_name
                if font_color is not None:
                    run.font.color.rgb = _rgb_tuple_to_color(font_color)

        return {
            "slide_index": slide_index,
            "shape_id": shape.shape_id,
            "shape_type": shape_type,
            "text": text,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }
    
    def add_line(
            self,
            slide_index: int,
            x1: int,
            y1: int,
            x2: int,
            y2: int,
            line_color: Optional[Tuple[int, int, int]] = None,
            line_width: Optional[int] = None,
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)
        slide = self.prs.slides[slide_index]

        shape = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Emu(x1),
            Emu(y1),
            Emu(x2),
            Emu(y2),
        )

        if line_color is not None:
            shape.line.color.rgb = _rgb_tuple_to_color(line_color)

        if line_width is not None:
            shape.line.width = Emu(line_width)

        return {
            "slide_index": slide_index,
            "shape_id": shape.shape_id,
            "type": "line",
            "x1": x1,
            "y1": y1,
            "x2": x2,
            "y2": y2,
        }
    
    def add_arrow(
            self,
            slide_index: int,
            left: int,
            top: int,
            width: int,
            height: int,
            direction: str = "right",
            text: str = "",
            fill_color: Optional[Tuple[int, int, int]] = None,
            line_color: Optional[Tuple[int, int, int]] = None,
            line_width: Optional[int] = None,
            font_size: int = 18,
            bold: bool = False,
            font_name: Optional[str] = None,
            font_color: Optional[Tuple[int, int, int]] = None,
        ) -> Dict[str, Any]:
        direction_map = {
            "right": "right_arrow",
            "left": "left_arrow",
            "up": "up_arrow",
            "down": "down_arrow",
        }

        if direction not in direction_map:
            raise ValueError("direction 必須是 right / left / up / down")

        return self.add_shape(
            slide_index=slide_index,
            shape_type=direction_map[direction],
            left=left,
            top=top,
            width=width,
            height=height,
            text=text,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width,
            font_size=font_size,
            bold=bold,
            font_name=font_name,
            font_color=font_color,
        )
    
    def set_slide_background_color(
            self,
            slide_index: int,
            rgb: Tuple[int, int, int],
        ) -> Dict[str, Any]:
        """
        設定指定頁的純色背景
        """
        _validate_slide_index(self.prs, slide_index)

        if len(rgb) != 3:
            raise ValueError("rgb 必須是 (R, G, B)")

        slide = self.prs.slides[slide_index]
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(int(rgb[0]), int(rgb[1]), int(rgb[2]))

        return {
            "slide_index": slide_index,
            "background_type": "color",
            "rgb": [int(rgb[0]), int(rgb[1]), int(rgb[2])],
        }

    def set_slide_background_image(
            self,
            slide_index: int,
            image_path: str,
        ) -> Dict[str, Any]:
        """
        用滿版圖片模擬投影片背景
        """
        _validate_slide_index(self.prs, slide_index)

        if not image_path:
            raise ValueError("image_path 不可為空")

        image_path = os.path.abspath(image_path)
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"找不到圖片: {image_path}")

        slide = self.prs.slides[slide_index]

        picture = slide.shapes.add_picture(
            image_path,
            0,
            0,
            width=self.prs.slide_width,
            height=self.prs.slide_height,
        )

        return {
            "slide_index": slide_index,
            "background_type": "image",
            "image_path": image_path,
            "shape_id": picture.shape_id,
        }

    def set_slides_background_color(
            self,
            slide_indices: List[int],
            rgb: Tuple[int, int, int],
        ) -> List[Dict[str, Any]]:
        results = []
        for slide_index in slide_indices:
            results.append(self.set_slide_background_color(slide_index, rgb))
        return results

    def set_slides_background_image(
            self,
            slide_indices: List[int],
            image_path: str,
        ) -> List[Dict[str, Any]]:
        results = []
        for slide_index in slide_indices:
            results.append(self.set_slide_background_image(slide_index, image_path))
        return results

    def get_info(self) -> Dict[str, Any]:
        return {
            "file_path": self.file_path,
            "slide_count": len(self.prs.slides),
            "slide_width": int(self.prs.slide_width),
            "slide_height": int(self.prs.slide_height),
            "layout_count": len(self.prs.slide_layouts),
        }

    def list_slides(self) -> List[Dict[str, Any]]:
        results = []
        for idx, slide in enumerate(self.prs.slides):
            shape_infos = []
            title_text = None

            for shape in slide.shapes:
                shape_info = {
                    "shape_id": getattr(shape, "shape_id", None),
                    "name": getattr(shape, "name", None),
                    "shape_type": str(getattr(shape, "shape_type", "")),
                    "has_text_frame": bool(getattr(shape, "has_text_frame", False)),
                }
                if getattr(shape, "has_text_frame", False):
                    try:
                        text = shape.text_frame.text.strip()
                    except Exception:
                        text = ""
                    shape_info["text_preview"] = text[:120]
                    if title_text is None and text:
                        title_text = text[:120]
                shape_infos.append(shape_info)

            results.append({
                "slide_index": idx,
                "shape_count": len(slide.shapes),
                "title_preview": title_text,
                "shapes": shape_infos,
            })
        return results

    def delete_slide(self, slide_index: int) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)

        slide_id_list = self.prs.slides._sldIdLst
        slide_id = slide_id_list[slide_index]
        r_id = slide_id.rId
        self.prs.part.drop_rel(r_id)
        del slide_id_list[slide_index]

        return {
            "deleted_slide_index": slide_index,
            "remaining_slide_count": len(self.prs.slides),
        }

    def duplicate_slide(self, slide_index: int) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)

        source_slide = self.prs.slides[slide_index]
        blank_layout = _get_blank_layout(self.prs)
        new_slide = self.prs.slides.add_slide(blank_layout)

        # 複製 shape XML
        for shape in source_slide.shapes:
            new_el = deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

        # 複製背景（盡量）
        try:
            if source_slide.background.fill.type is not None:
                src_fill = source_slide.background.fill
                dst_fill = new_slide.background.fill
                if hasattr(src_fill, "fore_color") and src_fill.fore_color.type is not None:
                    dst_fill.solid()
                    if hasattr(src_fill.fore_color, "rgb") and src_fill.fore_color.rgb:
                        dst_fill.fore_color.rgb = src_fill.fore_color.rgb
        except Exception:
            pass

        return {
            "source_slide_index": slide_index,
            "new_slide_index": len(self.prs.slides) - 1,
            "slide_count": len(self.prs.slides),
        }

    def replace_text(
            self,
            old_text: str,
            new_text: str,
            slide_indices: Optional[List[int]] = None,
            exact_match: bool = False,
            case_sensitive: bool = True,
        ) -> Dict[str, Any]:
        if not old_text:
            raise ValueError("old_text 不可為空")

        if slide_indices is None:
            target_indices = list(range(len(self.prs.slides)))
        else:
            target_indices = slide_indices
            for idx in target_indices:
                _validate_slide_index(self.prs, idx)

        total_replacements = 0
        matched_shapes = []

        def _replace(src: str) -> Tuple[str, int]:
            if exact_match:
                if case_sensitive:
                    if src == old_text:
                        return new_text, 1
                    return src, 0
                else:
                    if src.lower() == old_text.lower():
                        return new_text, 1
                    return src, 0
            else:
                if case_sensitive:
                    count = src.count(old_text)
                    return src.replace(old_text, new_text), count
                else:
                    import re
                    pattern = re.compile(re.escape(old_text), re.IGNORECASE)
                    replaced, count = pattern.subn(new_text, src)
                    return replaced, count

        for slide_index in target_indices:
            slide = self.prs.slides[slide_index]
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue

                changed = False
                shape_replace_count = 0

                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        original = run.text or ""
                        replaced, cnt = _replace(original)
                        if cnt > 0:
                            run.text = replaced
                            shape_replace_count += cnt
                            changed = True

                    # 如果 paragraph 沒有 runs，直接處理 paragraph.text
                    if len(paragraph.runs) == 0:
                        original = paragraph.text or ""
                        replaced, cnt = _replace(original)
                        if cnt > 0:
                            paragraph.text = replaced
                            shape_replace_count += cnt
                            changed = True

                if changed:
                    total_replacements += shape_replace_count
                    matched_shapes.append({
                        "slide_index": slide_index,
                        "shape_id": getattr(shape, "shape_id", None),
                        "replace_count": shape_replace_count,
                    })

        return {
            "old_text": old_text,
            "new_text": new_text,
            "total_replacements": total_replacements,
            "matched_shapes": matched_shapes,
        }

    def add_bullets(
            self,
            slide_index: int,
            items: List[str],
            left: int,
            top: int,
            width: int,
            height: int,
            font_size: int = 20,
            level: int = 0,
            bold: bool = False,
            font_name: Optional[str] = None,
            font_color: Optional[Tuple[int, int, int]] = None,
        ) -> Dict[str, Any]:
        _validate_slide_index(self.prs, slide_index)

        if not items:
            raise ValueError("items 不可為空")

        slide = self.prs.slides[slide_index]
        shape = slide.shapes.add_textbox(
            Emu(left),
            Emu(top),
            Emu(width),
            Emu(height),
        )

        tf = shape.text_frame
        tf.clear()

        for idx, item in enumerate(items):
            if idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = "" if item is None else str(item)
            p.level = max(0, int(level))
            p.bullet = True

            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.bold = bold
                if font_name:
                    run.font.name = font_name
                color = _rgb_tuple_to_color(font_color)
                if color:
                    run.font.color.rgb = color

        return {
            "slide_index": slide_index,
            "shape_id": shape.shape_id,
            "item_count": len(items),
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }

    def add_title_slide(
            self,
            title: str,
            subtitle: str = "",
        ) -> Dict[str, Any]:
        layout = None

        # 盡量找 Title Slide layout
        for candidate in self.prs.slide_layouts:
            try:
                placeholder_types = []
                for ph in candidate.placeholders:
                    try:
                        placeholder_types.append(ph.placeholder_format.type)
                    except Exception:
                        pass
                if PP_PLACEHOLDER.TITLE in placeholder_types:
                    layout = candidate
                    break
            except Exception:
                continue

        if layout is None:
            layout = self.prs.slide_layouts[0]

        slide = self.prs.slides.add_slide(layout)

        if slide.shapes.title is not None:
            slide.shapes.title.text = title or ""

        # 找副標題 placeholder
        subtitle_set = False
        if subtitle:
            for shape in slide.placeholders:
                try:
                    ph_type = shape.placeholder_format.type
                    if ph_type == PP_PLACEHOLDER.SUBTITLE:
                        shape.text = subtitle
                        subtitle_set = True
                        break
                except Exception:
                    pass

            if not subtitle_set:
                if not subtitle_set and subtitle:
                    tb = slide.shapes.add_textbox(
                        left=Emu(500000),
                        top=Emu(1500000),
                        width=Emu(6000000),
                        height=Emu(800000),
                    )
                    tb.text_frame.text = subtitle

        return {
            "slide_index": len(self.prs.slides) - 1,
            "title": title,
            "subtitle": subtitle,
        }

    def reorder_slides(self, new_order: List[int]) -> Dict[str, Any]:
        slide_count = len(self.prs.slides)

        if len(new_order) != slide_count:
            raise ValueError(f"new_order 長度必須等於目前頁數 {slide_count}")

        if sorted(new_order) != list(range(slide_count)):
            raise ValueError("new_order 必須是 0 到 slide_count-1 的完整排列")

        sld_id_lst = self.prs.slides._sldIdLst
        current_nodes = list(sld_id_lst)

        # 先清空再依新順序放回
        for _ in range(len(sld_id_lst)):
            del sld_id_lst[0]

        for old_index in new_order:
            sld_id_lst.append(current_nodes[old_index])

        return {
            "slide_count": slide_count,
            "new_order": new_order,
        }

def new(
        file_path: str,
        plank_page_num: int = 1,
        plank_page_width: int = 1080,
        plank_page_height: int = 1920,
        dpi: int = DEFAULT_DPI,
    ) -> str:
    """
    建立新的 pptx 檔案

    Args:
        file_path: 輸出檔案路徑
        plank_page_num: 頁面數量
        plank_page_width: 頁面寬度（px）
        plank_page_height: 頁面高度（px）
        dpi: 像素轉尺寸時使用的 DPI

    Returns:
        str: 建立後的檔案路徑
    """
    doc = new_document(
        file_path=file_path,
        plank_page_num=plank_page_num,
        plank_page_width=plank_page_width,
        plank_page_height=plank_page_height,
        dpi=dpi,
    )
    return doc.save(file_path)

def new_document(
        file_path: Optional[str] = None,
        plank_page_num: int = 1,
        plank_page_width: int = 1080,
        plank_page_height: int = 1920,
        dpi: int = DEFAULT_DPI,
    ) -> PPTDocument:
    _ensure_pptx_available()

    if plank_page_num < 1:
        raise ValueError("plank_page_num 至少要 >= 1")
    if plank_page_width <= 0 or plank_page_height <= 0:
        raise ValueError("頁面寬高必須 > 0")

    prs = Presentation()
    prs.slide_width = _px_to_emu(plank_page_width, dpi=dpi)
    prs.slide_height = _px_to_emu(plank_page_height, dpi=dpi)

    _remove_all_slides(prs)

    doc = PPTDocument(prs=prs, file_path=file_path)
    doc.add_blank_slides(plank_page_num)
    return doc

def open_presentation(file_path: str) -> PPTDocument:
    _ensure_pptx_available()

    file_path = _normalize_file_path(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"檔案不存在: {file_path}")

    prs = Presentation(file_path)
    return PPTDocument(prs=prs, file_path=file_path)


def save(document: PPTDocument, file_path: Optional[str] = None) -> str:
    return document.save(file_path=file_path)


def add_blank_slide(document: PPTDocument) -> int:
    return document.add_blank_slide()


def add_blank_slides(document: PPTDocument, page_num: int = 1) -> List[int]:
    return document.add_blank_slides(page_num=page_num)


def add_text(
        document: PPTDocument,
        slide_index: int,
        text: str,
        left: int,
        top: int,
        width: int,
        height: int,
        font_size: int = 20,
        bold: bool = False,
        italic: bool = False,
        font_name: Optional[str] = None,
        font_color: Optional[Tuple[int, int, int]] = None,
        align: str = "left",
    ) -> Dict[str, Any]:
    return document.add_textbox(
        slide_index=slide_index,
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        bold=bold,
        italic=italic,
        font_name=font_name,
        font_color=font_color,
        align=align,
    )


def add_image(
        document: PPTDocument,
        slide_index: int,
        image_path: str,
        left: int,
        top: int,
        width: Optional[int] = None,
        height: Optional[int] = None,
        keep_aspect_ratio: bool = True,
    ) -> Dict[str, Any]:
    return document.add_image(
        slide_index=slide_index,
        image_path=image_path,
        left=left,
        top=top,
        width=width,
        height=height,
        keep_aspect_ratio=keep_aspect_ratio,
    )


def add_table(
        document: PPTDocument,
        slide_index: int,
        rows: int,
        cols: int,
        left: int,
        top: int,
        width: int,
        height: int,
        data: Optional[List[List[Any]]] = None,
        first_row_as_header: bool = False,
        font_size: int = 14,
    ) -> Dict[str, Any]:
    return document.add_table(
        slide_index=slide_index,
        rows=rows,
        cols=cols,
        left=left,
        top=top,
        width=width,
        height=height,
        data=data,
        first_row_as_header=first_row_as_header,
        font_size=font_size,
    )


def add_shape(
        document: PPTDocument,
        slide_index: int,
        shape_type: str,
        left: int,
        top: int,
        width: int,
        height: int,
        text: str = "",
        fill_color: Optional[Tuple[int, int, int]] = None,
        line_color: Optional[Tuple[int, int, int]] = None,
        line_width: Optional[int] = None,
        font_size: int = 18,
        bold: bool = False,
        font_name: Optional[str] = None,
        font_color: Optional[Tuple[int, int, int]] = None,
    ) -> Dict[str, Any]:
    return document.add_shape(
        slide_index=slide_index,
        shape_type=shape_type,
        left=left,
        top=top,
        width=width,
        height=height,
        text=text,
        fill_color=fill_color,
        line_color=line_color,
        line_width=line_width,
        font_size=font_size,
        bold=bold,
        font_name=font_name,
        font_color=font_color,
    )


def add_line(
        document: PPTDocument,
        slide_index: int,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        line_color: Optional[Tuple[int, int, int]] = None,
        line_width: Optional[int] = None,
    ) -> Dict[str, Any]:
    return document.add_line(
        slide_index=slide_index,
        x1=x1,
        y1=y1,
        x2=x2,
        y2=y2,
        line_color=line_color,
        line_width=line_width,
    )


def add_arrow(
        document: PPTDocument,
        slide_index: int,
        left: int,
        top: int,
        width: int,
        height: int,
        direction: str = "right",
        text: str = "",
        fill_color: Optional[Tuple[int, int, int]] = None,
        line_color: Optional[Tuple[int, int, int]] = None,
        line_width: Optional[int] = None,
        font_size: int = 18,
        bold: bool = False,
        font_name: Optional[str] = None,
        font_color: Optional[Tuple[int, int, int]] = None,
    ) -> Dict[str, Any]:
    return document.add_arrow(
        slide_index=slide_index,
        left=left,
        top=top,
        width=width,
        height=height,
        direction=direction,
        text=text,
        fill_color=fill_color,
        line_color=line_color,
        line_width=line_width,
        font_size=font_size,
        bold=bold,
        font_name=font_name,
        font_color=font_color,
    )


def get_info(document: PPTDocument) -> Dict[str, Any]:
    return document.get_info()


def list_slides(document: PPTDocument) -> List[Dict[str, Any]]:
    return document.list_slides()


def set_slide_background_color(
        document: PPTDocument,
        slide_index: int,
        rgb: Tuple[int, int, int],
    ) -> dict:
    return document.set_slide_background_color(slide_index=slide_index, rgb=rgb)


def set_slide_background_image(
        document: PPTDocument,
        slide_index: int,
        image_path: str,
    ) -> dict:
    return document.set_slide_background_image(slide_index=slide_index, image_path=image_path)


def delete_slide(document: PPTDocument, slide_index: int) -> Dict[str, Any]:
    return document.delete_slide(slide_index)


def duplicate_slide(document: PPTDocument, slide_index: int) -> Dict[str, Any]:
    return document.duplicate_slide(slide_index)


def replace_text(
        document: PPTDocument,
        old_text: str,
        new_text: str,
        slide_indices: Optional[List[int]] = None,
        exact_match: bool = False,
        case_sensitive: bool = True,
    ) -> Dict[str, Any]:
    return document.replace_text(
        old_text=old_text,
        new_text=new_text,
        slide_indices=slide_indices,
        exact_match=exact_match,
        case_sensitive=case_sensitive,
    )


def add_bullets(
        document: PPTDocument,
        slide_index: int,
        items: List[str],
        left: int,
        top: int,
        width: int,
        height: int,
        font_size: int = 20,
        level: int = 0,
        bold: bool = False,
        font_name: Optional[str] = None,
        font_color: Optional[Tuple[int, int, int]] = None,
    ) -> Dict[str, Any]:
    return document.add_bullets(
        slide_index=slide_index,
        items=items,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        level=level,
        bold=bold,
        font_name=font_name,
        font_color=font_color,
    )


def add_title_slide(
        document: PPTDocument,
        title: str,
        subtitle: str = "",
    ) -> Dict[str, Any]:
    return document.add_title_slide(title=title, subtitle=subtitle)


def reorder_slides(document: PPTDocument, new_order: List[int]) -> Dict[str, Any]:
    return document.reorder_slides(new_order)


def set_slides_background_color(
        document: PPTDocument,
        slide_indices: List[int],
        rgb: Tuple[int, int, int],
    ) -> List[Dict[str, Any]]:
    return document.set_slides_background_color(slide_indices=slide_indices, rgb=rgb)


def set_slides_background_image(
        document: PPTDocument,
        slide_indices: List[int],
        image_path: str,
    ) -> List[Dict[str, Any]]:
    return document.set_slides_background_image(slide_indices=slide_indices, image_path=image_path)


def _find_libreoffice_executable(explicit_path: Optional[str] = None) -> str:
    """
    尋找 LibreOffice / soffice 可執行檔
    """
    candidates = []

    if explicit_path:
        candidates.append(explicit_path)

    # PATH 中常見名稱
    for name in ["soffice", "libreoffice", "soffice.exe"]:
        found = shutil.which(name)
        if found:
            candidates.append(found)

    # Windows 常見安裝路徑
    candidates.extend([
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ])

    for path in candidates:
        if path and os.path.exists(path):
            return os.path.abspath(path)

    raise FileNotFoundError("找不到 LibreOffice/soffice，可明確傳入 libreoffice_path")


def convert_pptx_to_pdf(
        pptx_path: str,
        output_dir: Optional[str] = None,
        libreoffice_path: Optional[str] = None,
        overwrite: bool = True,
    ) -> str:
    """
    用 LibreOffice headless 將 PPTX 轉成 PDF
    """
    pptx_path = os.path.abspath(pptx_path)
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"找不到 PPTX: {pptx_path}")

    soffice = _find_libreoffice_executable(libreoffice_path)

    if output_dir is None:
        output_dir = os.path.dirname(pptx_path) or os.getcwd()
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    expected_pdf = os.path.join(
        output_dir,
        f"{Path(pptx_path).stem}.pdf"
    )

    if os.path.exists(expected_pdf) and overwrite:
        os.remove(expected_pdf)

    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        pptx_path,
    ]

    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )

    if proc.returncode != 0:
        raise RuntimeError(
            f"LibreOffice 轉 PDF 失敗\n"
            f"cmd={' '.join(cmd)}\n"
            f"stdout={proc.stdout}\n"
            f"stderr={proc.stderr}"
        )

    if not os.path.exists(expected_pdf):
        raise FileNotFoundError(
            f"LibreOffice 指令執行完畢，但找不到輸出的 PDF: {expected_pdf}\n"
            f"stdout={proc.stdout}\n"
            f"stderr={proc.stderr}"
        )

    return expected_pdf


def render_slide_to_image(
        pptx_path: str,
        slide_index: int,
        output_path: str,
        dpi: int = 150,
        libreoffice_path: Optional[str] = None,
        temp_dir: Optional[str] = None,
    ) -> Dict[str, Any]:
    """
    將指定頁投影片輸出成圖片
    - 先轉 PDF
    - 再用 PyMuPDF 將指定頁 rasterize 成 png
    """
    if slide_index < 0:
        raise ValueError("slide_index 必須 >= 0")

    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with tempfile.TemporaryDirectory(dir=temp_dir) as tmpdir:
        pdf_path = convert_pptx_to_pdf(
            pptx_path=pptx_path,
            output_dir=tmpdir,
            libreoffice_path=libreoffice_path,
            overwrite=True,
        )

        doc = fitz.open(pdf_path)
        try:
            if slide_index >= len(doc):
                raise IndexError(f"slide_index 超出範圍: {slide_index}, 總頁數={len(doc)}")

            page = doc[slide_index]
            scale = dpi / 72.0
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            pix.save(output_path)

            return {
                "pptx_path": os.path.abspath(pptx_path),
                "pdf_path": pdf_path,
                "slide_index": slide_index,
                "page_count": len(doc),
                "dpi": dpi,
                "output_path": output_path,
                "width": pix.width,
                "height": pix.height,
            }
        finally:
            doc.close()


def render_slides_to_grid_image(
        pptx_path: str,
        slide_indices: List[int],
        output_path: str,
        cols: int = 2,
        dpi: int = 150,
        libreoffice_path: Optional[str] = None,
        temp_dir: Optional[str] = None,
        add_page_title: bool = True,
        figure_title: Optional[str] = None,
    ) -> Dict[str, Any]:
    """
    將多頁投影片輸出並拼成一張 grid 圖
    """
    if not slide_indices:
        raise ValueError("slide_indices 不可為空")
    if cols < 1:
        raise ValueError("cols 必須 >= 1")

    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with tempfile.TemporaryDirectory(dir=temp_dir) as tmpdir:
        pdf_path = convert_pptx_to_pdf(
            pptx_path=pptx_path,
            output_dir=tmpdir,
            libreoffice_path=libreoffice_path,
            overwrite=True,
        )

        doc = fitz.open(pdf_path)
        image_paths = []
        try:
            page_count = len(doc)

            for slide_index in slide_indices:
                if slide_index < 0 or slide_index >= page_count:
                    raise IndexError(f"slide_index 超出範圍: {slide_index}, 總頁數={page_count}")

            # 先逐頁輸出暫存 png
            for slide_index in slide_indices:
                page = doc[slide_index]
                scale = dpi / 72.0
                matrix = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=matrix, alpha=False)

                img_path = os.path.join(tmpdir, f"slide_{slide_index}.png")
                pix.save(img_path)
                image_paths.append(img_path)
        finally:
            doc.close()

        rows = math.ceil(len(image_paths) / cols)
        fig, axes = plt.subplots(rows, cols, figsize=(cols * 4.5, rows * 3.2), squeeze=False)

        for ax in axes.flat:
            ax.axis("off")

        for i, img_path in enumerate(image_paths):
            r = i // cols
            c = i % cols
            ax = axes[r][c]
            img = mpimg.imread(img_path)
            ax.imshow(img)
            ax.axis("off")
            if add_page_title:
                ax.set_title(f"Slide {slide_indices[i]}", fontsize=10)

        if figure_title:
            fig.suptitle(figure_title)

        plt.tight_layout()
        fig.savefig(output_path, dpi=dpi, bbox_inches="tight")
        plt.close(fig)

        return {
            "pptx_path": os.path.abspath(pptx_path),
            "slide_indices": slide_indices,
            "count": len(slide_indices),
            "cols": cols,
            "rows": rows,
            "dpi": dpi,
            "output_path": output_path,
        }


if __name__ == "__main__":
    # 簡單測試
    out = new("test_output/demo_new_ppt", plank_page_num=2, plank_page_width=1080, plank_page_height=1920)
    doc = open_presentation(out)

    add_text(
        document=doc,
        slide_index=0,
        text="Hello PPT Core",
        left=1000000,
        top=500000,
        width=5000000,
        height=800000,
        font_size=24,
        bold=True,
    )

    add_table(
        document=doc,
        slide_index=1,
        rows=3,
        cols=3,
        left=800000,
        top=800000,
        width=6000000,
        height=2000000,
        data=[
            ["欄位A", "欄位B", "欄位C"],
            [1, 2, 3],
            [4, 5, 6],
        ],
        first_row_as_header=True,
    )

    add_title_slide(doc, "主標題", "副標題")
    add_bullets(
        document=doc,
        slide_index=0,
        items=["第一點", "第二點", "第三點"],
        left=800000,
        top=1200000,
        width=5000000,
        height=2000000,
        font_size=24,
    )
    replace_text(doc, "第二點", "第二點（已更新）")
    duplicate_slide(doc, 0)
    reorder_slides(doc, [1, 0, 2, 3])

    save(doc)
    print(get_info(doc))
