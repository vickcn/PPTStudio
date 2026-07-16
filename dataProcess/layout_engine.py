# -*- coding: utf-8 -*-
"""PPT 版面演算：投影片座標、文字量測、字型擬合、區塊分割、圖片 px→EMU、z-order、manifest、預檢。"""
from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

EMU_PER_INCH = 914400.0
PT_PER_INCH = 72.0
CJK_EM_WIDTH = 1.02
LATIN_EM_WIDTH = 0.52
DIGIT_EM_WIDTH = 0.58
SPACE_EM_WIDTH = 0.35
BOLD_WIDTH_FACTOR = 1.06
DEFAULT_LEADING = 1.28
DEFAULT_INNER_PAD_IN = 0.08


@dataclass(frozen=True)
class RectPx:
    x: float
    y: float
    w: float
    h: float

    def padded(self, pad_px: float) -> "RectPx":
        if pad_px <= 0:
            return self
        return RectPx(
            x=self.x - pad_px,
            y=self.y - pad_px,
            w=self.w + pad_px * 2,
            h=self.h + pad_px * 2,
        )


@dataclass(frozen=True)
class RectIn:
    left: float
    top: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height


@dataclass(frozen=True)
class RectEmu:
    left: int
    top: int
    width: int
    height: int


@dataclass(frozen=True)
class ImagePlacement:
    """圖片在 picture shape 內的實際顯示區（含 letterbox）。"""
    content_left_in: float
    content_top_in: float
    content_width_in: float
    content_height_in: float
    image_width_px: int
    image_height_px: int


@dataclass(frozen=True)
class LayoutIssue:
    code: str
    message: str
    shape_name: Optional[str] = None
    shape_id: Optional[int] = None


def geometry_to_rect_in(geometry: Dict[str, Any]) -> RectIn:
    return RectIn(
        left=float(geometry.get("left_in") or 0.0),
        top=float(geometry.get("top_in") or 0.0),
        width=float(geometry.get("width_in") or 0.0),
        height=float(geometry.get("height_in") or 0.0),
    )


def rect_in_to_emu(rect: RectIn) -> RectEmu:
    return RectEmu(
        left=int(round(rect.left * EMU_PER_INCH)),
        top=int(round(rect.top * EMU_PER_INCH)),
        width=int(round(rect.width * EMU_PER_INCH)),
        height=int(round(rect.height * EMU_PER_INCH)),
    )


def geometry_dict_from_rect_in(rect: RectIn) -> Dict[str, Union[int, float]]:
    emu = rect_in_to_emu(rect)
    return {
        "left_in": round(rect.left, 3),
        "top_in": round(rect.top, 3),
        "width_in": round(rect.width, 3),
        "height_in": round(rect.height, 3),
        "left_emu": emu.left,
        "top_emu": emu.top,
        "width_emu": emu.width,
        "height_emu": emu.height,
    }


def compute_image_placement_in(
    picture_geometry: Dict[str, Any],
    image_width_px: int,
    image_height_px: int,
) -> ImagePlacement:
    box = geometry_to_rect_in(picture_geometry)
    if box.width <= 0 or box.height <= 0:
        raise ValueError("picture geometry 寬高必須 > 0")
    if image_width_px <= 0 or image_height_px <= 0:
        raise ValueError("image_size_px 必須 > 0")

    box_aspect = box.width / box.height
    img_aspect = image_width_px / image_height_px

    if img_aspect >= box_aspect:
        content_width_in = box.width
        content_height_in = box.width / img_aspect
        content_left_in = box.left
        content_top_in = box.top + (box.height - content_height_in) / 2.0
    else:
        content_height_in = box.height
        content_width_in = box.height * img_aspect
        content_left_in = box.left + (box.width - content_width_in) / 2.0
        content_top_in = box.top

    return ImagePlacement(
        content_left_in=content_left_in,
        content_top_in=content_top_in,
        content_width_in=content_width_in,
        content_height_in=content_height_in,
        image_width_px=int(image_width_px),
        image_height_px=int(image_height_px),
    )


def map_rect_px_to_slide_in(
    picture_geometry: Dict[str, Any],
    image_size_px: Tuple[int, int],
    rect_px: RectPx,
    pad_px: float = 0,
) -> RectIn:
    img_w, img_h = int(image_size_px[0]), int(image_size_px[1])
    placement = compute_image_placement_in(picture_geometry, img_w, img_h)
    region = rect_px.padded(pad_px)

    rel_x = region.x / img_w
    rel_y = region.y / img_h
    rel_w = region.w / img_w
    rel_h = region.h / img_h

    return RectIn(
        left=placement.content_left_in + rel_x * placement.content_width_in,
        top=placement.content_top_in + rel_y * placement.content_height_in,
        width=rel_w * placement.content_width_in,
        height=rel_h * placement.content_height_in,
    )


def map_rect_px_to_slide_geometry(
    picture_geometry: Dict[str, Any],
    image_size_px: Tuple[int, int],
    rect_px: RectPx,
    pad_px: float = 0,
) -> Dict[str, Union[int, float]]:
    rect_in = map_rect_px_to_slide_in(picture_geometry, image_size_px, rect_px, pad_px=pad_px)
    return geometry_dict_from_rect_in(rect_in)


def manifest_path_for_image(image_path: Union[str, Path]) -> Path:
    path = Path(image_path)
    return path.with_suffix(path.suffix + ".layout.json")


def write_layout_manifest(
    image_path: Union[str, Path],
    image_width: int,
    image_height: int,
    regions: Dict[str, Dict[str, float]],
    meta: Optional[Dict[str, Any]] = None,
) -> Path:
    out = manifest_path_for_image(image_path)
    payload = {
        "version": 1,
        "image_path": str(Path(image_path).resolve()),
        "image_width": int(image_width),
        "image_height": int(image_height),
        "regions": regions,
        "meta": meta or {},
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def load_layout_manifest(path: Union[str, Path]) -> Dict[str, Any]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if "regions" not in data:
        raise ValueError(f"manifest 缺少 regions: {path}")
    return data


def resolve_anchor_rect_px(
    manifest: Dict[str, Any],
    anchor_id: str,
    pad_px: float = 0,
) -> RectPx:
    regions = manifest.get("regions") or {}
    key = str(anchor_id)
    if key not in regions:
        raise KeyError(f"manifest 找不到 anchor: {anchor_id}")
    item = regions[key]
    rect = RectPx(
        x=float(item["x"]),
        y=float(item["y"]),
        w=float(item["w"]),
        h=float(item["h"]),
    )
    return rect.padded(pad_px)


def plan_overlay_geometry(
    picture_geometry: Dict[str, Any],
    manifest: Dict[str, Any],
    anchor_id: str,
    pad_px: float = 0,
) -> Dict[str, Union[int, float]]:
    img_w = int(manifest.get("image_width") or 0)
    img_h = int(manifest.get("image_height") or 0)
    if img_w <= 0 or img_h <= 0:
        raise ValueError("manifest 缺少 image_width / image_height")
    rect_px = resolve_anchor_rect_px(manifest, anchor_id, pad_px=pad_px)
    return map_rect_px_to_slide_geometry(
        picture_geometry,
        (img_w, img_h),
        rect_px,
        pad_px=0,
    )


def plan_overlay_from_paths(
    picture_geometry: Dict[str, Any],
    image_path: Union[str, Path],
    anchor_id: str,
    pad_px: float = 0,
) -> Dict[str, Union[int, float]]:
    manifest_path = manifest_path_for_image(image_path)
    if not manifest_path.exists():
        raise FileNotFoundError(f"找不到 layout manifest: {manifest_path}")
    manifest = load_layout_manifest(manifest_path)
    return plan_overlay_geometry(picture_geometry, manifest, anchor_id, pad_px=pad_px)


def is_picture_shape(shape: Dict[str, Any]) -> bool:
    shape_type = str(shape.get("shape_type") or "").lower()
    if shape_type == "picture":
        return True
    return str(shape.get("xml_tag") or "").lower() == "pic"


def is_overlay_shape(shape: Dict[str, Any]) -> bool:
    if shape.get("fill_none"):
        return True
    name = str(shape.get("name") or "")
    if name in {"ExtractHighlight"}:
        return True
    text = shape.get("text")
    if text == "" and shape.get("line_color") and not shape.get("fill_color"):
        return True
    return False


def shape_layer_rank(shape: Dict[str, Any]) -> int:
    if is_picture_shape(shape):
        return 10
    if is_overlay_shape(shape):
        return 20
    return 30


def order_shapes_for_build(shapes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    indexed = list(enumerate(shapes or []))
    indexed.sort(key=lambda item: (shape_layer_rank(item[1]), item[0]))
    return [shape for _, shape in indexed]


def rects_overlap(a: RectIn, b: RectIn, margin_in: float = 0) -> bool:
    return not (
        a.right <= b.left + margin_in
        or b.right <= a.left + margin_in
        or a.bottom <= b.top + margin_in
        or b.bottom <= a.top + margin_in
    )


def validate_slide_layout(
    slide: Dict[str, Any],
    slide_size: Optional[Dict[str, Any]] = None,
) -> List[LayoutIssue]:
    issues: List[LayoutIssue] = []
    slide_w = float((slide_size or {}).get("width_in") or 10.0)
    slide_h = float((slide_size or {}).get("height_in") or 7.5)
    shapes = slide.get("shapes") or []

    rects: List[Tuple[Dict[str, Any], RectIn]] = []
    for shape in shapes:
        geom = shape.get("geometry") or {}
        if not geom.get("width_in"):
            continue
        rect = geometry_to_rect_in(geom)
        if rect.width <= 0 or rect.height <= 0:
            issues.append(
                LayoutIssue(
                    code="zero_geometry",
                    message="shape 寬或高 <= 0",
                    shape_name=str(shape.get("name") or ""),
                    shape_id=shape.get("shape_id"),
                )
            )
            continue
        if rect.left < -0.01 or rect.top < -0.01:
            issues.append(
                LayoutIssue(
                    code="out_of_slide",
                    message="shape 左上角超出投影片",
                    shape_name=str(shape.get("name") or ""),
                    shape_id=shape.get("shape_id"),
                )
            )
        if rect.right > slide_w + 0.05 or rect.bottom > slide_h + 0.05:
            issues.append(
                LayoutIssue(
                    code="out_of_slide",
                    message="shape 超出投影片範圍",
                    shape_name=str(shape.get("name") or ""),
                    shape_id=shape.get("shape_id"),
                )
            )
        rects.append((shape, rect))

    pictures = [item for item in rects if is_picture_shape(item[0])]
    overlays = [item for item in rects if is_overlay_shape(item[0])]

    for overlay_shape, overlay_rect in overlays:
        if not pictures:
            issues.append(
                LayoutIssue(
                    code="overlay_without_picture",
                    message="overlay 形狀但本頁沒有 picture",
                    shape_name=str(overlay_shape.get("name") or ""),
                    shape_id=overlay_shape.get("shape_id"),
                )
            )
            continue
        picture_rect = pictures[0][1]
        if not rects_overlap(overlay_rect, picture_rect, margin_in=-0.02):
            issues.append(
                LayoutIssue(
                    code="overlay_off_picture",
                    message="overlay 未落在 picture 顯示區內",
                    shape_name=str(overlay_shape.get("name") or ""),
                    shape_id=overlay_shape.get("shape_id"),
                )
            )

    return issues


def issues_to_dicts(issues: List[LayoutIssue]) -> List[Dict[str, Any]]:
    return [asdict(item) for item in issues]


# ---------------------------------------------------------------------------
# 通用投影片 / 文字版面演算（圖案、文字框、文字藝術師皆適用）
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SlideCanvas:
    """投影片可排版區域（含邊界）。"""
    width_in: float = 10.0
    height_in: float = 7.5
    margin_left_in: float = 0.5
    margin_top_in: float = 0.35
    margin_right_in: float = 0.5
    margin_bottom_in: float = 0.3

    @property
    def content(self) -> RectIn:
        return RectIn(
            left=self.margin_left_in,
            top=self.margin_top_in,
            width=max(self.width_in - self.margin_left_in - self.margin_right_in, 0.1),
            height=max(self.height_in - self.margin_top_in - self.margin_bottom_in, 0.1),
        )


@dataclass(frozen=True)
class TextBlockSpec:
    """單一文字區塊需求（標題、正文、標籤、文字藝術師文案）。"""
    text: str
    role: str = "body"
    single_line: bool = False
    bold: bool = False
    max_lines: Optional[int] = None
    min_font_pt: float = 9.0
    max_font_pt: float = 44.0
    font_size_pt: Optional[float] = None
    leading: float = DEFAULT_LEADING
    inner_pad_in: float = DEFAULT_INNER_PAD_IN


@dataclass(frozen=True)
class PlannedTypography:
    font_size_pt: float
    line_count: int
    line_height_in: float
    block_height_in: float
    text_width_in: float
    fits: bool


@dataclass(frozen=True)
class PlannedPlacement:
    """演算後的方塊位置 + 建議字型。"""
    rect: RectIn
    typography: PlannedTypography
    role: str = "body"
    single_line: bool = False
    bold: bool = False
    text: str = ""

    def geometry_dict(self) -> Dict[str, Union[int, float]]:
        return geometry_dict_from_rect_in(self.rect)


def pt_to_in(font_pt: float) -> float:
    return float(font_pt) / PT_PER_INCH


def in_to_pt(value_in: float) -> float:
    return float(value_in) * PT_PER_INCH


def slide_canvas_from_size(
    slide_size: Optional[Dict[str, Any]] = None,
    margins: Optional[Tuple[float, float, float, float]] = None,
) -> SlideCanvas:
    size = slide_size or {}
    width_in = float(size.get("width_in") or 10.0)
    height_in = float(size.get("height_in") or 7.5)
    if margins is None:
        return SlideCanvas(width_in=width_in, height_in=height_in)
    left, top, right, bottom = margins
    return SlideCanvas(
        width_in=width_in,
        height_in=height_in,
        margin_left_in=left,
        margin_top_in=top,
        margin_right_in=right,
        margin_bottom_in=bottom,
    )


def inset_rect(rect: RectIn, pad_in: float) -> RectIn:
    pad = max(float(pad_in), 0.0)
    return RectIn(
        left=rect.left + pad,
        top=rect.top + pad,
        width=max(rect.width - pad * 2.0, 0.05),
        height=max(rect.height - pad * 2.0, 0.05),
    )


def char_width_units(ch: str) -> float:
    if ch == " ":
        return SPACE_EM_WIDTH
    if ch == "\n" or ch == "\r" or ch == "\t":
        return 0.0
    code = ord(ch)
    if code < 128:
        if ch.isdigit():
            return DIGIT_EM_WIDTH
        return LATIN_EM_WIDTH
    return 1.0


def text_width_units(text: str) -> float:
    return sum(char_width_units(ch) for ch in text)


def em_width_in(font_pt: float, bold: bool = False) -> float:
    factor = BOLD_WIDTH_FACTOR if bold else 1.0
    return pt_to_in(font_pt) * CJK_EM_WIDTH * factor


def estimate_text_width_in(text: str, font_pt: float, bold: bool = False) -> float:
    return text_width_units(text) * em_width_in(font_pt, bold=bold)


def line_height_in(font_pt: float, leading: float = DEFAULT_LEADING) -> float:
    return pt_to_in(font_pt) * leading


def chars_capacity_per_line(width_in: float, font_pt: float, bold: bool = False) -> float:
    unit_w = em_width_in(font_pt, bold=bold)
    if unit_w <= 0:
        return 1.0
    return max(width_in / unit_w, 1.0)


def estimate_wrapped_line_count(
    text: str,
    width_in: float,
    font_pt: float,
    bold: bool = False,
) -> int:
    if width_in <= 0 or font_pt <= 0:
        return 1
    lines = 0
    paragraphs = str(text or "").split("\n")
    if not paragraphs:
        return 1
    cap = chars_capacity_per_line(width_in, font_pt, bold=bold)
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
            continue
        units = text_width_units(paragraph)
        lines += max(1, int(math.ceil(units / cap)))
    return max(lines, 1)


def estimate_block_height_in(
    text: str,
    width_in: float,
    font_pt: float,
    bold: bool = False,
    leading: float = DEFAULT_LEADING,
    single_line: bool = False,
) -> float:
    lines = 1 if single_line else estimate_wrapped_line_count(text, width_in, font_pt, bold=bold)
    return lines * line_height_in(font_pt, leading=leading)


def text_fits_box(
    text: str,
    width_in: float,
    height_in: float,
    font_pt: float,
    bold: bool = False,
    leading: float = DEFAULT_LEADING,
    single_line: bool = False,
    inner_pad_in: float = DEFAULT_INNER_PAD_IN,
) -> bool:
    inner = inset_rect(RectIn(0, 0, width_in, height_in), inner_pad_in)
    if inner.width <= 0 or inner.height <= 0:
        return False
    if single_line:
        if estimate_text_width_in(text.replace("\n", " "), font_pt, bold=bold) > inner.width:
            return False
        return line_height_in(font_pt, leading=leading) <= inner.height
    block_h = estimate_block_height_in(
        text,
        inner.width,
        font_pt,
        bold=bold,
        leading=leading,
        single_line=False,
    )
    return block_h <= inner.height


def fit_font_size(
    text: str,
    width_in: float,
    height_in: float,
    min_font_pt: float = 9.0,
    max_font_pt: float = 44.0,
    bold: bool = False,
    leading: float = DEFAULT_LEADING,
    single_line: bool = False,
    inner_pad_in: float = DEFAULT_INNER_PAD_IN,
) -> float:
    lo = float(min_font_pt)
    hi = float(max_font_pt)
    if lo > hi:
        lo, hi = hi, lo
    best = lo
    for _ in range(28):
        mid = (lo + hi) / 2.0
        if text_fits_box(
            text,
            width_in,
            height_in,
            mid,
            bold=bold,
            leading=leading,
            single_line=single_line,
            inner_pad_in=inner_pad_in,
        ):
            best = mid
            lo = mid
        else:
            hi = mid
    return round(best, 1)


def split_vertical_equal(region: RectIn, count: int, gap_in: float = 0.0) -> List[RectIn]:
    if count <= 0:
        return []
    gap = max(float(gap_in), 0.0)
    total_gap = gap * max(count - 1, 0)
    row_h = (region.height - total_gap) / float(count)
    rects: List[RectIn] = []
    top = region.top
    for _ in range(count):
        rects.append(RectIn(region.left, top, region.width, row_h))
        top += row_h + gap
    return rects


def split_horizontal_weights(
    region: RectIn,
    weights: List[float],
    gap_in: float = 0.0,
) -> List[RectIn]:
    if not weights:
        return []
    gap = max(float(gap_in), 0.0)
    total_weight = sum(max(float(w), 0.01) for w in weights)
    total_gap = gap * max(len(weights) - 1, 0)
    usable = region.width - total_gap
    rects: List[RectIn] = []
    left = region.left
    for weight in weights:
        w = usable * (max(float(weight), 0.01) / total_weight)
        rects.append(RectIn(left, region.top, w, region.height))
        left += w + gap
    return rects


def center_rect_in(parent: RectIn, child_width: float, child_height: float) -> RectIn:
    return RectIn(
        left=parent.left + (parent.width - child_width) / 2.0,
        top=parent.top + (parent.height - child_height) / 2.0,
        width=child_width,
        height=child_height,
    )


def reserve_text_column(
    slide_width_in: float,
    picture_left_in: float,
    text_left_in: float,
    margin_in: float = 0.15,
) -> float:
    """側圖投影片：依 picture 左緣反推文字欄寬（吋）。"""
    return max(picture_left_in - text_left_in - margin_in, 0.5)


def plan_typography(
    spec: TextBlockSpec,
    box: RectIn,
) -> PlannedTypography:
    font_pt = spec.font_size_pt
    if font_pt is None:
        font_pt = fit_font_size(
            spec.text,
            box.width,
            box.height,
            min_font_pt=spec.min_font_pt,
            max_font_pt=spec.max_font_pt,
            bold=spec.bold,
            leading=spec.leading,
            single_line=spec.single_line,
            inner_pad_in=spec.inner_pad_in,
        )
    inner = inset_rect(box, spec.inner_pad_in)
    line_count = 1 if spec.single_line else estimate_wrapped_line_count(
        spec.text,
        inner.width,
        font_pt,
        bold=spec.bold,
    )
    if spec.max_lines is not None:
        line_count = min(line_count, int(spec.max_lines))
    lh = line_height_in(font_pt, leading=spec.leading)
    block_h = line_count * lh
    text_w = estimate_text_width_in(spec.text.replace("\n", " "), font_pt, bold=spec.bold)
    fits = text_fits_box(
        spec.text,
        box.width,
        box.height,
        font_pt,
        bold=spec.bold,
        leading=spec.leading,
        single_line=spec.single_line,
        inner_pad_in=spec.inner_pad_in,
    )
    return PlannedTypography(
        font_size_pt=float(font_pt),
        line_count=int(line_count),
        line_height_in=lh,
        block_height_in=block_h,
        text_width_in=text_w,
        fits=fits,
    )


def plan_text_block(spec: TextBlockSpec, box: RectIn) -> PlannedPlacement:
    typo = plan_typography(spec, box)
    inner = inset_rect(box, spec.inner_pad_in)
    if spec.single_line:
        top = inner.top + max((inner.height - typo.line_height_in) / 2.0, 0.0)
        rect = RectIn(inner.left, top, inner.width, max(typo.line_height_in, inner.height))
    else:
        top = inner.top + max((inner.height - typo.block_height_in) / 2.0, 0.0)
        rect = RectIn(inner.left, top, inner.width, max(typo.block_height_in, typo.line_height_in))
    return PlannedPlacement(
        rect=rect,
        typography=typo,
        role=spec.role,
        single_line=spec.single_line,
        bold=spec.bold,
        text=spec.text,
    )


def plan_stacked_text_blocks(
    specs: List[TextBlockSpec],
    region: RectIn,
    gap_in: float = 0.08,
    weights: Optional[List[float]] = None,
) -> List[PlannedPlacement]:
    if not specs:
        return []
    gap = max(float(gap_in), 0.0)
    total_gap = gap * max(len(specs) - 1, 0)
    usable_h = region.height - total_gap
    if weights is None:
        weights = [1.0 for _ in specs]
    weight_sum = sum(max(float(w), 0.01) for w in weights)
    placements: List[PlannedPlacement] = []
    top = region.top
    for spec, weight in zip(specs, weights):
        h = usable_h * (max(float(weight), 0.01) / weight_sum)
        box = RectIn(region.left, top, region.width, h)
        placements.append(plan_text_block(spec, box))
        top += h + gap
    return placements


def plan_header_band(
    canvas: SlideCanvas,
    title: str,
    subtitle: str = "",
    band_height_in: Optional[float] = None,
    title_weight: float = 0.62,
) -> Tuple[PlannedPlacement, Optional[PlannedPlacement], float]:
    """頁首標題帶：回傳 (title, subtitle|None, 下一區塊 top_in)。"""
    content = canvas.content
    band_h = band_height_in if band_height_in is not None else min(content.height * 0.18, 1.35)
    band = RectIn(content.left, content.top, content.width, band_h)
    if subtitle:
        blocks = plan_stacked_text_blocks(
            [
                TextBlockSpec(
                    text=title,
                    role="title",
                    single_line=True,
                    bold=True,
                    max_font_pt=34.0,
                    min_font_pt=18.0,
                ),
                TextBlockSpec(
                    text=subtitle,
                    role="subtitle",
                    single_line=True,
                    bold=False,
                    max_font_pt=14.0,
                    min_font_pt=9.0,
                ),
            ],
            band,
            gap_in=0.06,
            weights=[title_weight, 1.0 - title_weight],
        )
        return blocks[0], blocks[1], band.bottom + 0.08
    title_only = plan_text_block(
        TextBlockSpec(
            text=title,
            role="title",
            single_line=True,
            bold=True,
            max_font_pt=34.0,
            min_font_pt=18.0,
        ),
        band,
    )
    return title_only, None, band.bottom + 0.08


def plan_square_badge(
    label: str,
    host: RectIn,
    fill_ratio: float = 0.72,
    min_font_pt: float = 12.0,
    max_font_pt: float = 22.0,
) -> PlannedPlacement:
    size = min(host.width, host.height) * max(min(fill_ratio, 1.0), 0.2)
    box = center_rect_in(host, size, size)
    return plan_text_block(
        TextBlockSpec(
            text=label,
            role="badge",
            single_line=True,
            bold=True,
            min_font_pt=min_font_pt,
            max_font_pt=max_font_pt,
        ),
        box,
    )


def plan_row_with_columns(
    region: RectIn,
    column_weights: List[float],
    gap_in: float = 0.1,
) -> List[RectIn]:
    return split_horizontal_weights(region, column_weights, gap_in=gap_in)


def placement_to_dict(placement: PlannedPlacement) -> Dict[str, Any]:
    return {
        "role": placement.role,
        "text": placement.text,
        "single_line": placement.single_line,
        "bold": placement.bold,
        "geometry": placement.geometry_dict(),
        "typography": asdict(placement.typography),
    }


def placements_to_dicts(placements: List[PlannedPlacement]) -> List[Dict[str, Any]]:
    return [placement_to_dict(item) for item in placements]


def canvas_to_dict(canvas: SlideCanvas) -> Dict[str, Any]:
    payload = asdict(canvas)
    payload["content"] = geometry_dict_from_rect_in(canvas.content)
    return payload

