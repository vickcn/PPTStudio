# -*- coding: utf-8 -*-
"""
api_server.py

FastAPI 包裝 PPT 核心功能
啟動:
    uvicorn api_server:app --host 0.0.0.0 --port 8010

依賴:
    pip install fastapi uvicorn python-pptx pillow
"""

import os
import json
import traceback
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from dataProcess.ppt_stdio import (
    new,
    open_presentation,
    save,
    add_blank_slide,
    add_blank_slides,
    add_text,
    add_wordart_like_textbox,
    update_wordart_text,
    add_image,
    add_table,
    add_line,
    add_arrow,
    add_shape,
    get_info,
    list_slides,
    get_textbox_style,
    get_slide_textbox_styles,
    set_textbox_style,
    delete_textbox,
    delete_shape,
    clone_named_shape_from_template,
    get_slide_text_fonts,
    scan_presentation_text_fonts,
    get_presentation_theme_info,
    get_slide_background_info,
    scan_presentation_backgrounds,
    set_slide_background_color,
    set_slide_background_image,
    set_slides_background_color,
    set_slides_background_image,
    delete_slide,
    duplicate_slide,
    replace_text,
    add_bullets,
    add_title_slide,
    reorder_slides,
    render_slide_to_image,
    render_slides_to_grid_image,
)


app = FastAPI(
    title="PPT API Server",
    version="0.1.0",
    description="提供 PPTX 建立與編輯的 API，之後可再包成 MCP tools",
)


def _load_server_config() -> Dict[str, Any]:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if not os.path.exists(config_path):
        return {}
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


# -------------------------
# Request / Response Models
# -------------------------

class CreatePPTRequest(BaseModel):
    file_path: str = Field(..., description="輸出 pptx 路徑")
    plank_page_num: int = Field(1, ge=1)
    plank_page_width: int = Field(1080, gt=0)
    plank_page_height: int = Field(1920, gt=0)
    dpi: int = Field(96, gt=0)


class AddBlankSlidesRequest(BaseModel):
    file_path: str
    page_num: int = Field(1, ge=1)
    save_as: Optional[str] = None


class AddTextRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    text: str = ""
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    font_size: int = Field(20, gt=0)
    bold: bool = False
    italic: bool = False
    font_name: Optional[str] = None
    font_color: Optional[List[int]] = None
    align: str = "left"
    save_as: Optional[str] = None


class AddWordartLikeTextboxRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    text: str = ""
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    font_size: int = Field(28, gt=0)
    bold: bool = True
    font_name: Optional[str] = None
    font_color: Optional[List[int]] = None
    align: str = "center"
    save_as: Optional[str] = None


class UpdateWordartTextRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    new_text: str = ""
    shape_name: Optional[str] = None
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class AddImageRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    image_path: str
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: Optional[int] = Field(None, gt=0)
    height: Optional[int] = Field(None, gt=0)
    keep_aspect_ratio: bool = True
    save_as: Optional[str] = None


class AddTableRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    rows: int = Field(..., ge=1)
    cols: int = Field(..., ge=1)
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    data: Optional[List[List[Any]]] = None
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)
    save_as: Optional[str] = None


class DeleteSlideRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    save_as: Optional[str] = None


class DeleteTextboxRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class DeleteShapeRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class CloneNamedShapeFromTemplateRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_name: str
    new_text: str = ""
    left: Optional[int] = Field(None, ge=0)
    top: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class DuplicateSlideRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    save_as: Optional[str] = None


class ReplaceTextRequest(BaseModel):
    file_path: str
    old_text: str
    new_text: str
    slide_indices: Optional[List[int]] = None
    exact_match: bool = False
    case_sensitive: bool = True
    save_as: Optional[str] = None


class AddBulletsRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    items: List[str]
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    font_size: int = Field(20, gt=0)
    level: int = Field(0, ge=0)
    bold: bool = False
    font_name: Optional[str] = None
    font_color: Optional[List[int]] = None
    save_as: Optional[str] = None


class AddTitleSlideRequest(BaseModel):
    file_path: str
    title: str
    subtitle: str = ""
    save_as: Optional[str] = None


class AddShapeRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_type: str
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    text: str = ""
    fill_color: Optional[List[int]] = None
    line_color: Optional[List[int]] = None
    line_width: Optional[int] = Field(None, gt=0)
    font_size: int = Field(18, gt=0)
    bold: bool = False
    font_name: Optional[str] = None
    font_color: Optional[List[int]] = None
    save_as: Optional[str] = None


class AddLineRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    x1: int = Field(..., ge=0)
    y1: int = Field(..., ge=0)
    x2: int = Field(..., ge=0)
    y2: int = Field(..., ge=0)
    line_color: Optional[List[int]] = None
    line_width: Optional[int] = Field(None, gt=0)
    save_as: Optional[str] = None


class AddArrowRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: int = Field(..., gt=0)
    height: int = Field(..., gt=0)
    direction: str = "right"
    text: str = ""
    fill_color: Optional[List[int]] = None
    line_color: Optional[List[int]] = None
    line_width: Optional[int] = Field(None, gt=0)
    font_size: int = Field(18, gt=0)
    bold: bool = False
    font_name: Optional[str] = None
    font_color: Optional[List[int]] = None
    save_as: Optional[str] = None


class ReorderSlidesRequest(BaseModel):
    file_path: str
    new_order: List[int]
    save_as: Optional[str] = None


class SetSlidesBackgroundColorRequest(BaseModel):
    file_path: str
    slide_indices: List[int]
    rgb: List[int] = Field(..., min_length=3, max_length=3)
    save_as: Optional[str] = None


class SetSlidesBackgroundImageRequest(BaseModel):
    file_path: str
    slide_indices: List[int]
    image_path: str
    save_as: Optional[str] = None


class SaveAsRequest(BaseModel):
    file_path: str
    save_as: str


class SetSlideBackgroundColorRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    rgb: List[int] = Field(..., min_length=3, max_length=3, description="背景色 [R, G, B]")
    save_as: Optional[str] = None


class SetSlideBackgroundImageRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    image_path: str
    save_as: Optional[str] = None


class SetTextboxStyleRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    fill_color: Optional[List[int]] = Field(None, min_length=3, max_length=3)
    fill_transparency: Optional[float] = Field(None, ge=0.0, le=1.0)
    line_style: Optional[str] = None
    line_color: Optional[List[int]] = Field(None, min_length=3, max_length=3)
    line_width: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class RenderSlideToImageRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    output_path: str
    dpi: int = Field(150, ge=72)
    libreoffice_path: Optional[str] = None


class RenderSlidesToGridImageRequest(BaseModel):
    file_path: str
    slide_indices: List[int]
    output_path: str
    cols: int = Field(2, ge=1)
    dpi: int = Field(150, ge=72)
    libreoffice_path: Optional[str] = None
    add_page_title: bool = True
    figure_title: Optional[str] = None


def _ok(data: Any = None, message: str = "success") -> Dict[str, Any]:
    return {
        "ok": True,
        "message": message,
        "data": data,
    }


def _err_to_http(e: Exception):
    raise HTTPException(
        status_code=500,
        detail={
            "error": str(e),
            "traceback": traceback.format_exc(),
        }
    )


# -------------------------
# Health
# -------------------------

@app.get("/")
def root():
    return _ok({
        "service": "PPT API Server",
        "version": "0.1.0",
    })


@app.get("/health")
def health():
    return _ok({"status": "healthy"})


# -------------------------
# PPT Create / Load Info
# -------------------------

@app.post("/ppt/new")
def create_new_ppt(req: CreatePPTRequest):
    try:
        out_path = new(
            file_path=req.file_path,
            plank_page_num=req.plank_page_num,
            plank_page_width=req.plank_page_width,
            plank_page_height=req.plank_page_height,
            dpi=req.dpi,
        )
        doc = open_presentation(out_path)
        return _ok({
            "file_path": out_path,
            "info": get_info(doc),
            "slides": list_slides(doc),
        }, message="ppt 建立成功")
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/info")
def ppt_info(file_path: str):
    try:
        doc = open_presentation(file_path)
        return _ok(get_info(doc))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slides")
def ppt_slides(file_path: str):
    try:
        doc = open_presentation(file_path)
        return _ok(list_slides(doc))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/textbox_style")
def ppt_textbox_style(file_path: str, slide_index: int = 0, shape_id: Optional[int] = None, shape_index: Optional[int] = None):
    try:
        doc = open_presentation(file_path)
        return _ok(get_textbox_style(doc, slide_index=slide_index, shape_id=shape_id, shape_index=shape_index))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slide_textbox_styles")
def ppt_slide_textbox_styles(file_path: str, slide_index: int = 0):
    try:
        doc = open_presentation(file_path)
        return _ok(get_slide_textbox_styles(doc, slide_index=slide_index))
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_textbox_style")
def ppt_set_textbox_style(req: SetTextboxStyleRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_textbox_style(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            fill_color=tuple(req.fill_color) if req.fill_color is not None else None,
            fill_transparency=req.fill_transparency,
            line_style=req.line_style,
            line_color=tuple(req.line_color) if req.line_color is not None else None,
            line_width=req.line_width,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="更新文字框樣式成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_textbox")
def ppt_delete_textbox(req: DeleteTextboxRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_textbox(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="刪除文字框成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_shape")
def ppt_delete_shape(req: DeleteShapeRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_shape(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="刪除 shape 成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/clone_named_shape_from_template")
def ppt_clone_named_shape_from_template(req: CloneNamedShapeFromTemplateRequest):
    try:
        doc = open_presentation(req.file_path)
        result = clone_named_shape_from_template(
            document=doc,
            slide_index=req.slide_index,
            shape_name=req.shape_name,
            new_text=req.new_text,
            left=req.left,
            top=req.top,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="複製命名 shape 成功")
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slide_fonts")
def ppt_slide_fonts(file_path: str, slide_index: int = 0):
    try:
        doc = open_presentation(file_path)
        return _ok(get_slide_text_fonts(doc, slide_index))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slides_fonts")
def ppt_slides_fonts(file_path: str):
    try:
        doc = open_presentation(file_path)
        return _ok(scan_presentation_text_fonts(doc))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/theme_info")
def ppt_theme_info(file_path: str):
    """讀取簡報佈景主題（theme）相關資訊；底層實作於 dataProcess.ppt_stdio.get_presentation_theme_info。"""
    try:
        doc = open_presentation(file_path)
        return _ok(get_presentation_theme_info(doc))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slide_background")
def ppt_slide_background(file_path: str, slide_index: int = 0):
    """讀取單頁投影片背景；底層實作於 get_slide_background_info。"""
    try:
        doc = open_presentation(file_path)
        return _ok(get_slide_background_info(doc, slide_index))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slides_backgrounds")
def ppt_slides_backgrounds(file_path: str):
    """掃描整份簡報各頁背景；底層實作於 scan_presentation_backgrounds。"""
    try:
        doc = open_presentation(file_path)
        return _ok(scan_presentation_backgrounds(doc))
    except Exception as e:
        _err_to_http(e)


# -------------------------
# PPT Save
# -------------------------

@app.post("/ppt/save_as")
def ppt_save_as(req: SaveAsRequest):
    try:
        doc = open_presentation(req.file_path)
        out_path = save(doc, req.save_as)
        return _ok({
            "file_path": out_path,
            "info": get_info(doc),
        }, message="另存成功")
    except Exception as e:
        _err_to_http(e)


# -------------------------
# Slide Operations
# -------------------------

@app.post("/ppt/add_blank_slide")
def ppt_add_blank_slide(req: AddBlankSlidesRequest):
    try:
        doc = open_presentation(req.file_path)
        slide_index = add_blank_slide(doc)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "added_slide_index": slide_index,
            "info": get_info(doc),
        }, message="新增空白頁成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_blank_slides")
def ppt_add_blank_slides(req: AddBlankSlidesRequest):
    try:
        doc = open_presentation(req.file_path)
        indices = add_blank_slides(doc, page_num=req.page_num)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "added_slide_indices": indices,
            "info": get_info(doc),
        }, message="新增多頁空白頁成功")
    except Exception as e:
        _err_to_http(e)


# -------------------------
# Shape Operations
# -------------------------

@app.post("/ppt/add_text")
def ppt_add_text(req: AddTextRequest):
    try:
        doc = open_presentation(req.file_path)

        font_color = None
        if req.font_color is not None:
            if len(req.font_color) != 3:
                raise ValueError("font_color 必須是 [R, G, B]")
            font_color = (req.font_color[0], req.font_color[1], req.font_color[2])

        result = add_text(
            document=doc,
            slide_index=req.slide_index,
            text=req.text,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_name=req.font_name,
            font_color=font_color,
            align=req.align,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增文字成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_wordart_like_textbox")
def ppt_add_wordart_like_textbox(req: AddWordartLikeTextboxRequest):
    try:
        doc = open_presentation(req.file_path)

        font_color = None
        if req.font_color is not None:
            if len(req.font_color) != 3:
                raise ValueError("font_color 格式必須是 [R, G, B]")
            font_color = (req.font_color[0], req.font_color[1], req.font_color[2])

        result = add_wordart_like_textbox(
            document=doc,
            slide_index=req.slide_index,
            text=req.text,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            font_size=req.font_size,
            bold=req.bold,
            font_name=req.font_name,
            font_color=font_color,
            align=req.align,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增 wordart_like_textbox 成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_wordart_text")
def ppt_update_wordart_text(req: UpdateWordartTextRequest):
    try:
        doc = open_presentation(req.file_path)
        result = update_wordart_text(
            document=doc,
            slide_index=req.slide_index,
            new_text=req.new_text,
            shape_name=req.shape_name,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="更新 wordart 文字成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_image")
def ppt_add_image(req: AddImageRequest):
    try:
        doc = open_presentation(req.file_path)
        result = add_image(
            document=doc,
            slide_index=req.slide_index,
            image_path=req.image_path,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            keep_aspect_ratio=req.keep_aspect_ratio,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增圖片成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_table")
def ppt_add_table(req: AddTableRequest):
    try:
        doc = open_presentation(req.file_path)
        result = add_table(
            document=doc,
            slide_index=req.slide_index,
            rows=req.rows,
            cols=req.cols,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            data=req.data,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增表格成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_shape")
def ppt_add_shape(req: AddShapeRequest):
    try:
        doc = open_presentation(req.file_path)

        fill_color = tuple(req.fill_color) if req.fill_color is not None else None
        line_color = tuple(req.line_color) if req.line_color is not None else None
        font_color = tuple(req.font_color) if req.font_color is not None else None

        result = add_shape(
            document=doc,
            slide_index=req.slide_index,
            shape_type=req.shape_type,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            text=req.text,
            fill_color=fill_color,
            line_color=line_color,
            line_width=req.line_width,
            font_size=req.font_size,
            bold=req.bold,
            font_name=req.font_name,
            font_color=font_color,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增圖形成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_line")
def ppt_add_line(req: AddLineRequest):
    try:
        doc = open_presentation(req.file_path)

        line_color = tuple(req.line_color) if req.line_color is not None else None

        result = add_line(
            document=doc,
            slide_index=req.slide_index,
            x1=req.x1,
            y1=req.y1,
            x2=req.x2,
            y2=req.y2,
            line_color=line_color,
            line_width=req.line_width,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增線段成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_arrow")
def ppt_add_arrow(req: AddArrowRequest):
    try:
        doc = open_presentation(req.file_path)

        fill_color = tuple(req.fill_color) if req.fill_color is not None else None
        line_color = tuple(req.line_color) if req.line_color is not None else None
        font_color = tuple(req.font_color) if req.font_color is not None else None

        result = add_arrow(
            document=doc,
            slide_index=req.slide_index,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            direction=req.direction,
            text=req.text,
            fill_color=fill_color,
            line_color=line_color,
            line_width=req.line_width,
            font_size=req.font_size,
            bold=req.bold,
            font_name=req.font_name,
            font_color=font_color,
        )
        out_path = save(doc, req.save_as or req.file_path)

        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增箭頭成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_slide")
def ppt_delete_slide(req: DeleteSlideRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_slide(doc, req.slide_index)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
            "slides": list_slides(doc),
        }, message="刪除頁面成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/duplicate_slide")
def ppt_duplicate_slide(req: DuplicateSlideRequest):
    try:
        doc = open_presentation(req.file_path)
        result = duplicate_slide(doc, req.slide_index)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
            "slides": list_slides(doc),
        }, message="複製頁面成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/replace_text")
def ppt_replace_text(req: ReplaceTextRequest):
    try:
        doc = open_presentation(req.file_path)
        result = replace_text(
            document=doc,
            old_text=req.old_text,
            new_text=req.new_text,
            slide_indices=req.slide_indices,
            exact_match=req.exact_match,
            case_sensitive=req.case_sensitive,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="取代文字成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_bullets")
def ppt_add_bullets(req: AddBulletsRequest):
    try:
        doc = open_presentation(req.file_path)

        font_color = None
        if req.font_color is not None:
            if len(req.font_color) != 3:
                raise ValueError("font_color 必須是 [R, G, B]")
            font_color = (req.font_color[0], req.font_color[1], req.font_color[2])

        result = add_bullets(
            document=doc,
            slide_index=req.slide_index,
            items=req.items,
            left=req.left,
            top=req.top,
            width=req.width,
            height=req.height,
            font_size=req.font_size,
            level=req.level,
            bold=req.bold,
            font_name=req.font_name,
            font_color=font_color,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增項目符號成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_title_slide")
def ppt_add_title_slide(req: AddTitleSlideRequest):
    try:
        doc = open_presentation(req.file_path)
        result = add_title_slide(
            document=doc,
            title=req.title,
            subtitle=req.subtitle,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
            "slides": list_slides(doc),
        }, message="新增標題頁成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/reorder_slides")
def ppt_reorder_slides(req: ReorderSlidesRequest):
    try:
        doc = open_presentation(req.file_path)
        result = reorder_slides(doc, req.new_order)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
            "slides": list_slides(doc),
        }, message="重排頁面成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_slides_background_color")
def ppt_set_slides_background_color(req: SetSlidesBackgroundColorRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_slides_background_color(doc, req.slide_indices, tuple(req.rgb))
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="設定多頁背景顏色成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_slides_background_image")
def ppt_set_slides_background_image(req: SetSlidesBackgroundImageRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_slides_background_image(doc, req.slide_indices, req.image_path)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="設定多頁背景圖片成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_slide_background_color")
def ppt_set_slide_background_color(req: SetSlideBackgroundColorRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_slide_background_color(doc, req.slide_index, req.rgb)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="設定頁面背景顏色成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_slide_background_image")
def ppt_set_slide_background_image(req: SetSlideBackgroundImageRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_slide_background_image(doc, req.slide_index, req.image_path)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="設定頁面背景圖片成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/render_slide_to_image")
def ppt_render_slide_to_image(req: RenderSlideToImageRequest):
    try:
        result = render_slide_to_image(
            pptx_path=req.file_path,
            slide_index=req.slide_index,
            output_path=req.output_path,
            dpi=req.dpi,
            libreoffice_path=req.libreoffice_path,
        )
        return _ok(result, message="投影片截圖成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/render_slides_to_grid_image")
def ppt_render_slides_to_grid_image(req: RenderSlidesToGridImageRequest):
    try:
        result = render_slides_to_grid_image(
            pptx_path=req.file_path,
            slide_indices=req.slide_indices,
            output_path=req.output_path,
            cols=req.cols,
            dpi=req.dpi,
            libreoffice_path=req.libreoffice_path,
            add_page_title=req.add_page_title,
            figure_title=req.figure_title,
        )
        return _ok(result, message="多頁拼圖成功")
    except Exception as e:
        _err_to_http(e)


if __name__ == "__main__":
    import uvicorn

    server_config = _load_server_config()
    host = str(server_config.get("hostIP", "10.1.3.127"))
    port = int(server_config.get("hostPort", 6414))

    uvicorn.run("api_server:app", host=host, port=port, reload=False)
