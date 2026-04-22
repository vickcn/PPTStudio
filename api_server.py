# -*- coding: utf-8 -*-
"""
api_server.py

FastAPI PPT

Usage:
    uvicorn api_server:app --host 0.0.0.0 --port 8010

Dependencies:
    pip install fastapi uvicorn python-pptx pillow
"""

import os
import json
import logging
import traceback
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

import dataProcess.ppt_stdio as ppt_stdio_mod
try:
    from dataProcess.file_importers import FileImportManager
except Exception:
    FileImportManager = None

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
    list_slide_tables,
    get_table_detail,
    update_table_cell,
    set_table_cell_style,
    update_table_row,
    update_table_column,
    set_table_row_style,
    set_table_column_style,
    set_table_row_height,
    set_table_column_width,
    distribute_table_column_widths,
    distribute_table_row_heights,
    delete_table,
    rebuild_table_with_modified_structure,
    insert_table_row,
    delete_table_row,
    insert_table_column,
    delete_table_column,
    add_line,
    add_arrow,
    add_shape,
    get_info,
    list_slides,
    get_textbox_style,
    get_slide_textbox_styles,
    set_textbox_style,
    drag_shape,
    reorder_shape_layer,
    drag_textbox,
    delete_textbox,
    get_slide_animations,
    get_shape_animations,
    add_shape_animation,
    update_shape_animation,
    delete_shape_animation,
    clear_shape_animations,
    clear_slide_animations,
    reorder_slide_animations,
    get_slide_transition,
    set_slide_transition,
    clear_slide_transition,
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
    set_all_slides_background_color,
    set_all_slides_background_image,
    delete_slide,
    duplicate_slide,
    replace_text,
    add_bullets,
    add_title_slide,
    reorder_slides,
    render_slide_to_image,
    render_slides_to_grid_image,
    parse_math_expression,
    add_equation,
    add_equation_omml,
    update_equation,
    update_equation_omml,
    delete_equation,
    delete_equation_omml,
)


app = FastAPI(
    title="PPT API Server",
    version="0.1.0",
    description="提供 PPTX 建立與編輯的 API，之後可再包成 MCP tools",
)

logger = logging.getLogger("ppt_api")
_LOG_LEVEL_NAME = os.environ.get("PPT_API_LOG_LEVEL", "WARNING").upper()
_LOG_LEVEL = getattr(logging, _LOG_LEVEL_NAME, logging.WARNING)
if not logging.getLogger().handlers:
    logging.basicConfig(level=_LOG_LEVEL)
logger.setLevel(_LOG_LEVEL)
logger.info("Loaded ppt_stdio module: %s", getattr(ppt_stdio_mod, "__file__", "<unknown>"))

_file_import_manager = FileImportManager() if FileImportManager is not None else None


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


class _TableShapeLocator(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class UpdateTableCellRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    col_idx: int = Field(..., ge=0)
    text: Optional[str] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None
    border_color: Optional[List[int]] = None
    border_width: Optional[float] = None
    border_style: Optional[str] = None
    clear_text: bool = False


class SetTableCellStyleRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    col_idx: int = Field(..., ge=0)
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None
    border_color: Optional[List[int]] = None
    border_width: Optional[float] = None
    border_style: Optional[str] = None


class UpdateTableRowRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    row_text: Optional[str] = None
    cell_texts: Optional[List[str]] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None


class UpdateTableColumnRequest(_TableShapeLocator):
    col_idx: int = Field(..., ge=0)
    column_text: Optional[str] = None
    cell_texts: Optional[List[str]] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None


class SetTableRowStyleRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None
    border_color: Optional[List[int]] = None
    border_width: Optional[float] = None
    border_style: Optional[str] = None


class SetTableColumnStyleRequest(_TableShapeLocator):
    col_idx: int = Field(..., ge=0)
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = None
    fill_color: Optional[List[int]] = None
    h_align: Optional[str] = None
    v_align: Optional[str] = None
    border_color: Optional[List[int]] = None
    border_width: Optional[float] = None
    border_style: Optional[str] = None


class SetTableRowHeightRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    height_emu: int = Field(..., gt=0)


class SetTableColumnWidthRequest(_TableShapeLocator):
    col_idx: int = Field(..., ge=0)
    width_emu: int = Field(..., gt=0)


class DistributeTableColumnWidthsRequest(_TableShapeLocator):
    column_indices: Optional[List[int]] = None


class DistributeTableRowHeightsRequest(_TableShapeLocator):
    row_indices: Optional[List[int]] = None


class RebuildTableStructureRequest(_TableShapeLocator):
    new_rows: int = Field(..., ge=1)
    new_cols: int = Field(..., ge=1)
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)


class InsertTableRowRequest(_TableShapeLocator):
    insert_before: int = Field(..., ge=0)
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)


class DeleteTableRowRequest(_TableShapeLocator):
    row_idx: int = Field(..., ge=0)
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)


class InsertTableColumnRequest(_TableShapeLocator):
    insert_before: int = Field(..., ge=0)
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)


class DeleteTableColumnRequest(_TableShapeLocator):
    col_idx: int = Field(..., ge=0)
    first_row_as_header: bool = False
    font_size: int = Field(14, gt=0)


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


class ParseMathExpressionRequest(BaseModel):
    """Parse spoken math or LaTeX only; does not modify PPTX."""
    input_text: str = Field(..., description="Raw input text")
    input_type: str = Field("latex", description="spoken or latex")


class EquationTextRun(BaseModel):
    text: str
    font_name: Optional[str] = None
    font_size: Optional[int] = Field(None, gt=0)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[List[int]] = Field(None, min_length=3, max_length=3)


class AddEquationRequest(BaseModel):
    """Add equation; default is OMML, image mode keeps M1 compatibility."""
    file_path: str
    slide_index: int = Field(..., ge=0)
    input_text: str = Field(..., description="Spoken math or LaTeX")
    input_type: str = Field("latex", description="spoken or latex")
    left: int = Field(..., ge=0)
    top: int = Field(..., ge=0)
    width: Optional[int] = Field(None, gt=0)
    height: Optional[int] = Field(None, gt=0)
    font_size: Optional[int] = Field(None, gt=0)
    color: Optional[List[int]] = Field(None, description="RGB color")
    prefix_runs: Optional[List[EquationTextRun]] = None
    suffix_runs: Optional[List[EquationTextRun]] = None
    render_mode: str = Field("omml", description="omml or image")
    save_as: Optional[str] = None


class UpdateEquationRequest(BaseModel):
    """Update equation; default is OMML, image mode updates image equations."""
    file_path: str
    input_text: str
    input_type: str = Field("latex", description="spoken or latex")
    expr_id: Optional[str] = None
    shape_id: Optional[int] = None
    slide_index: Optional[int] = Field(None, ge=0, description="Optional hint for locating target")
    prefix_runs: Optional[List[EquationTextRun]] = None
    suffix_runs: Optional[List[EquationTextRun]] = None
    render_mode: str = Field("omml", description="omml or image")
    save_as: Optional[str] = None


class DeleteEquationRequest(BaseModel):
    """Delete equation; default is OMML, image mode deletes image equations."""
    file_path: str
    expr_id: Optional[str] = None
    shape_id: Optional[int] = None
    slide_index: Optional[int] = Field(None, ge=0)
    render_mode: str = Field("omml", description="omml or image")
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


class SetAllSlidesBackgroundColorRequest(BaseModel):
    file_path: str
    rgb: List[int] = Field(..., min_length=3, max_length=3)
    save_as: Optional[str] = None


class SetSlidesBackgroundImageRequest(BaseModel):
    file_path: str
    slide_indices: List[int]
    image_path: str
    save_as: Optional[str] = None


class SetAllSlidesBackgroundImageRequest(BaseModel):
    file_path: str
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


class DragTextboxRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    left: Optional[int] = Field(None, ge=0, description="Horizontal offset in EMUs")
    top: Optional[int] = Field(None, ge=0, description="Vertical offset in EMUs")
    delta_x: Optional[int] = Field(None, description="Horizontal delta in EMUs")
    delta_y: Optional[int] = Field(None, description="Vertical delta in EMUs")
    width: Optional[int] = Field(None, gt=0, description="New width in EMUs")
    height: Optional[int] = Field(None, gt=0, description="New height in EMUs")
    save_as: Optional[str] = None


class DragShapeRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    left: Optional[int] = Field(None, ge=0)
    top: Optional[int] = Field(None, ge=0)
    delta_x: Optional[int] = None
    delta_y: Optional[int] = None
    width: Optional[int] = Field(None, gt=0)
    height: Optional[int] = Field(None, gt=0)
    save_as: Optional[str] = None


class ReorderShapeLayerRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    action: str = Field(..., description="to_front | to_back | forward | backward")
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class AddShapeAnimationRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    effect_type: str = "fade"
    trigger: str = "on_click"
    duration_ms: int = Field(500, gt=0)
    delay_ms: int = Field(0, ge=0)
    save_as: Optional[str] = None


class UpdateShapeAnimationRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    animation_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    effect_type: Optional[str] = None
    trigger: Optional[str] = None
    duration_ms: Optional[int] = Field(None, gt=0)
    delay_ms: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class DeleteShapeAnimationRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    animation_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class ClearShapeAnimationsRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    shape_id: Optional[int] = None
    shape_index: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class ClearSlideAnimationsRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    save_as: Optional[str] = None


class ReorderSlideAnimationsRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    new_order: List[int]
    save_as: Optional[str] = None


class SetSlideTransitionRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
    transition_type: str = "fade"
    duration_ms: Optional[int] = Field(None, ge=0)
    advance_on_click: bool = True
    advance_after_ms: Optional[int] = Field(None, ge=0)
    save_as: Optional[str] = None


class ClearSlideTransitionRequest(BaseModel):
    file_path: str
    slide_index: int = Field(..., ge=0)
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


class ImportFileRequest(BaseModel):
    file_path: str
    file_type: Optional[str] = None
    options: Optional[Dict[str, Any]] = None


class ProcessFileRequest(BaseModel):
    file_path: str
    filename: Optional[str] = None
    options: Optional[Dict[str, Any]] = None


class RunStageExtractRequest(BaseModel):
    local_path: str
    filename: Optional[str] = None
    config: Optional[Dict[str, Any]] = None


class RunStageSegmentRequest(BaseModel):
    import_result: Dict[str, Any]
    options: Optional[Dict[str, Any]] = None


class RunStageChunkRequest(BaseModel):
    segment_result: Dict[str, Any]
    options: Optional[Dict[str, Any]] = None


class RunParserPipelineRequest(BaseModel):
    file_path: str
    file_type: Optional[str] = None
    import_options: Optional[Dict[str, Any]] = None
    segment_options: Optional[Dict[str, Any]] = None
    chunk_options: Optional[Dict[str, Any]] = None


def _tuple3_opt(v: Optional[List[int]]) -> Optional[tuple]:
    if v is None:
        return None
    if len(v) != 3:
        raise ValueError("RGB 必須為長度 3 的整數陣列")
    return (int(v[0]), int(v[1]), int(v[2]))


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


def _serialize_runs(runs: Optional[List[EquationTextRun]]) -> Optional[List[Dict[str, Any]]]:
    if not runs:
        return None
    result: List[Dict[str, Any]] = []
    for run in runs:
        if hasattr(run, "dict"):
            result.append(run.dict(exclude_none=True))
        else:
            result.append(dict(run))
    return result


def _get_file_import_manager() -> Any:
    if _file_import_manager is None:
        raise RuntimeError("file_importers is not available in this environment")
    return _file_import_manager


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
# Context Parser / File Importers
# -------------------------

@app.post("/ppt/import_file")
def ppt_import_file(req: ImportFileRequest):
    try:
        manager = _get_file_import_manager()
        options = req.options or {}
        result = manager.import_file(
            file_path=req.file_path,
            file_type=req.file_type,
            **options,
        )
        return _ok(result, message="import_file completed")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/process_file")
def ppt_process_file(req: ProcessFileRequest):
    try:
        manager = _get_file_import_manager()
        options = req.options or {}
        ret: Dict[str, Any] = {}
        ok = manager.process_file(
            ret=ret,
            local_path=req.file_path,
            filename=req.filename,
            **options,
        )
        if not ok:
            raise RuntimeError(ret.get("error", "process_file failed"))
        return _ok(ret, message="process_file completed")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/run_stage_extract")
def ppt_run_stage_extract(req: RunStageExtractRequest):
    try:
        manager = _get_file_import_manager()
        result = manager.run_stage_extract(
            local_path=req.local_path,
            filename=req.filename,
            config=req.config,
        )
        if not result.get("success", False):
            raise RuntimeError(result.get("error", "run_stage_extract failed"))
        return _ok(result, message="run_stage_extract completed")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/run_stage_segment")
def ppt_run_stage_segment(req: RunStageSegmentRequest):
    try:
        manager = _get_file_import_manager()
        options = req.options or {}
        result = manager.run_stage_segment(req.import_result, **options)
        if not result.get("success", False):
            raise RuntimeError(result.get("error", "run_stage_segment failed"))
        return _ok(result, message="run_stage_segment completed")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/run_stage_chunk")
def ppt_run_stage_chunk(req: RunStageChunkRequest):
    try:
        manager = _get_file_import_manager()
        options = req.options or {}
        result = manager.run_stage_chunk(req.segment_result, **options)
        if not result.get("success", False):
            raise RuntimeError(result.get("error", "run_stage_chunk failed"))
        return _ok(result, message="run_stage_chunk completed")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/run_parser_pipeline")
def ppt_run_parser_pipeline(req: RunParserPipelineRequest):
    try:
        manager = _get_file_import_manager()
        import_options = req.import_options or {}
        segment_options = req.segment_options or {}
        chunk_options = req.chunk_options or {}

        import_result = manager.import_file(
            file_path=req.file_path,
            file_type=req.file_type,
            **import_options,
        )
        if not import_result.get("success", False):
            raise RuntimeError(import_result.get("error", "import_file failed"))

        segment_result = manager.run_stage_segment(import_result, **segment_options)
        if not segment_result.get("success", False):
            raise RuntimeError(segment_result.get("error", "run_stage_segment failed"))

        chunk_result = manager.run_stage_chunk(segment_result, **chunk_options)
        if not chunk_result.get("success", False):
            raise RuntimeError(chunk_result.get("error", "run_stage_chunk failed"))

        return _ok(
            {
                "import_result": import_result,
                "segment_result": segment_result,
                "chunk_result": chunk_result,
            },
            message="run_parser_pipeline completed",
        )
    except Exception as e:
        _err_to_http(e)


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


@app.post("/ppt/drag_textbox")
def ppt_drag_textbox(req: DragTextboxRequest):
    try:
        doc = open_presentation(req.file_path)
        result = drag_textbox(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            left=req.left,
            top=req.top,
            delta_x=req.delta_x,
            delta_y=req.delta_y,
            width=req.width,
            height=req.height,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="拖曳文字框成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/drag_shape")
def ppt_drag_shape(req: DragShapeRequest):
    try:
        doc = open_presentation(req.file_path)
        result = drag_shape(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            left=req.left,
            top=req.top,
            delta_x=req.delta_x,
            delta_y=req.delta_y,
            width=req.width,
            height=req.height,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="shape 已拖拉")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/reorder_shape_layer")
def ppt_reorder_shape_layer(req: ReorderShapeLayerRequest):
    try:
        doc = open_presentation(req.file_path)
        result = reorder_shape_layer(
            document=doc,
            slide_index=req.slide_index,
            action=req.action,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="shape 圖層調整成功")
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


@app.get("/ppt/slide_animations")
def ppt_slide_animations(file_path: str, slide_index: int = 0):
    try:
        doc = open_presentation(file_path)
        return _ok(get_slide_animations(doc, slide_index))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/shape_animations")
def ppt_shape_animations(
        file_path: str,
        slide_index: int = 0,
        shape_id: Optional[int] = None,
        shape_index: Optional[int] = None,
    ):
    try:
        doc = open_presentation(file_path)
        return _ok(
            get_shape_animations(
                doc,
                slide_index=slide_index,
                shape_id=shape_id,
                shape_index=shape_index,
            )
        )
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_shape_animation")
def ppt_add_shape_animation(req: AddShapeAnimationRequest):
    try:
        doc = open_presentation(req.file_path)
        result = add_shape_animation(
            document=doc,
            slide_index=req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            effect_type=req.effect_type,
            trigger=req.trigger,
            duration_ms=req.duration_ms,
            delay_ms=req.delay_ms,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="新增 shape 動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_shape_animation")
def ppt_update_shape_animation(req: UpdateShapeAnimationRequest):
    try:
        doc = open_presentation(req.file_path)
        result = update_shape_animation(
            document=doc,
            slide_index=req.slide_index,
            animation_index=req.animation_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            effect_type=req.effect_type,
            trigger=req.trigger,
            duration_ms=req.duration_ms,
            delay_ms=req.delay_ms,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="更新 shape 動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_shape_animation")
def ppt_delete_shape_animation(req: DeleteShapeAnimationRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_shape_animation(
            document=doc,
            slide_index=req.slide_index,
            animation_index=req.animation_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="刪除 shape 動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/clear_shape_animations")
def ppt_clear_shape_animations(req: ClearShapeAnimationsRequest):
    try:
        doc = open_presentation(req.file_path)
        result = clear_shape_animations(
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
        }, message="清除 shape 動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/clear_slide_animations")
def ppt_clear_slide_animations(req: ClearSlideAnimationsRequest):
    try:
        doc = open_presentation(req.file_path)
        result = clear_slide_animations(document=doc, slide_index=req.slide_index)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="清除投影片動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/reorder_slide_animations")
def ppt_reorder_slide_animations(req: ReorderSlideAnimationsRequest):
    try:
        doc = open_presentation(req.file_path)
        result = reorder_slide_animations(
            document=doc,
            slide_index=req.slide_index,
            new_order=req.new_order,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="重排投影片動畫成功")
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/slide_transition")
def ppt_slide_transition(file_path: str, slide_index: int = 0):
    try:
        doc = open_presentation(file_path)
        return _ok(get_slide_transition(doc, slide_index=slide_index))
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_slide_transition")
def ppt_set_slide_transition(req: SetSlideTransitionRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_slide_transition(
            document=doc,
            slide_index=req.slide_index,
            transition_type=req.transition_type,
            duration_ms=req.duration_ms,
            advance_on_click=req.advance_on_click,
            advance_after_ms=req.advance_after_ms,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="設定投影片轉場成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/clear_slide_transition")
def ppt_clear_slide_transition(req: ClearSlideTransitionRequest):
    try:
        doc = open_presentation(req.file_path)
        result = clear_slide_transition(document=doc, slide_index=req.slide_index)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="清除投影片轉場成功")
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


@app.get("/ppt/slide_tables")
def ppt_slide_tables(file_path: str, slide_index: int = 0):
    try:
        doc = open_presentation(file_path)
        return _ok(list_slide_tables(doc, slide_index))
    except Exception as e:
        _err_to_http(e)


@app.get("/ppt/table_detail")
def ppt_table_detail(file_path: str, slide_index: int, shape_id: Optional[int] = None, shape_index: Optional[int] = None):
    try:
        doc = open_presentation(file_path)
        return _ok(get_table_detail(doc, slide_index, shape_id=shape_id, shape_index=shape_index))
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_table_cell")
def ppt_update_table_cell(req: UpdateTableCellRequest):
    try:
        doc = open_presentation(req.file_path)
        result = update_table_cell(
            doc,
            req.slide_index,
            req.row_idx,
            req.col_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            text=req.text,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
            border_color=_tuple3_opt(req.border_color),
            border_width=req.border_width,
            border_style=req.border_style,
            clear_text=req.clear_text,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="更新儲存格成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_table_cell_style")
def ppt_set_table_cell_style(req: SetTableCellStyleRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_table_cell_style(
            doc,
            req.slide_index,
            req.row_idx,
            req.col_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
            border_color=_tuple3_opt(req.border_color),
            border_width=req.border_width,
            border_style=req.border_style,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="設定儲存格樣式成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_table_row")
def ppt_update_table_row(req: UpdateTableRowRequest):
    try:
        doc = open_presentation(req.file_path)
        result = update_table_row(
            doc,
            req.slide_index,
            req.row_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            row_text=req.row_text,
            cell_texts=req.cell_texts,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="更新整列成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_table_column")
def ppt_update_table_column(req: UpdateTableColumnRequest):
    try:
        doc = open_presentation(req.file_path)
        result = update_table_column(
            doc,
            req.slide_index,
            req.col_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            column_text=req.column_text,
            cell_texts=req.cell_texts,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="更新整欄成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_table_row_style")
def ppt_set_table_row_style(req: SetTableRowStyleRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_table_row_style(
            doc,
            req.slide_index,
            req.row_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
            border_color=_tuple3_opt(req.border_color),
            border_width=req.border_width,
            border_style=req.border_style,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="設定整列樣式成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_table_column_style")
def ppt_set_table_column_style(req: SetTableColumnStyleRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_table_column_style(
            doc,
            req.slide_index,
            req.col_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            font_name=req.font_name,
            font_size=req.font_size,
            bold=req.bold,
            italic=req.italic,
            font_color=_tuple3_opt(req.font_color),
            fill_color=_tuple3_opt(req.fill_color),
            h_align=req.h_align,
            v_align=req.v_align,
            border_color=_tuple3_opt(req.border_color),
            border_width=req.border_width,
            border_style=req.border_style,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="設定整欄樣式成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_table_row_height")
def ppt_set_table_row_height(req: SetTableRowHeightRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_table_row_height(
            doc,
            req.slide_index,
            req.row_idx,
            req.height_emu,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="設定列高成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/set_table_column_width")
def ppt_set_table_column_width(req: SetTableColumnWidthRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_table_column_width(
            doc,
            req.slide_index,
            req.col_idx,
            req.width_emu,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="設定欄寬成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/distribute_table_column_widths")
def ppt_distribute_table_column_widths(req: DistributeTableColumnWidthsRequest):
    try:
        doc = open_presentation(req.file_path)
        result = distribute_table_column_widths(
            doc,
            req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            column_indices=req.column_indices,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="平均欄寬成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/distribute_table_row_heights")
def ppt_distribute_table_row_heights(req: DistributeTableRowHeightsRequest):
    try:
        doc = open_presentation(req.file_path)
        result = distribute_table_row_heights(
            doc,
            req.slide_index,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            row_indices=req.row_indices,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="平均列高成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_table")
def ppt_delete_table(req: _TableShapeLocator):
    try:
        doc = open_presentation(req.file_path)
        result = delete_table(doc, req.slide_index, shape_id=req.shape_id, shape_index=req.shape_index)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="刪除表格成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/rebuild_table_structure")
def ppt_rebuild_table_structure(req: RebuildTableStructureRequest):
    try:
        doc = open_presentation(req.file_path)
        result = rebuild_table_with_modified_structure(
            doc,
            req.slide_index,
            req.new_rows,
            req.new_cols,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="重建表格成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/insert_table_row")
def ppt_insert_table_row(req: InsertTableRowRequest):
    try:
        doc = open_presentation(req.file_path)
        result = insert_table_row(
            doc,
            req.slide_index,
            req.insert_before,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="插入列成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_table_row")
def ppt_delete_table_row(req: DeleteTableRowRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_table_row(
            doc,
            req.slide_index,
            req.row_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="刪除列成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/insert_table_column")
def ppt_insert_table_column(req: InsertTableColumnRequest):
    try:
        doc = open_presentation(req.file_path)
        result = insert_table_column(
            doc,
            req.slide_index,
            req.insert_before,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="插入欄成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_table_column")
def ppt_delete_table_column(req: DeleteTableColumnRequest):
    try:
        doc = open_presentation(req.file_path)
        result = delete_table_column(
            doc,
            req.slide_index,
            req.col_idx,
            shape_id=req.shape_id,
            shape_index=req.shape_index,
            first_row_as_header=req.first_row_as_header,
            font_size=req.font_size,
        )
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({"file_path": out_path, "result": result, "info": get_info(doc)}, message="刪除欄成功")
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


@app.post("/ppt/set_all_slides_background_color")
def ppt_set_all_slides_background_color(req: SetAllSlidesBackgroundColorRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_all_slides_background_color(doc, tuple(req.rgb))
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="套用全部背景顏色成功")
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


@app.post("/ppt/set_all_slides_background_image")
def ppt_set_all_slides_background_image(req: SetAllSlidesBackgroundImageRequest):
    try:
        doc = open_presentation(req.file_path)
        result = set_all_slides_background_image(doc, req.image_path)
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message="套用全部背景圖片成功")
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


@app.post("/ppt/parse_math_expression")
def ppt_parse_math_expression(req: ParseMathExpressionRequest):
    """解析數學公式：默認使用 OMML 路徑，image 模式保持與 M1 兼容。"""
    try:
        data = parse_math_expression(req.input_text, req.input_type)
        return _ok(data, message="解析數學公式成功")
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/add_equation")
def ppt_add_equation(req: AddEquationRequest):
    """Add equation: default OMML path, image mode keeps M1 compatibility."""
    try:
        doc = open_presentation(req.file_path)
        color_rgb = tuple(req.color) if req.color is not None else None
        mode = (req.render_mode or "omml").strip().lower()
        logger.info(
            "add_equation request mode=%s file=%s slide=%s input_type=%s text=%s",
            mode,
            req.file_path,
            req.slide_index,
            req.input_type,
            (req.input_text or "")[:120],
        )

        if mode == "image":
            result = add_equation(
                document=doc,
                slide_index=req.slide_index,
                input_text=req.input_text,
                input_type=req.input_type,
                left=req.left,
                top=req.top,
                width=req.width,
                height=req.height,
                font_size=req.font_size,
                color=color_rgb,
            )
            msg = "add equation (image) success"
        else:
            prefix_runs = _serialize_runs(req.prefix_runs)
            suffix_runs = _serialize_runs(req.suffix_runs)
            result = add_equation_omml(
                document=doc,
                slide_index=req.slide_index,
                input_text=req.input_text,
                input_type=req.input_type,
                left=req.left,
                top=req.top,
                width=req.width,
                height=req.height,
                font_size=req.font_size,
                color=color_rgb,
                prefix_runs=prefix_runs,
                suffix_runs=suffix_runs,
            )
            msg = "add equation (omml) success"

        if isinstance(result, dict):
            omml_ref = result.get("omml_fragment_ref")
            logger.info(
                "add_equation result mode=%s expr_id=%s shape_id=%s slide=%s render_mode=%s omml_write_mode=%s omml_prefix=%s",
                mode,
                result.get("expr_id"),
                result.get("shape_id"),
                result.get("slide_index"),
                result.get("render_mode"),
                result.get("omml_write_mode"),
                (omml_ref[:160] if isinstance(omml_ref, str) else None),
            )

        out_path = save(doc, req.save_as or req.file_path)
        return _ok(
            {
                "file_path": out_path,
                "result": result,
                "info": get_info(doc),
            },
            message=msg,
        )
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/update_equation")
def ppt_update_equation(req: UpdateEquationRequest):
    """更新數學公式：默認使用 OMML 路徑，image 模式更新圖像公式。"""
    try:
        doc = open_presentation(req.file_path)
        mode = (req.render_mode or "omml").strip().lower()
        if mode == "image":
            result = update_equation(
                document=doc,
                input_text=req.input_text,
                input_type=req.input_type,
                expr_id=req.expr_id,
                shape_id=req.shape_id,
                slide_index=req.slide_index,
            )
            msg = "更新數學公式 (image) 成功"
        else:
            prefix_runs = _serialize_runs(req.prefix_runs)
            suffix_runs = _serialize_runs(req.suffix_runs)
            result = update_equation_omml(
                document=doc,
                input_text=req.input_text,
                input_type=req.input_type,
                expr_id=req.expr_id,
                shape_id=req.shape_id,
                slide_index=req.slide_index,
                prefix_runs=prefix_runs,
                suffix_runs=suffix_runs,
            )
            msg = "更新數學公式 (OMML) 成功"
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message=msg)
    except Exception as e:
        _err_to_http(e)


@app.post("/ppt/delete_equation")
def ppt_delete_equation(req: DeleteEquationRequest):
    """刪除數學公式：默認使用 OMML 路徑，image 模式刪除圖像公式。"""
    try:
        doc = open_presentation(req.file_path)
        mode = (req.render_mode or "omml").strip().lower()
        if mode == "image":
            result = delete_equation(
                document=doc,
                expr_id=req.expr_id,
                shape_id=req.shape_id,
                slide_index=req.slide_index,
            )
            msg = "刪除數學公式 (image) 成功"
        else:
            result = delete_equation_omml(
                document=doc,
                expr_id=req.expr_id,
                shape_id=req.shape_id,
                slide_index=req.slide_index,
            )
            msg = "刪除數學公式 (OMML) 成功"
        out_path = save(doc, req.save_as or req.file_path)
        return _ok({
            "file_path": out_path,
            "result": result,
            "info": get_info(doc),
        }, message=msg)
    except Exception as e:
        _err_to_http(e)


if __name__ == "__main__":
    import uvicorn

    server_config = _load_server_config()
    host = str(server_config.get("hostIP", "10.1.3.127"))
    port = int(server_config.get("hostPort", 6414))

    uvicorn.run("api_server:app", host=host, port=port, reload=False)
