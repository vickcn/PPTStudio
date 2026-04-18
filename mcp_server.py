# -*- coding: utf-8 -*-
"""
mcp_server.py

Usage:
FastAPI PPT API MCP server using stdio transport for Cursor / Claude Desktop / MCP Gateway using streamable-http transport.
Dependencies:
    pip install "mcp[cli]" httpx

Commands:
    python mcp_server.py
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport stdio
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport streamable-http --host 10.1.3.127 --port 6414

Usage:
1. Start FastAPI PPT API server:       uvicorn api_server:app --host 10.1.3.127 --port 6414
2. Start MCP server:
    python mcp_server.py
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport stdio
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport streamable-http --host 10.1.3.127 --port 6414
"""

from __future__ import annotations

import argparse
import os
from typing import Any, Dict, List, Optional

import httpx
from mcp.server.fastmcp import FastMCP


DEFAULT_API_BASE = os.getenv("PPT_API_BASE", "http://10.1.3.127:6414")
DEFAULT_TIMEOUT = float(os.getenv("PPT_API_TIMEOUT", "180"))

mcp = FastMCP(
    "ppt",
    instructions=(
        "MCP server for PPT operations backed by an existing FastAPI PPT API. "
        "Use these tools to create, inspect, edit, reorder, render, and save PPTX files."
    ),
)

_API_BASE = DEFAULT_API_BASE.rstrip("/")
_TIMEOUT = DEFAULT_TIMEOUT


def set_runtime_config(api_base: str, timeout: float) -> None:
    global _API_BASE, _TIMEOUT
    _API_BASE = api_base.rstrip("/")
    _TIMEOUT = timeout


async def _request(method: str, path: str, *, params: Optional[Dict[str, Any]] = None, json_body: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    url = f"{_API_BASE}{path}"
    async with httpx.AsyncClient(timeout=_TIMEOUT) as client:
        resp = await client.request(method, url, params=params, json=json_body)

    try:
        data = resp.json()
    except Exception:
        resp.raise_for_status()
        return {
            "ok": False,
            "status_code": resp.status_code,
            "text": resp.text,
            "url": url,
        }

    if resp.is_success:
        return data

    return {
        "ok": False,
        "status_code": resp.status_code,
        "url": url,
        "error": data,
    }


def _clean_rgb(rgb: Optional[List[int]]) -> Optional[List[int]]:
    if rgb is None:
        return None
    if len(rgb) != 3:
        raise ValueError("rgb / color must be [R, G, B]")
    vals = [int(v) for v in rgb]
    for v in vals:
        if v < 0 or v > 255:
            raise ValueError("rgb / color must be between 0 and 255")
    return vals


@mcp.tool()
async def health() -> Dict[str, Any]:
    """Check whether the PPT API server is healthy."""
    return await _request("GET", "/health")


@mcp.tool()
async def root_info() -> Dict[str, Any]:
    """Get basic root info from the PPT API server."""
    return await _request("GET", "/")


@mcp.tool()
async def import_file(
    file_path: str,
    file_type: Optional[str] = None,
    options: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Import file via file_importers (ppt_parser-backed for .pptx)."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "file_type": file_type,
        "options": options,
    }
    return await _request("POST", "/ppt/import_file", json_body=body)


@mcp.tool()
async def process_file(
    file_path: str,
    filename: Optional[str] = None,
    options: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Process file via file_importers.process_file."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "filename": filename,
        "options": options,
    }
    return await _request("POST", "/ppt/process_file", json_body=body)


@mcp.tool()
async def run_stage_extract(
    local_path: str,
    filename: Optional[str] = None,
    config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Run extract stage via file_importers.run_stage_extract."""
    body: Dict[str, Any] = {
        "local_path": local_path,
        "filename": filename,
        "config": config,
    }
    return await _request("POST", "/ppt/run_stage_extract", json_body=body)


@mcp.tool()
async def run_stage_segment(
    import_result: Dict[str, Any],
    options: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Run segment stage via file_importers.run_stage_segment."""
    body: Dict[str, Any] = {
        "import_result": import_result,
        "options": options,
    }
    return await _request("POST", "/ppt/run_stage_segment", json_body=body)


@mcp.tool()
async def run_stage_chunk(
    segment_result: Dict[str, Any],
    options: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Run chunk stage via file_importers.run_stage_chunk."""
    body: Dict[str, Any] = {
        "segment_result": segment_result,
        "options": options,
    }
    return await _request("POST", "/ppt/run_stage_chunk", json_body=body)


@mcp.tool()
async def run_parser_pipeline(
    file_path: str,
    file_type: Optional[str] = None,
    import_options: Optional[Dict[str, Any]] = None,
    segment_options: Optional[Dict[str, Any]] = None,
    chunk_options: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Run import->segment->chunk pipeline via file_importers."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "file_type": file_type,
        "import_options": import_options,
        "segment_options": segment_options,
        "chunk_options": chunk_options,
    }
    return await _request("POST", "/ppt/run_parser_pipeline", json_body=body)

@mcp.tool()
async def new_presentation(
    file_path: str,
    plank_page_num: int = 1,
    plank_page_width: int = 1080,
    plank_page_height: int = 1920,
    dpi: int = 96,
) -> Dict[str, Any]:
    """Create a new PPTX file."""
    return await _request(
        "POST",
        "/ppt/new",
        json_body={
            "file_path": file_path,
            "plank_page_num": plank_page_num,
            "plank_page_width": plank_page_width,
            "plank_page_height": plank_page_height,
            "dpi": dpi,
        },
    )


@mcp.tool()
async def get_presentation_info(file_path: str) -> Dict[str, Any]:
    """Get presentation info."""
    return await _request("GET", "/ppt/info", params={"file_path": file_path})


@mcp.tool()
async def list_presentation_slides(file_path: str) -> Dict[str, Any]:
    """List presentation slides."""
    return await _request("GET", "/ppt/slides", params={"file_path": file_path})


@mcp.tool()
async def save_as(file_path: str, save_as: str) -> Dict[str, Any]:
    """Save as."""
    return await _request(
        "POST",
        "/ppt/save_as",
        json_body={
            "file_path": file_path,
            "save_as": save_as,
        },
    )


@mcp.tool()
async def add_blank_slide(file_path: str, save_as: Optional[str] = None) -> Dict[str, Any]:
    """Add a blank slide."""
    body: Dict[str, Any] = {"file_path": file_path, "page_num": 1}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_blank_slide", json_body=body)


@mcp.tool()
async def add_blank_slides(file_path: str, page_num: int = 1, save_as: Optional[str] = None) -> Dict[str, Any]:
    """Add blank slides."""
    body: Dict[str, Any] = {"file_path": file_path, "page_num": page_num}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_blank_slides", json_body=body)


@mcp.tool()
async def add_text(
    file_path: str,
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
    font_color: Optional[List[int]] = None,
    align: str = "left",
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """Add text."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "text": text,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_name": font_name,
        "font_color": _clean_rgb(font_color),
        "align": align,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_text", json_body=body)


@mcp.tool()
async def add_image(
    file_path: str,
    slide_index: int,
    image_path: str,
    left: int,
    top: int,
    width: Optional[int] = None,
    height: Optional[int] = None,
    keep_aspect_ratio: bool = True,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """Add image."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "image_path": image_path,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "keep_aspect_ratio": keep_aspect_ratio,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_image", json_body=body)


@mcp.tool()
async def add_table(
    file_path: str,
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
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """Add table."""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "rows": rows,
        "cols": cols,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "data": data,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_table", json_body=body)


@mcp.tool()
async def slide_tables(file_path: str, slide_index: int = 0) -> Dict[str, Any]:
    """列出指定頁所有表格（GET /ppt/slide_tables）。"""
    return await _request(
        "GET",
        "/ppt/slide_tables",
        params={"file_path": file_path, "slide_index": slide_index},
    )


@mcp.tool()
async def table_detail(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
) -> Dict[str, Any]:
    """取得表格詳細（GET /ppt/table_detail）。"""
    params: Dict[str, Any] = {"file_path": file_path, "slide_index": slide_index}
    if shape_id is not None:
        params["shape_id"] = shape_id
    if shape_index is not None:
        params["shape_index"] = shape_index
    return await _request("GET", "/ppt/table_detail", params=params)


@mcp.tool()
async def update_table_cell(
    file_path: str,
    slide_index: int,
    row_idx: int,
    col_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    text: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    border_color: Optional[List[int]] = None,
    border_width: Optional[float] = None,
    border_style: Optional[str] = None,
    clear_text: bool = False,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """更新單一儲存格（POST /ppt/update_table_cell）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "col_idx": col_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
        "border_color": _clean_rgb(border_color) if border_color is not None else None,
        "border_width": border_width,
        "border_style": border_style,
        "clear_text": clear_text,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/update_table_cell", json_body=body)


@mcp.tool()
async def set_table_cell_style(
    file_path: str,
    slide_index: int,
    row_idx: int,
    col_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    border_color: Optional[List[int]] = None,
    border_width: Optional[float] = None,
    border_style: Optional[str] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定單格樣式（POST /ppt/set_table_cell_style）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "col_idx": col_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
        "border_color": _clean_rgb(border_color) if border_color is not None else None,
        "border_width": border_width,
        "border_style": border_style,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_table_cell_style", json_body=body)


@mcp.tool()
async def update_table_row(
    file_path: str,
    slide_index: int,
    row_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    row_text: Optional[str] = None,
    cell_texts: Optional[List[str]] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """更新整列（POST /ppt/update_table_row）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "row_text": row_text,
        "cell_texts": cell_texts,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/update_table_row", json_body=body)


@mcp.tool()
async def update_table_column(
    file_path: str,
    slide_index: int,
    col_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    column_text: Optional[str] = None,
    cell_texts: Optional[List[str]] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """更新整欄（POST /ppt/update_table_column）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "col_idx": col_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "column_text": column_text,
        "cell_texts": cell_texts,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/update_table_column", json_body=body)


@mcp.tool()
async def set_table_row_style(
    file_path: str,
    slide_index: int,
    row_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    border_color: Optional[List[int]] = None,
    border_width: Optional[float] = None,
    border_style: Optional[str] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定整列樣式（POST /ppt/set_table_row_style）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
        "border_color": _clean_rgb(border_color) if border_color is not None else None,
        "border_width": border_width,
        "border_style": border_style,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_table_row_style", json_body=body)


@mcp.tool()
async def set_table_column_style(
    file_path: str,
    slide_index: int,
    col_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[List[int]] = None,
    fill_color: Optional[List[int]] = None,
    h_align: Optional[str] = None,
    v_align: Optional[str] = None,
    border_color: Optional[List[int]] = None,
    border_width: Optional[float] = None,
    border_style: Optional[str] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定整欄樣式（POST /ppt/set_table_column_style）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "col_idx": col_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "h_align": h_align,
        "v_align": v_align,
        "border_color": _clean_rgb(border_color) if border_color is not None else None,
        "border_width": border_width,
        "border_style": border_style,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_table_column_style", json_body=body)


@mcp.tool()
async def distribute_table_column_widths(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    column_indices: Optional[List[int]] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """平均欄寬（POST /ppt/distribute_table_column_widths）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "column_indices": column_indices,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/distribute_table_column_widths", json_body=body)


@mcp.tool()
async def distribute_table_row_heights(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    row_indices: Optional[List[int]] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """平均列高（POST /ppt/distribute_table_row_heights）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "row_indices": row_indices,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/distribute_table_row_heights", json_body=body)


@mcp.tool()
async def delete_table(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """刪除整個表格（POST /ppt/delete_table）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_table", json_body=body)


@mcp.tool()
async def set_table_row_height(
    file_path: str,
    slide_index: int,
    row_idx: int,
    height_emu: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定列高（POST /ppt/set_table_row_height）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "height_emu": height_emu,
        "shape_id": shape_id,
        "shape_index": shape_index,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_table_row_height", json_body=body)


@mcp.tool()
async def set_table_column_width(
    file_path: str,
    slide_index: int,
    col_idx: int,
    width_emu: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定欄寬（POST /ppt/set_table_column_width）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "col_idx": col_idx,
        "width_emu": width_emu,
        "shape_id": shape_id,
        "shape_index": shape_index,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_table_column_width", json_body=body)


@mcp.tool()
async def rebuild_table_structure(
    file_path: str,
    slide_index: int,
    new_rows: int,
    new_cols: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    first_row_as_header: bool = False,
    font_size: int = 14,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """重建表格列欄數（POST /ppt/rebuild_table_structure）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "new_rows": new_rows,
        "new_cols": new_cols,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/rebuild_table_structure", json_body=body)


@mcp.tool()
async def insert_table_row(
    file_path: str,
    slide_index: int,
    insert_before: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    first_row_as_header: bool = False,
    font_size: int = 14,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """插入列（POST /ppt/insert_table_row）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "insert_before": insert_before,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/insert_table_row", json_body=body)


@mcp.tool()
async def delete_table_row(
    file_path: str,
    slide_index: int,
    row_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    first_row_as_header: bool = False,
    font_size: int = 14,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """刪除列（POST /ppt/delete_table_row）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "row_idx": row_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_table_row", json_body=body)


@mcp.tool()
async def insert_table_column(
    file_path: str,
    slide_index: int,
    insert_before: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    first_row_as_header: bool = False,
    font_size: int = 14,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """插入欄（POST /ppt/insert_table_column）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "insert_before": insert_before,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/insert_table_column", json_body=body)


@mcp.tool()
async def delete_table_column(
    file_path: str,
    slide_index: int,
    col_idx: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    first_row_as_header: bool = False,
    font_size: int = 14,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """刪除欄（POST /ppt/delete_table_column）。"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "col_idx": col_idx,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "first_row_as_header": first_row_as_header,
        "font_size": font_size,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_table_column", json_body=body)


@mcp.tool()
async def add_shape(
    file_path: str,
    slide_index: int,
    shape_type: str,
    left: int,
    top: int,
    width: int,
    height: int,
    text: str = "",
    fill_color: Optional[List[int]] = None,
    line_color: Optional[List[int]] = None,
    line_width: Optional[int] = None,
    font_size: int = 18,
    bold: bool = False,
    font_name: Optional[str] = None,
    font_color: Optional[List[int]] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """新增形狀圖案"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_type": shape_type,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "text": text,
        "fill_color": _clean_rgb(fill_color),
        "line_color": _clean_rgb(line_color),
        "line_width": line_width,
        "font_size": font_size,
        "bold": bold,
        "font_name": font_name,
        "font_color": _clean_rgb(font_color),
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_shape", json_body=body)


@mcp.tool()
async def add_line(
    file_path: str,
    slide_index: int,
    x1: int,
    y1: int,
    x2: int,
    y2: int,
    line_color: Optional[List[int]] = None,
    line_width: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """增加線段"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "x1": x1,
        "y1": y1,
        "x2": x2,
        "y2": y2,
        "line_color": _clean_rgb(line_color),
        "line_width": line_width,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_line", json_body=body)


@mcp.tool()
async def add_arrow(
    file_path: str,
    slide_index: int,
    left: int,
    top: int,
    width: int,
    height: int,
    direction: str = "right",
    text: str = "",
    fill_color: Optional[List[int]] = None,
    line_color: Optional[List[int]] = None,
    line_width: Optional[int] = None,
    font_size: int = 18,
    bold: bool = False,
    font_name: Optional[str] = None,
    font_color: Optional[List[int]] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """新增箭頭"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "direction": direction,
        "text": text,
        "fill_color": _clean_rgb(fill_color),
        "line_color": _clean_rgb(line_color),
        "line_width": line_width,
        "font_size": font_size,
        "bold": bold,
        "font_name": font_name,
        "font_color": _clean_rgb(font_color),
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_arrow", json_body=body)


@mcp.tool()
async def delete_slide(file_path: str, slide_index: int, save_as: Optional[str] = None) -> Dict[str, Any]:
    """刪除投影片"""
    body: Dict[str, Any] = {"file_path": file_path, "slide_index": slide_index}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_slide", json_body=body)


@mcp.tool()
async def duplicate_slide(file_path: str, slide_index: int, save_as: Optional[str] = None) -> Dict[str, Any]:
    """複製投影片"""
    body: Dict[str, Any] = {"file_path": file_path, "slide_index": slide_index}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/duplicate_slide", json_body=body)


@mcp.tool()
async def replace_text(
    file_path: str,
    old_text: str,
    new_text: str,
    slide_indices: Optional[List[int]] = None,
    exact_match: bool = False,
    case_sensitive: bool = True,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """替換文字"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "old_text": old_text,
        "new_text": new_text,
        "slide_indices": slide_indices,
        "exact_match": exact_match,
        "case_sensitive": case_sensitive,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/replace_text", json_body=body)


@mcp.tool()
async def add_bullets(
    file_path: str,
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
    font_color: Optional[List[int]] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """新增列表"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "items": items,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "font_size": font_size,
        "level": level,
        "bold": bold,
        "font_name": font_name,
        "font_color": _clean_rgb(font_color),
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_bullets", json_body=body)


@mcp.tool()
async def add_title_slide(file_path: str, title: str, subtitle: str = "", save_as: Optional[str] = None) -> Dict[str, Any]:
    """新增標題頁"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "title": title,
        "subtitle": subtitle,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_title_slide", json_body=body)


@mcp.tool()
async def reorder_slides(file_path: str, new_order: List[int], save_as: Optional[str] = None) -> Dict[str, Any]:
    """重新排序投影片"""
    body: Dict[str, Any] = {"file_path": file_path, "new_order": new_order}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/reorder_slides", json_body=body)


@mcp.tool()
async def set_slide_background_color(
    file_path: str,
    slide_index: int,
    rgb: List[int],
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定投影片背景顏色"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "rgb": _clean_rgb(rgb),
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_slide_background_color", json_body=body)


@mcp.tool()
async def set_slide_background_image(
    file_path: str,
    slide_index: int,
    image_path: str,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定投影片背景圖片"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "image_path": image_path,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_slide_background_image", json_body=body)


@mcp.tool()
async def set_slides_background_color(
    file_path: str,
    slide_indices: List[int],
    rgb: List[int],
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定多張投影片背景顏色"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_indices": slide_indices,
        "rgb": _clean_rgb(rgb),
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_slides_background_color", json_body=body)


@mcp.tool()
async def set_slides_background_image(
    file_path: str,
    slide_indices: List[int],
    image_path: str,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定多張投影片背景圖片"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_indices": slide_indices,
        "image_path": image_path,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/set_slides_background_image", json_body=body)


@mcp.tool()
async def ppt_theme_info(file_path: str) -> Dict[str, Any]:
    """
    獲取PPTX主題信息
    """
    return await _request("GET", "/ppt/theme_info", params={"file_path": file_path})


@mcp.tool()
async def ppt_slide_background(file_path: str, slide_index: int = 0) -> Dict[str, Any]:
    """
    獲取投影片背景
    """
    return await _request(
        "GET",
        "/ppt/slide_background",
        params={"file_path": file_path, "slide_index": slide_index},
    )


@mcp.tool()
async def ppt_slides_backgrounds(file_path: str) -> Dict[str, Any]:
    """
    獲取多張投影片背景
    """
    return await _request("GET", "/ppt/slides_backgrounds", params={"file_path": file_path})


@mcp.tool()
async def ppt_textbox_style(
    file_path: str,
    slide_index: int = 0,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
) -> Dict[str, Any]:
    """獲取文本框樣式"""
    return await _request(
        "GET",
        "/ppt/textbox_style",
        params={
            "file_path": file_path,
            "slide_index": slide_index,
            "shape_id": shape_id,
            "shape_index": shape_index,
        },
    )


@mcp.tool()
async def ppt_slide_textbox_styles(file_path: str, slide_index: int = 0) -> Dict[str, Any]:
    """獲取投影片文本框樣式"""
    return await _request(
        "GET",
        "/ppt/slide_textbox_styles",
        params={"file_path": file_path, "slide_index": slide_index},
    )


@mcp.tool()
async def set_textbox_style(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    fill_color: Optional[List[int]] = None,
    fill_transparency: Optional[float] = None,
    line_style: Optional[str] = None,
    line_color: Optional[List[int]] = None,
    line_width: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """設定文本框樣式"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "fill_color": _clean_rgb(fill_color) if fill_color is not None else None,
        "fill_transparency": fill_transparency,
        "line_style": line_style,
        "line_color": _clean_rgb(line_color) if line_color is not None else None,
        "line_width": line_width,
        "save_as": save_as,
    }
    return await _request("POST", "/ppt/set_textbox_style", json_body=body)


@mcp.tool()
async def drag_shape(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    left: Optional[int] = None,
    top: Optional[int] = None,
    delta_x: Optional[int] = None,
    delta_y: Optional[int] = None,
    width: Optional[int] = None,
    height: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """拖動形狀"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "left": left,
        "top": top,
        "delta_x": delta_x,
        "delta_y": delta_y,
        "width": width,
        "height": height,
        "save_as": save_as,
    }
    return await _request("POST", "/ppt/drag_shape", json_body=body)


@mcp.tool()
async def drag_textbox(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    left: Optional[int] = None,
    top: Optional[int] = None,
    delta_x: Optional[int] = None,
    delta_y: Optional[int] = None,
    width: Optional[int] = None,
    height: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """拖動文本框"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "left": left,
        "top": top,
        "delta_x": delta_x,
        "delta_y": delta_y,
        "width": width,
        "height": height,
        "save_as": save_as,
    }
    return await _request("POST", "/ppt/drag_textbox", json_body=body)


@mcp.tool()
async def delete_textbox(
    file_path: str,
    slide_index: int,
    shape_id: Optional[int] = None,
    shape_index: Optional[int] = None,
    save_as: Optional[str] = None,
) -> Dict[str, Any]:
    """刪除文本框"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
        "save_as": save_as,
    }
    return await _request("POST", "/ppt/delete_textbox", json_body=body)


@mcp.tool()
async def add_wordart_like_textbox(
        file_path: str,
        slide_index: int,
        text: str,
        left: int,
        top: int,
        width: int,
        height: int,
        font_size: int = 28,
        bold: bool = True,
        font_name: Optional[str] = None,
        font_color: Optional[List[int]] = None,
        align: str = "center",
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """新增WordArt文本框"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "text": text,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "font_size": font_size,
        "bold": bold,
        "font_name": font_name,
        "font_color": _clean_rgb(font_color) if font_color is not None else None,
        "align": align,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_wordart_like_textbox", json_body=body)


@mcp.tool()
async def update_wordart_text(
        file_path: str,
        slide_index: int,
        new_text: str,
        shape_name: Optional[str] = None,
        shape_id: Optional[int] = None,
        shape_index: Optional[int] = None,
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """更新WordArt文本框文字"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "new_text": new_text,
        "shape_name": shape_name,
        "shape_id": shape_id,
        "shape_index": shape_index,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/update_wordart_text", json_body=body)


@mcp.tool()
async def delete_shape(
        file_path: str,
        slide_index: int,
        shape_id: Optional[int] = None,
        shape_index: Optional[int] = None,
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """刪除形狀"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_id": shape_id,
        "shape_index": shape_index,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_shape", json_body=body)


@mcp.tool()
async def clone_named_shape_from_template(
        file_path: str,
        slide_index: int,
        shape_name: str,
        new_text: str = "",
        left: Optional[int] = None,
        top: Optional[int] = None,
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """從模板複製形狀"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "shape_name": shape_name,
        "new_text": new_text,
        "left": left,
        "top": top,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/clone_named_shape_from_template", json_body=body)


@mcp.tool()
async def parse_math_expression(input_text: str, input_type: str = "latex") -> Dict[str, Any]:
    """解析數學公式"""
    return await _request(
        "POST",
        "/ppt/parse_math_expression",
        json_body={"input_text": input_text, "input_type": input_type},
    )


@mcp.tool()
async def add_equation(
        file_path: str,
        slide_index: int,
        input_text: str,
        input_type: str = "latex",
        left: int = 0,
        top: int = 0,
        width: Optional[int] = None,
        height: Optional[int] = None,
        font_size: Optional[int] = None,
        color: Optional[List[int]] = None,
        prefix_runs: Optional[List[Dict[str, Any]]] = None,
        suffix_runs: Optional[List[Dict[str, Any]]] = None,
        render_mode: str = "omml",
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """新增數學公式"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "slide_index": slide_index,
        "input_text": input_text,
        "input_type": input_type,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "font_size": font_size,
        "color": _clean_rgb(color) if color is not None else None,
        "prefix_runs": prefix_runs,
        "suffix_runs": suffix_runs,
        "render_mode": render_mode,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_equation", json_body=body)


@mcp.tool()
async def update_equation(
        file_path: str,
        input_text: str,
        input_type: str = "latex",
        expr_id: Optional[str] = None,
        shape_id: Optional[int] = None,
        slide_index: Optional[int] = None,
        prefix_runs: Optional[List[Dict[str, Any]]] = None,
        suffix_runs: Optional[List[Dict[str, Any]]] = None,
        render_mode: str = "omml",
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """更新數學公式"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "input_text": input_text,
        "input_type": input_type,
        "expr_id": expr_id,
        "shape_id": shape_id,
        "slide_index": slide_index,
        "prefix_runs": prefix_runs,
        "suffix_runs": suffix_runs,
        "render_mode": render_mode,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/update_equation", json_body=body)


@mcp.tool()
async def delete_equation(
        file_path: str,
        expr_id: Optional[str] = None,
        shape_id: Optional[int] = None,
        slide_index: Optional[int] = None,
        render_mode: str = "omml",
        save_as: Optional[str] = None,
    ) -> Dict[str, Any]:
    """刪除數學公式"""
    body: Dict[str, Any] = {
        "file_path": file_path,
        "expr_id": expr_id,
        "shape_id": shape_id,
        "slide_index": slide_index,
        "render_mode": render_mode,
    }
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_equation", json_body=body)


@mcp.tool()
async def render_slide_to_image(
    file_path: str,
    slide_index: int,
    output_path: str,
    dpi: int = 150,
    libreoffice_path: Optional[str] = None,
) -> Dict[str, Any]:
    """渲染投影片為圖片"""
    return await _request(
        "POST",
        "/ppt/render_slide_to_image",
        json_body={
            "file_path": file_path,
            "slide_index": slide_index,
            "output_path": output_path,
            "dpi": dpi,
            "libreoffice_path": libreoffice_path,
        },
    )


@mcp.tool()
async def render_slides_to_grid_image(
    file_path: str,
    slide_indices: List[int],
    output_path: str,
    cols: int = 2,
    dpi: int = 150,
    libreoffice_path: Optional[str] = None,
    add_page_title: bool = True,
    figure_title: Optional[str] = None,
) -> Dict[str, Any]:
    """渲染多張投影片為圖片"""
    return await _request(
        "POST",
        "/ppt/render_slides_to_grid_image",
        json_body={
            "file_path": file_path,
            "slide_indices": slide_indices,
            "output_path": output_path,
            "cols": cols,
            "dpi": dpi,
            "libreoffice_path": libreoffice_path,
            "add_page_title": add_page_title,
            "figure_title": figure_title,
        },
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="MCP server for PPT FastAPI backend")
    parser.add_argument("--api-base", default=DEFAULT_API_BASE, help="PPT API base URL, default: http://10.1.3.127:6414")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, help="HTTP timeout seconds")
    parser.add_argument(
        "--transport",
        default="stdio",
        choices=["stdio", "streamable-http", "sse"],
        help="MCP transport",
    )
    parser.add_argument("--host", default="10.1.3.127", help="streamable-http / sse host, default: 10.1.3.127")
    parser.add_argument("--port", type=int, default=6414, help="streamable-http / sse port, default: 6414")
    parser.add_argument("--path", default="/mcp", help="streamable-http path, default: /mcp")
    args = parser.parse_args()

    set_runtime_config(args.api_base, args.timeout)

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "streamable-http":
        mcp.run(transport="streamable-http", host=args.host, port=args.port, path=args.path)
    else:
        mcp.run(transport="sse", host=args.host, port=args.port, path=args.path)


if __name__ == "__main__":
    main()

