# -*- coding: utf-8 -*-
"""
mcp_server.py

把既有的 FastAPI PPT API 包成 MCP server。
預設以 stdio transport 執行，適合 Cursor / Claude Desktop / 自建 MCP Gateway。
也可切成 streamable-http transport。

需求：
    pip install "mcp[cli]" httpx

範例：
    python mcp_server.py
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport stdio
    python mcp_server.py --api-base http://10.1.3.127:6414 --transport streamable-http --host 10.1.3.127 --port 6414

注意：
1. 你的 api_server.py 需先啟動，例如：
       uvicorn api_server:app --host 10.1.3.127 --port 6414
2. 目前 api_server.py 內的 /ppt/add_shape 端點會依賴 add_shape，
   但原檔 import 區沒有匯入 add_shape；若未補上，該工具呼叫時會失敗。
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
        raise ValueError("rgb / color 必須是 3 個整數 [R, G, B]")
    vals = [int(v) for v in rgb]
    for v in vals:
        if v < 0 or v > 255:
            raise ValueError("rgb / color 每個值都必須介於 0~255")
    return vals


@mcp.tool()
async def health() -> Dict[str, Any]:
    """檢查底層 PPT API server 是否存活。"""
    return await _request("GET", "/health")


@mcp.tool()
async def root_info() -> Dict[str, Any]:
    """取得底層 PPT API server 的基本資訊。"""
    return await _request("GET", "/")


@mcp.tool()
async def new_presentation(
    file_path: str,
    plank_page_num: int = 1,
    plank_page_width: int = 1080,
    plank_page_height: int = 1920,
    dpi: int = 96,
) -> Dict[str, Any]:
    """建立新的 PPTX 檔案。"""
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
    """讀取 PPT 基本資訊。"""
    return await _request("GET", "/ppt/info", params={"file_path": file_path})


@mcp.tool()
async def list_presentation_slides(file_path: str) -> Dict[str, Any]:
    """列出所有投影片與 shape 摘要。"""
    return await _request("GET", "/ppt/slides", params={"file_path": file_path})


@mcp.tool()
async def save_as(file_path: str, save_as: str) -> Dict[str, Any]:
    """另存 PPT 到新路徑。"""
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
    """新增一張空白投影片。"""
    body: Dict[str, Any] = {"file_path": file_path, "page_num": 1}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/add_blank_slide", json_body=body)


@mcp.tool()
async def add_blank_slides(file_path: str, page_num: int = 1, save_as: Optional[str] = None) -> Dict[str, Any]:
    """新增多張空白投影片。"""
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
    """在指定投影片加入文字方塊。"""
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
    """在指定投影片加入圖片。"""
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
    """在指定投影片加入表格。"""
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
    """在指定投影片加入一般圖形，例如 rect / ellipse / diamond。"""
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
    """在指定投影片加入直線。"""
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
    """在指定投影片加入箭頭圖形。"""
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
    """刪除指定投影片。"""
    body: Dict[str, Any] = {"file_path": file_path, "slide_index": slide_index}
    if save_as:
        body["save_as"] = save_as
    return await _request("POST", "/ppt/delete_slide", json_body=body)


@mcp.tool()
async def duplicate_slide(file_path: str, slide_index: int, save_as: Optional[str] = None) -> Dict[str, Any]:
    """複製指定投影片。"""
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
    """在整份或指定投影片中取代文字。"""
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
    """在指定投影片加入項目符號清單。"""
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
    """新增標題頁。"""
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
    """重新排列投影片順序。"""
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
    """設定單頁背景顏色。"""
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
    """設定單頁背景圖片。"""
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
    """設定多頁背景顏色。"""
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
    """設定多頁背景圖片。"""
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
    呼叫 GET /ppt/theme_info，讀取 PPTX 佈景主題（theme）、色彩配置、字型配置等摘要。
    需先啟動 api_server；底層邏輯由 ppt_stdio.get_presentation_theme_info 提供。
    """
    return await _request("GET", "/ppt/theme_info", params={"file_path": file_path})


@mcp.tool()
async def ppt_slide_background(file_path: str, slide_index: int = 0) -> Dict[str, Any]:
    """
    呼叫 GET /ppt/slide_background，讀取指定頁之背景類型（繼承母片、純色、圖片、滿版圖片模擬等）。
    """
    return await _request(
        "GET",
        "/ppt/slide_background",
        params={"file_path": file_path, "slide_index": slide_index},
    )


@mcp.tool()
async def ppt_slides_backgrounds(file_path: str) -> Dict[str, Any]:
    """
    呼叫 GET /ppt/slides_backgrounds，一次掃描整份簡報各頁背景並附帶 theme 摘要。
    """
    return await _request("GET", "/ppt/slides_backgrounds", params={"file_path": file_path})


@mcp.tool()
async def render_slide_to_image(
    file_path: str,
    slide_index: int,
    output_path: str,
    dpi: int = 150,
    libreoffice_path: Optional[str] = None,
) -> Dict[str, Any]:
    """把指定投影片輸出成單張圖片。"""
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
    """把多頁投影片輸出成 grid 拼圖。"""
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
    parser.add_argument("--api-base", default=DEFAULT_API_BASE, help="底層 PPT API base URL，例如 http://10.1.3.127:6414")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, help="HTTP timeout seconds")
    parser.add_argument(
        "--transport",
        default="stdio",
        choices=["stdio", "streamable-http", "sse"],
        help="MCP transport",
    )
    parser.add_argument("--host", default="10.1.3.127", help="streamable-http / sse 綁定 host")
    parser.add_argument("--port", type=int, default=6414, help="streamable-http / sse 綁定 port")
    parser.add_argument("--path", default="/mcp", help="streamable-http 路徑")
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
