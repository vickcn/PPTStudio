#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EMU_PER_INCH = 914400


if __package__ in {None, ""}:
    from ppt_stdio import (
        get_presentation_theme_info,
        open_presentation,
        scan_presentation_backgrounds,
        scan_presentation_text_fonts,
    )
else:
    from .ppt_stdio import (
        get_presentation_theme_info,
        open_presentation,
        scan_presentation_backgrounds,
        scan_presentation_text_fonts,
    )


def emu_to_inches(value: int) -> float:
    return round(int(value) / EMU_PER_INCH, 4)


def qn(prefix: str, tag: str) -> str:
    return "{%s}%s" % (NS[prefix], tag)


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def font_size_pt_from_rpr(rpr) -> float | None:
    if rpr is None or not rpr.get("sz"):
        return None
    return round(int(rpr.get("sz")) / 100, 2)


def font_typefaces_from_rpr(rpr) -> dict:
    if rpr is None:
        return {}

    result = {}
    latin = rpr.find("a:latin", NS)
    ea = rpr.find("a:ea", NS)
    cs = rpr.find("a:cs", NS)
    if latin is not None and latin.get("typeface"):
        result["latin_font_name"] = latin.get("typeface")
    if ea is not None and ea.get("typeface"):
        result["east_asian_font_name"] = ea.get("typeface")
    if cs is not None and cs.get("typeface"):
        result["complex_script_font_name"] = cs.get("typeface")
    if "latin_font_name" in result:
        result["font_name"] = result["latin_font_name"]
    elif "east_asian_font_name" in result:
        result["font_name"] = result["east_asian_font_name"]
    return result


def api_slide_to_json_slide(api_slide_index: int) -> int:
    return int(api_slide_index) + 1


def normalize_api_slide_payload(payload):
    if isinstance(payload, dict):
        normalized = {}
        for key, value in payload.items():
            if key == "slide_index" and value is not None:
                normalized[key] = api_slide_to_json_slide(value)
            else:
                normalized[key] = normalize_api_slide_payload(value)
        return normalized
    if isinstance(payload, list):
        return [normalize_api_slide_payload(item) for item in payload]
    return payload


def maybe_fix_mojibake(value):
    if isinstance(value, dict):
        return {k: maybe_fix_mojibake(v) for k, v in value.items()}
    if isinstance(value, list):
        return [maybe_fix_mojibake(v) for v in value]
    if not isinstance(value, str) or not value:
        return value

    bad_markers = (
        "Ã", "æ", "ç", "é", "å", "ï", "ð", "¾", "¼", "¢", "€",
        "\x81", "\x82", "\x83", "\x84", "\x85", "\x86", "\x87", "\x88", "\x89",
        "\x8a", "\x8b", "\x8c", "\x8d", "\x8e", "\x8f", "\x90", "\x91", "\x92",
        "\x93", "\x94", "\x95", "\x96", "\x97", "\x98", "\x99", "\x9a", "\x9b",
        "\x9c", "\x9d", "\x9e", "\x9f",
    )
    current = value
    for _ in range(3):
        if not any(ch in current for ch in bad_markers):
            break
        try:
            fixed = current.encode("latin1").decode("utf-8")
        except Exception:
            break
        if fixed == current:
            break
        current = fixed
    return current


def read_rels(zf: zipfile.ZipFile, rels_path: str) -> dict:
    rels = {}
    if rels_path not in zf.namelist():
        return rels
    root = ET.fromstring(zf.read(rels_path))
    for rel in root:
        rid = rel.get("Id")
        rels[rid] = {
            "type": rel.get("Type"),
            "target": rel.get("Target"),
        }
    return rels


def slide_num_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def parse_xfrm(node) -> dict | None:
    if node is None:
        return None
    off = node.find("a:off", NS)
    ext = node.find("a:ext", NS)
    if off is None or ext is None:
        return None
    x = int(off.get("x", 0))
    y = int(off.get("y", 0))
    cx = int(ext.get("cx", 0))
    cy = int(ext.get("cy", 0))
    return {
        "left_emu": x,
        "top_emu": y,
        "width_emu": cx,
        "height_emu": cy,
        "left_in": emu_to_inches(x),
        "top_in": emu_to_inches(y),
        "width_in": emu_to_inches(cx),
        "height_in": emu_to_inches(cy),
    }


def parse_rotation_deg(node) -> float | None:
    if node is None:
        return None
    raw = node.get("rot")
    if raw is None:
        return None
    try:
        return round(int(raw) / 60000.0, 4)
    except Exception:
        return None


def parse_crop(blip_fill) -> dict[str, float] | None:
    if blip_fill is None:
        return None
    src_rect = blip_fill.find("a:srcRect", NS)
    if src_rect is None:
        return None

    crop: dict[str, float] = {}
    for xml_key, out_key in (("l", "left"), ("r", "right"), ("t", "top"), ("b", "bottom")):
        raw = src_rect.get(xml_key)
        if raw is None:
            continue
        try:
            crop[out_key] = round(int(raw) / 100000.0, 6)
        except Exception:
            continue
    return crop or None


def parse_text_frame_layout(tx_body) -> dict[str, Any]:
    if tx_body is None:
        return {}

    body_pr = tx_body.find("a:bodyPr", NS)
    if body_pr is None:
        return {}

    layout: dict[str, Any] = {}

    wrap = body_pr.get("wrap")
    if wrap is not None:
        layout["word_wrap"] = wrap.lower() != "none"

    if body_pr.find("a:noAutofit", NS) is not None:
        layout["auto_fit"] = "none"
    elif body_pr.find("a:normAutofit", NS) is not None:
        layout["auto_fit"] = "shrink_text"
    elif body_pr.find("a:spAutoFit", NS) is not None:
        layout["auto_fit"] = "shape_to_fit_text"

    return layout


def resolve_zip_target(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    normalized = posixpath.normpath(posixpath.join(base_dir, target))
    return normalized.lstrip("/")


def export_zip_asset(
    zf: zipfile.ZipFile,
    zip_path: str | None,
    export_dir: Path | None,
    output_name: str,
) -> str | None:
    if not export_dir or not zip_path or zip_path not in zf.namelist():
        return None
    export_dir.mkdir(parents=True, exist_ok=True)
    out_file = export_dir / output_name
    if not out_file.exists():
        out_file.write_bytes(zf.read(zip_path))
    return str(out_file)


def make_export_name(prefix: str, zip_path: str | None, default_suffix: str = ".png") -> str:
    if zip_path:
        source = Path(zip_path)
        return f"{prefix}_{source.stem}{source.suffix or default_suffix}"
    return f"{prefix}{default_suffix}"


def parse_bg_element(
    bg_node,
    rels: dict,
    base_part: str,
    zf: zipfile.ZipFile,
    export_dir: Path | None,
    export_name: str,
) -> dict:
    if bg_node is None:
        return {"mode": "inherit"}

    info = {"mode": "explicit"}
    bg_pr = bg_node.find("p:bgPr", NS)
    if bg_pr is None:
        info["xml"] = ET.tostring(bg_node, encoding="unicode")
        return info

    solid = bg_pr.find("a:solidFill", NS)
    if solid is not None:
        srgb = solid.find("a:srgbClr", NS)
        scheme = solid.find("a:schemeClr", NS)
        if srgb is not None:
            rgb = srgb.get("val")
            info.update({"type": "solid", "color_rgb": [int(rgb[i:i + 2], 16) for i in (0, 2, 4)]})
        elif scheme is not None:
            info.update({"type": "scheme", "scheme_color": scheme.get("val")})

    blip_fill = bg_pr.find("a:blipFill", NS)
    if blip_fill is not None:
        blip = blip_fill.find("a:blip", NS)
        rid = blip.get(qn("r", "embed")) if blip is not None else None
        image_target = rels.get(rid, {}).get("target") if rid else None
        image_zip_path = resolve_zip_target(base_part, image_target) if image_target else None
        export_path = export_zip_asset(zf, image_zip_path, export_dir, export_name)
        info.update(
            {
                "type": "image",
                "image_rid": rid,
                "image_target": image_target,
                "image_zip_path": image_zip_path,
                "image_export_path": export_path,
            }
        )

    info["xml"] = ET.tostring(bg_node, encoding="unicode")
    return info


def get_text_runs(tx_body) -> tuple[str, list[dict]]:
    paragraphs = []
    full_text = []
    lst_style = tx_body.find("a:lstStyle", NS)
    for p_idx, p in enumerate(tx_body.findall("a:p", NS)):
        runs = []
        para_text = []
        run_index = 0
        ppr = p.find("a:pPr", NS)
        level = int(ppr.get("lvl", "0")) if ppr is not None and ppr.get("lvl") else 0
        lvl_tag = f"a:lvl{level + 1}pPr"
        lvl_ppr = lst_style.find(lvl_tag, NS) if lst_style is not None else None
        default_font_size_pt = None
        for candidate in (
            p.find("a:endParaRPr", NS),
            ppr.find("a:defRPr", NS) if ppr is not None else None,
            lvl_ppr.find("a:defRPr", NS) if lvl_ppr is not None else None,
        ):
            default_font_size_pt = font_size_pt_from_rpr(candidate)
            if default_font_size_pt is not None:
                break
        for r in list(p):
            lname = local_name(r.tag)
            if lname == "r":
                text = "".join(t.text or "" for t in r.findall("a:t", NS))
                para_text.append(text)
                rpr = r.find("a:rPr", NS)
                run_info = {
                    "run_index": run_index,
                    "text": text,
                    "lang": rpr.get("lang") if rpr is not None else None,
                    "font_size_pt": font_size_pt_from_rpr(rpr),
                    "effective_font_size_pt": font_size_pt_from_rpr(rpr) or default_font_size_pt,
                    "bold": rpr.get("b") == "1" if rpr is not None and rpr.get("b") is not None else None,
                    "italic": rpr.get("i") == "1" if rpr is not None and rpr.get("i") is not None else None,
                }
                run_info.update(font_typefaces_from_rpr(rpr))
                runs.append(run_info)
                run_index += 1
            elif lname == "br":
                para_text.append("\n")
            elif lname == "fld":
                text = "".join(t.text or "" for t in r.findall("a:t", NS))
                para_text.append(text)
                runs.append({"run_index": run_index, "text": text, "field": True})
                run_index += 1
        p_text = "".join(para_text)
        full_text.append(p_text)
        paragraphs.append({"paragraph_index": p_idx, "text": p_text, "runs": runs})
    return "\n".join(full_text), paragraphs


def parse_slide_master_background(zf: zipfile.ZipFile, export_media_dir: Path | None) -> dict:
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    pres_rels = read_rels(zf, "ppt/_rels/presentation.xml.rels")
    master_id = pres.find("p:sldMasterIdLst/p:sldMasterId", NS)
    if master_id is None:
        return {"mode": "unknown"}

    master_rid = master_id.get(qn("r", "id"))
    master_target = pres_rels.get(master_rid, {}).get("target")
    if not master_target:
        return {"mode": "unknown"}

    master_part = resolve_zip_target("ppt/presentation.xml", master_target)
    master_root = ET.fromstring(zf.read(master_part))
    master_rels = read_rels(
        zf,
        posixpath.join(posixpath.dirname(master_part), "_rels", Path(master_part).name + ".rels"),
    )
    bg_node = master_root.find("p:cSld/p:bg", NS)
    blip = master_root.find("p:cSld/p:bg/p:bgPr/a:blipFill/a:blip", NS)
    rid = blip.get(qn("r", "embed")) if blip is not None else None
    target = master_rels.get(rid, {}).get("target") if rid else None
    zip_path = resolve_zip_target(master_part, target) if target else None
    return parse_bg_element(
        bg_node,
        master_rels,
        master_part,
        zf,
        export_media_dir,
        make_export_name("slide_master_background", zip_path),
    )


def parse_layout_from_pptx(ppt_path: Path, export_media_dir: Path | None = None) -> dict:
    with zipfile.ZipFile(ppt_path) as zf:
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        sld_sz = pres.find("p:sldSz", NS)
        slide_width = int(sld_sz.get("cx"))
        slide_height = int(sld_sz.get("cy"))
        slide_master_background = parse_slide_master_background(zf, export_media_dir)

        slide_files = sorted(
            [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=slide_num_key,
        )

        exported_media: list[dict] = []
        result = {
            "file_path": str(ppt_path),
            "slide_count": len(slide_files),
            "slide_size": {
                "width_emu": slide_width,
                "height_emu": slide_height,
                "width_in": emu_to_inches(slide_width),
                "height_in": emu_to_inches(slide_height),
            },
            "slide_master_background": slide_master_background,
            "exported_media_dir": str(export_media_dir) if export_media_dir else None,
            "exported_media": exported_media,
            "slides": [],
        }

        for idx, slide_path in enumerate(slide_files):
            slide_root = ET.fromstring(zf.read(slide_path))
            rels = read_rels(zf, "ppt/slides/_rels/" + Path(slide_path).name + ".rels")
            bg_info = parse_bg_element(
                slide_root.find("p:cSld/p:bg", NS),
                rels,
                slide_path,
                zf,
                export_media_dir,
                make_export_name(f"slide_{idx:02d}_background", None),
            )

            sp_tree = slide_root.find("p:cSld/p:spTree", NS)
            shapes = []
            shape_index = 0

            for child in list(sp_tree):
                lname = local_name(child.tag)
                if lname in {"nvGrpSpPr", "grpSpPr"}:
                    continue

                entry = {"shape_index": shape_index, "xml_tag": lname}
                shape_index += 1

                if lname == "sp":
                    c_nv_pr = child.find("p:nvSpPr/p:cNvPr", NS)
                    xfrm = child.find("p:spPr/a:xfrm", NS)
                    tx_body = child.find("p:txBody", NS)
                    entry.update(
                        {
                            "shape_id": int(c_nv_pr.get("id")) if c_nv_pr is not None else None,
                            "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                            "shape_type": "text_or_auto_shape",
                            "geometry": parse_xfrm(xfrm),
                        }
                    )
                    prst = child.find("p:spPr/a:prstGeom", NS)
                    if prst is not None:
                        entry["preset_geometry"] = prst.get("prst")
                    if tx_body is not None:
                        text_frame_layout = parse_text_frame_layout(tx_body)
                        if text_frame_layout:
                            entry["text_frame"] = text_frame_layout
                        text, paragraphs = get_text_runs(tx_body)
                        entry["text"] = text
                        entry["paragraphs"] = paragraphs
                    shapes.append(entry)
                    continue

                if lname == "pic":
                    c_nv_pr = child.find("p:nvPicPr/p:cNvPr", NS)
                    xfrm = child.find("p:spPr/a:xfrm", NS)
                    blip_fill = child.find("p:blipFill", NS)
                    blip = child.find("p:blipFill/a:blip", NS)
                    rid = blip.get(qn("r", "embed")) if blip is not None else None
                    image_target = rels.get(rid, {}).get("target") if rid else None
                    image_zip_path = resolve_zip_target(slide_path, image_target) if image_target else None
                    export_path = export_zip_asset(
                        zf,
                        image_zip_path,
                        export_media_dir,
                        make_export_name(
                            f"slide_{idx:02d}_shape_{int(c_nv_pr.get('id')) if c_nv_pr is not None else shape_index}",
                            image_zip_path,
                        ),
                    )
                    if export_path:
                        exported_media.append(
                            {
                                "slide_index": api_slide_to_json_slide(idx),
                                "shape_id": int(c_nv_pr.get("id")) if c_nv_pr is not None else None,
                                "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                                "zip_path": image_zip_path,
                                "export_path": export_path,
                            }
                        )
                    entry.update(
                        {
                            "shape_id": int(c_nv_pr.get("id")) if c_nv_pr is not None else None,
                            "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                            "shape_type": "picture",
                            "geometry": parse_xfrm(xfrm),
                            "rotation_deg": parse_rotation_deg(xfrm),
                            "crop": parse_crop(blip_fill),
                            "embed_rid": rid,
                            "image_target": image_target,
                            "image_zip_path": image_zip_path,
                            "export_image_path": export_path,
                            "image_rel_type": rels.get(rid, {}).get("type"),
                        }
                    )
                    shapes.append(entry)
                    continue

                if lname == "graphicFrame":
                    c_nv_pr = child.find("p:nvGraphicFramePr/p:cNvPr", NS)
                    xfrm = child.find("p:xfrm", NS)
                    entry.update(
                        {
                            "shape_id": int(c_nv_pr.get("id")) if c_nv_pr is not None else None,
                            "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                            "shape_type": "graphic_frame",
                            "geometry": parse_xfrm(xfrm),
                        }
                    )
                    table = child.find("a:graphic/a:graphicData/a:tbl", NS)
                    if table is not None:
                        entry["graphic_kind"] = "table"
                    shapes.append(entry)
                    continue

                if lname == "grpSp":
                    c_nv_pr = child.find("p:nvGrpSpPr/p:cNvPr", NS)
                    xfrm = child.find("p:grpSpPr/a:xfrm", NS)
                    entry.update(
                        {
                            "shape_id": int(c_nv_pr.get("id")) if c_nv_pr is not None else None,
                            "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                            "shape_type": "group",
                            "geometry": parse_xfrm(xfrm),
                        }
                    )
                    shapes.append(entry)
                    continue

                shapes.append(entry)

            result["slides"].append(
                {
                    "slide_index": api_slide_to_json_slide(idx),
                    "slide_number": api_slide_to_json_slide(idx),
                    "background": bg_info,
                    "shape_count": len(shapes),
                    "shapes": shapes,
                }
            )

    return maybe_fix_mojibake(result)


def merge_api_data(layout: dict, ppt_path: Path) -> dict:
    doc = open_presentation(str(ppt_path))
    fonts = maybe_fix_mojibake(scan_presentation_text_fonts(doc))
    backgrounds = maybe_fix_mojibake(scan_presentation_backgrounds(doc))
    theme = maybe_fix_mojibake(get_presentation_theme_info(doc))

    font_data = fonts.get("data") if isinstance(fonts, dict) and "data" in fonts else fonts
    if font_data and isinstance(font_data, dict):
        font_slides = {
            api_slide_to_json_slide(s["slide_index"]): normalize_api_slide_payload(s)
            for s in font_data.get("slides", [])
        }
        layout["font_detection_summary"] = {
            "detected_font_count": font_data.get("detected_font_count"),
            "unresolved_run_count": font_data.get("unresolved_run_count"),
            "font_summary": font_data.get("font_summary", []),
        }

        for slide in layout.get("slides", []):
            slide_index = slide["slide_index"]
            fslide = font_slides.get(slide_index, {})
            slide["font_summary"] = fslide.get("font_summary", [])
            slide["detected_font_count"] = fslide.get("detected_font_count")

            font_shapes = [s for s in fslide.get("shapes", []) if s.get("kind") == "text_frame"]
            used = set()
            for shape in slide.get("shapes", []):
                if "text" not in shape:
                    continue

                match = None
                for idx, fs in enumerate(font_shapes):
                    if idx in used:
                        continue
                    if fs.get("shape_id") == shape.get("shape_id") and fs.get("name") == shape.get("name"):
                        match = (idx, fs)
                        break
                if match is None:
                    for idx, fs in enumerate(font_shapes):
                        if idx in used:
                            continue
                        if fs.get("shape_id") == shape.get("shape_id"):
                            match = (idx, fs)
                            break
                if match is None:
                    for idx, fs in enumerate(font_shapes):
                        if idx in used:
                            continue
                        if fs.get("name") == shape.get("name"):
                            match = (idx, fs)
                            break
                if match is None:
                    continue

                idx, fs = match
                used.add(idx)
                shape["font_detail"] = {"text_preview": shape.get("text", ""), "paragraphs": []}

                fs_paragraphs = fs.get("paragraphs", [])
                base_paragraphs = shape.get("paragraphs", [])
                for p_idx, base_paragraph in enumerate(base_paragraphs):
                    src_paragraph = fs_paragraphs[p_idx] if p_idx < len(fs_paragraphs) else {}
                    merged_paragraph = {
                        "paragraph_index": base_paragraph.get("paragraph_index", p_idx),
                        "text": base_paragraph.get("text", ""),
                        "runs": [],
                    }
                    src_runs = src_paragraph.get("runs", [])
                    base_runs = base_paragraph.get("runs", [])
                    for r_idx, src_run in enumerate(src_runs):
                        merged_run = dict(src_run)
                        if r_idx < len(base_runs):
                            merged_run["text"] = base_runs[r_idx].get("text", "")
                            merged_run["effective_font_size_pt"] = (
                                merged_run.get("font_size_pt")
                                if merged_run.get("font_size_pt") is not None
                                else base_runs[r_idx].get("effective_font_size_pt")
                            )
                            for key in (
                                "font_name",
                                "latin_font_name",
                                "east_asian_font_name",
                                "complex_script_font_name",
                            ):
                                if merged_run.get(key) is None and base_runs[r_idx].get(key) is not None:
                                    merged_run[key] = base_runs[r_idx].get(key)
                        merged_paragraph["runs"].append(merged_run)
                    if not merged_paragraph["runs"] and base_paragraph.get("text"):
                        merged_paragraph["runs"].append({"run_index": 0, "text": base_paragraph.get("text", "")})
                    shape["font_detail"]["paragraphs"].append(merged_paragraph)

                font_names = []
                font_sizes = []
                for paragraph in shape["font_detail"]["paragraphs"]:
                    for run in paragraph.get("runs", []):
                        if run.get("font_name"):
                            font_names.append(run["font_name"])
                        if run.get("font_size_pt") is not None:
                            font_sizes.append(run["font_size_pt"])
                shape["font_summary"] = {
                    "font_names": sorted(set(font_names)),
                    "font_sizes_pt": sorted(set(font_sizes)),
                }

    backgrounds_data = backgrounds.get("data") if isinstance(backgrounds, dict) and "data" in backgrounds else backgrounds
    theme_data = theme.get("data") if isinstance(theme, dict) and "data" in theme else theme
    bg_slides = {
        api_slide_to_json_slide(s["slide_index"]): normalize_api_slide_payload(s)
        for s in (backgrounds_data or {}).get("slides", [])
    }
    theme_slides = {
        api_slide_to_json_slide(s["slide_index"]): normalize_api_slide_payload(s)
        for s in (theme_data or {}).get("slides", [])
    }
    layout["theme_info"] = (theme_data or {}).get("theme_info", {})

    for slide in layout.get("slides", []):
        slide_index = slide["slide_index"]
        slide["background_api"] = bg_slides.get(slide_index, {})
        slide["background_theme"] = theme_slides.get(slide_index, {})

    return layout


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Parse PPTX structure and export layout JSON.")
    parser.add_argument("file_path", help="Target .pptx file path")
    parser.add_argument("-o", "--output", help="Output JSON path; default is sibling .json", default=None)
    parser.add_argument(
        "--export-media-dir",
        help="Directory to export picture assets; default is <json_stem>_assets beside output JSON",
    )
    parser.add_argument("--no-api", action="store_true", help="Skip font/background/theme merge")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    ppt_path = Path(args.file_path).resolve()
    if not ppt_path.exists():
        raise SystemExit(f"File not found: {ppt_path}")
    if ppt_path.suffix.lower() != ".pptx":
        raise SystemExit(f"Only .pptx is supported: {ppt_path}")

    output_path = Path(args.output).resolve() if args.output else ppt_path.with_suffix(".json")
    export_media_dir = (
        Path(args.export_media_dir).resolve()
        if args.export_media_dir
        else output_path.with_name(f"{output_path.stem}_assets")
    )

    layout = parse_layout_from_pptx(ppt_path, export_media_dir=export_media_dir)
    if not args.no_api:
        layout = merge_api_data(layout, ppt_path)

    output_path.write_text(json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
