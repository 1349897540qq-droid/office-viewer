#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Antigravity MCP App Server: Office Viewer (.docx / .xlsx / .pptx)
100% pure Python 3 standard library — zero external dependencies.
Implements MCP stdio JSON-RPC 2.0 + SEP-1865 MCP Apps (ui:// resource + aux_pane entrypoint).
"""

import argparse
import base64
import datetime
import html as html_lib
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
import xml.etree.ElementTree as ET
import zipfile

APP_URI = "ui://office-viewer/app.html"
SUPPORTED_EXTS = {".docx", ".xlsx", ".pptx"}
MAX_TRAJECTORY_BYTES = 420_000   # Keep preview_office_file structuredContent safely < 1 MiB
MAX_IMAGE_BYTES = 800_000        # Per embedded image cap
MAX_TOTAL_IMAGE_BYTES = 3_500_000
MAX_XLSX_ROWS = 500
MAX_XLSX_COLS = 60

# XML Namespaces
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

MIME_MAP = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
    ".bmp": "image/bmp",
}


def qn(prefix, tag):
    return "{%s}%s" % (NS[prefix], tag)


def safe_color(val):
    if not val:
        return None
    val = val.strip().lstrip("#")
    if val.lower() in ("auto", "none", ""):
        return None
    if re.fullmatch(r"[0-9a-fA-F]{6}", val):
        return "#" + val.upper()
    if re.fullmatch(r"[0-9a-fA-F]{8}", val):
        # ARGB (Excel) -> RGB
        return "#" + val[2:].upper()
    return None


def human_size(num_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < 1024.0 or unit == "GB":
            return "%.1f %s" % (num_bytes, unit) if unit != "B" else "%d B" % num_bytes
        num_bytes /= 1024.0
    return "%d B" % num_bytes


def format_mtime(ts):
    return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")


def parse_rels(zf, rels_path):
    rels = {}
    if rels_path not in zf.namelist():
        return rels
    try:
        root = ET.fromstring(zf.read(rels_path))
        for rel in root:
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            mode = rel.attrib.get("TargetMode", "")
            rtype = rel.attrib.get("Type", "")
            if rid:
                rels[rid] = {"target": target, "mode": mode, "type": rtype}
    except Exception:
        pass
    return rels


def resolve_zip_path(base_dir, target):
    if not target:
        return ""
    if target.startswith("/"):
        return target.lstrip("/")
    parts = [p for p in (base_dir.split("/") if base_dir else []) if p]
    for seg in target.split("/"):
        if seg == "" or seg == ".":
            continue
        elif seg == "..":
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return "/".join(parts)


def read_image_data_uri(zf, zip_path, budget):
    if not zip_path or zip_path not in zf.namelist():
        return None
    ext = os.path.splitext(zip_path)[1].lower()
    mime = MIME_MAP.get(ext)
    if not mime:
        return None
    try:
        raw = zf.read(zip_path)
    except Exception:
        return None
    if len(raw) > MAX_IMAGE_BYTES or budget["used"] + len(raw) > MAX_TOTAL_IMAGE_BYTES:
        budget["skipped"] += 1
        return None
    budget["used"] += len(raw)
    b64 = base64.b64encode(raw).decode("ascii")
    return "data:%s;base64,%s" % (mime, b64)


# =============================================================================
# DOCX PARSER
# =============================================================================
def parse_docx(file_path):
    with zipfile.ZipFile(file_path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            raise ValueError("Invalid .docx file: missing word/document.xml")

        rels = parse_rels(zf, "word/_rels/document.xml.rels")

        # Parse styles.xml to map styleId -> style metadata
        styles = {}
        if "word/styles.xml" in zf.namelist():
            try:
                sroot = ET.fromstring(zf.read("word/styles.xml"))
                for st in sroot.findall(qn("w", "style")):
                    sid = st.attrib.get(qn("w", "styleId"), "")
                    name_el = st.find(qn("w", "name"))
                    sname = name_el.attrib.get(qn("w", "val"), "") if name_el is not None else ""
                    rpr = st.find(qn("w", "rPr"))
                    sz = None
                    bold = False
                    color = None
                    if rpr is not None:
                        sz_el = rpr.find(qn("w", "sz"))
                        if sz_el is not None and sz_el.attrib.get(qn("w", "val"), "").isdigit():
                            sz = int(sz_el.attrib.get(qn("w", "val"))) / 2.0
                        if rpr.find(qn("w", "b")) is not None:
                            bold = True
                        col_el = rpr.find(qn("w", "color"))
                        if col_el is not None:
                            color = safe_color(col_el.attrib.get(qn("w", "val")))
                    styles[sid] = {"name": sname, "sz": sz, "bold": bold, "color": color}
            except Exception:
                pass

        # Parse numbering.xml to distinguish bullet vs numbered lists
        num_formats = {}
        if "word/numbering.xml" in zf.namelist():
            try:
                nroot = ET.fromstring(zf.read("word/numbering.xml"))
                abstract_fmts = {}
                for ab in nroot.findall(qn("w", "abstractNum")):
                    aid = ab.attrib.get(qn("w", "abstractNumId"))
                    lvls = {}
                    for lvl in ab.findall(qn("w", "lvl")):
                        ilvl = lvl.attrib.get(qn("w", "ilvl"), "0")
                        nf = lvl.find(qn("w", "numFmt"))
                        lvls[ilvl] = nf.attrib.get(qn("w", "val"), "bullet") if nf is not None else "bullet"
                    abstract_fmts[aid] = lvls
                for num in nroot.findall(qn("w", "num")):
                    nid = num.attrib.get(qn("w", "numId"))
                    aref = num.find(qn("w", "abstractNumId"))
                    if aref is not None:
                        aid = aref.attrib.get(qn("w", "val"))
                        num_formats[nid] = abstract_fmts.get(aid, {})
            except Exception:
                pass

        # Extract header text (if any)
        header_text = ""
        for hcand in ("word/header1.xml", "word/header2.xml"):
            if hcand in zf.namelist():
                try:
                    hroot = ET.fromstring(zf.read(hcand))
                    ht = "".join(t.text or "" for t in hroot.iter(qn("w", "t"))).strip()
                    if ht:
                        header_text = ht
                        break
                except Exception:
                    pass

        root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find(qn("w", "body"))
        if body is None:
            raise ValueError("Invalid .docx file: missing w:body")

        outline = []
        stats = {"paragraphs": 0, "tables": 0, "images": 0, "words": 0, "pageBreaks": 0}
        img_budget = {"used": 0, "skipped": 0}
        list_counters = {}

        def render_run(r_el, base_style=None):
            rpr = r_el.find(qn("w", "rPr"))
            css = []
            tag_open = ""
            tag_close = ""

            is_bold = base_style.get("bold", False) if base_style else False
            is_italic = False
            is_underline = False
            is_strike = False
            vert = None
            color = base_style.get("color") if base_style else None
            sz_pt = base_style.get("sz") if base_style else None
            highlight = None

            if rpr is not None:
                b_el = rpr.find(qn("w", "b"))
                if b_el is not None:
                    bval = b_el.attrib.get(qn("w", "val"), "1").lower()
                    is_bold = bval not in ("0", "false", "off")
                i_el = rpr.find(qn("w", "i"))
                if i_el is not None:
                    ival = i_el.attrib.get(qn("w", "val"), "1").lower()
                    is_italic = ival not in ("0", "false", "off")
                u_el = rpr.find(qn("w", "u"))
                if u_el is not None and u_el.attrib.get(qn("w", "val"), "single") != "none":
                    is_underline = True
                st_el = rpr.find(qn("w", "strike"))
                if st_el is not None and st_el.attrib.get(qn("w", "val"), "1").lower() not in ("0", "false"):
                    is_strike = True
                va_el = rpr.find(qn("w", "vertAlign"))
                if va_el is not None:
                    vert = va_el.attrib.get(qn("w", "val"))
                col_el = rpr.find(qn("w", "color"))
                if col_el is not None:
                    c = safe_color(col_el.attrib.get(qn("w", "val")))
                    if c:
                        color = c
                sz_el = rpr.find(qn("w", "sz"))
                if sz_el is not None and sz_el.attrib.get(qn("w", "val"), "").isdigit():
                    sz_pt = int(sz_el.attrib.get(qn("w", "val"))) / 2.0
                shd_el = rpr.find(qn("w", "shd"))
                if shd_el is not None:
                    hf = safe_color(shd_el.attrib.get(qn("w", "fill")))
                    if hf:
                        highlight = hf
                hl_el = rpr.find(qn("w", "highlight"))
                if hl_el is not None:
                    hl_map = {"yellow": "#FEF08A", "green": "#BBF7D0", "cyan": "#A5F3FC", "magenta": "#F5D0FE", "red": "#FECACA", "lightGray": "#E5E7EB"}
                    highlight = hl_map.get(hl_el.attrib.get(qn("w", "val"), ""), highlight)

            if is_bold:
                css.append("font-weight:700")
            if is_italic:
                css.append("font-style:italic")
            deco = []
            if is_underline:
                deco.append("underline")
            if is_strike:
                deco.append("line-through")
            if deco:
                css.append("text-decoration:" + " ".join(deco))
            if color:
                css.append("color:" + color)
            if sz_pt:
                css.append("font-size:%.1fpt" % max(7.5, min(sz_pt, 36.0)))
            if highlight:
                css.append("background-color:" + highlight + ";padding:0 2px;border-radius:2px")
            if vert == "superscript":
                tag_open, tag_close = "<sup>", "</sup>"
            elif vert == "subscript":
                tag_open, tag_close = "<sub>", "</sub>"

            pieces = []
            has_page_break = False
            for child in r_el:
                tag = child.tag
                if tag == qn("w", "t"):
                    if child.text:
                        pieces.append(html_lib.escape(child.text))
                elif tag == qn("w", "tab"):
                    pieces.append("&emsp;")
                elif tag == qn("w", "br"):
                    if child.attrib.get(qn("w", "type")) == "page":
                        has_page_break = True
                    else:
                        pieces.append("<br/>")
                elif tag == qn("w", "drawing") or tag == qn("w", "pict"):
                    for blip in child.iter(qn("a", "blip")):
                        embed = blip.attrib.get(qn("r", "embed"))
                        if embed and embed in rels:
                            zpath = resolve_zip_path("word", rels[embed]["target"])
                            uri = read_image_data_uri(zf, zpath, img_budget)
                            if uri:
                                stats["images"] += 1
                                pieces.append('<img class="ov-img" src="%s" alt="embedded image"/>' % uri)
                            else:
                                pieces.append('<span class="ov-img-skip">[内嵌图片过大或格式不支持: %s]</span>' % html_lib.escape(os.path.basename(zpath)))

            inner = "".join(pieces)
            if has_page_break and not inner:
                stats["pageBreaks"] += 1
                return '<div class="ov-page-break"><span>分页符</span></div>'
            if not inner:
                return ""
            if css:
                res = '<span style="%s">%s%s%s</span>' % (";".join(css), tag_open, inner, tag_close)
            elif tag_open:
                res = "%s%s%s" % (tag_open, inner, tag_close)
            else:
                res = inner
            if has_page_break:
                stats["pageBreaks"] += 1
                res += '<div class="ov-page-break"><span>分页符</span></div>'
            return res

        def detect_heading_level(p_el, style_id, style_info, plain_text):
            if not plain_text or len(plain_text) > 120:
                return 0
            sname = (style_info.get("name", "") if style_info else "") + " " + (style_id or "")
            m = re.search(r"(?:heading|标题)\s*([1-6])", sname, re.I)
            if m:
                return int(m.group(1))
            if re.search(r"\btitle\b", sname, re.I):
                return 1
            # Heuristic for hand-styled SOP documents (size >= 13.5pt bold -> h1, >= 11.5pt bold -> h2)
            r_els = p_el.findall(qn("w", "r"))
            if r_els:
                rpr = r_els[0].find(qn("w", "rPr"))
                if rpr is not None and rpr.find(qn("w", "b")) is not None:
                    sz_el = rpr.find(qn("w", "sz"))
                    if sz_el is not None and sz_el.attrib.get(qn("w", "val"), "").isdigit():
                        pt = int(sz_el.attrib.get(qn("w", "val"))) / 2.0
                        if pt >= 13.5:
                            return 1
                        if pt >= 11.2:
                            return 2
            return 0

        def render_paragraph(p_el, in_table=False):
            ppr = p_el.find(qn("w", "pPr"))
            style_id = ""
            style_info = None
            p_css = []
            list_prefix = ""

            if ppr is not None:
                ps = ppr.find(qn("w", "pStyle"))
                if ps is not None:
                    style_id = ps.attrib.get(qn("w", "val"), "")
                    style_info = styles.get(style_id)
                jc = ppr.find(qn("w", "jc"))
                if jc is not None:
                    jval = jc.attrib.get(qn("w", "val"), "")
                    align_map = {"center": "center", "right": "right", "both": "justify", "left": "left"}
                    if jval in align_map:
                        p_css.append("text-align:" + align_map[jval])
                sp = ppr.find(qn("w", "spacing"))
                if sp is not None:
                    sb = sp.attrib.get(qn("w", "before"), "")
                    sa = sp.attrib.get(qn("w", "after"), "")
                    if sb.isdigit() and int(sb) > 0:
                        p_css.append("margin-top:%.1fpt" % min(int(sb) / 20.0, 24.0))
                    if sa.isdigit():
                        p_css.append("margin-bottom:%.1fpt" % min(int(sa) / 20.0, 20.0))
                ind = ppr.find(qn("w", "ind"))
                if ind is not None:
                    left_v = ind.attrib.get(qn("w", "left"), "")
                    hang_v = ind.attrib.get(qn("w", "hanging"), "")
                    first_v = ind.attrib.get(qn("w", "firstLine"), "")
                    if left_v.isdigit() and int(left_v) > 0:
                        p_css.append("padding-left:%.1fpt" % min(int(left_v) / 20.0, 72.0))
                    if hang_v.isdigit() and int(hang_v) > 0:
                        p_css.append("text-indent:-%.1fpt" % min(int(hang_v) / 20.0, 36.0))
                    elif first_v.isdigit() and int(first_v) > 0:
                        p_css.append("text-indent:%.1fpt" % min(int(first_v) / 20.0, 36.0))
                shd = ppr.find(qn("w", "shd"))
                if shd is not None:
                    fill = safe_color(shd.attrib.get(qn("w", "fill")))
                    if fill:
                        p_css.append("background-color:%s;padding:4px 8px;border-radius:3px" % fill)
                pbdr = ppr.find(qn("w", "pBdr"))
                if pbdr is not None:
                    for side, css_prop in (("bottom", "border-bottom"), ("left", "border-left"), ("top", "border-top")):
                        s_el = pbdr.find(qn("w", side))
                        if s_el is not None and s_el.attrib.get(qn("w", "val"), "none") != "none":
                            bcol = safe_color(s_el.attrib.get(qn("w", "color"))) or "#1B365D"
                            sz_v = s_el.attrib.get(qn("w", "sz"), "6")
                            px = max(1, min(4, int(sz_v) // 6)) if sz_v.isdigit() else 1
                            p_css.append("%s:%dpx solid %s;padding-%s:4px" % (css_prop, px, bcol, side))
                numpr = ppr.find(qn("w", "numPr"))
                if numpr is not None:
                    ilvl_el = numpr.find(qn("w", "ilvl"))
                    nid_el = numpr.find(qn("w", "numId"))
                    ilvl = ilvl_el.attrib.get(qn("w", "val"), "0") if ilvl_el is not None else "0"
                    nid = nid_el.attrib.get(qn("w", "val"), "") if nid_el is not None else ""
                    if nid and nid != "0":
                        fmt = num_formats.get(nid, {}).get(ilvl, "bullet")
                        lvl_num = int(ilvl) if ilvl.isdigit() else 0
                        p_css.append("padding-left:%dpx;text-indent:-16px" % (22 + lvl_num * 18))
                        if fmt == "bullet":
                            list_prefix = '<span class="ov-bullet">%s</span>' % ("• " if lvl_num == 0 else "◦ ")
                        else:
                            key = (nid, ilvl)
                            list_counters[key] = list_counters.get(key, 0) + 1
                            list_prefix = '<span class="ov-bullet">%d. </span>' % list_counters[key]

            content_parts = []
            plain_parts = []
            for child in p_el:
                if child.tag == qn("w", "r"):
                    content_parts.append(render_run(child, style_info))
                    for t in child.findall(qn("w", "t")):
                        if t.text:
                            plain_parts.append(t.text)
                elif child.tag == qn("w", "hyperlink"):
                    rid = child.attrib.get(qn("r", "id"))
                    href = rels.get(rid, {}).get("target", "") if rid else ""
                    h_inner = "".join(render_run(r, style_info) for r in child.findall(qn("w", "r")))
                    for t in child.iter(qn("w", "t")):
                        if t.text:
                            plain_parts.append(t.text)
                    if href and href.startswith(("http://", "https://", "mailto:")):
                        content_parts.append('<a class="ov-link" href="%s" target="_blank" rel="noopener">%s</a>' % (html_lib.escape(href), h_inner))
                    else:
                        content_parts.append(h_inner)

            inner_html = "".join(content_parts)
            plain_text = "".join(plain_parts).strip()

            if inner_html.startswith('<div class="ov-page-break">') and not plain_text:
                return inner_html

            if plain_text:
                stats["paragraphs"] += 1
                # Count CJK chars + English words
                cjk = len(re.findall(r"[\u4e00-\u9fff]", plain_text))
                eng = len(re.findall(r"[A-Za-z0-9_]+", plain_text))
                stats["words"] += cjk + eng

            if not inner_html and not in_table:
                return '<div class="ov-spacer"></div>'

            anchor = ""
            if not in_table and plain_text:
                hlvl = detect_heading_level(p_el, style_id, style_info, plain_text)
                if hlvl > 0:
                    hid = "ov-h-%d" % len(outline)
                    outline.append({"id": hid, "level": hlvl, "text": plain_text[:80]})
                    anchor = ' id="%s" class="ov-p ov-h%d"' % (hid, hlvl)

            cls = anchor if anchor else ' class="ov-p"'
            style_attr = (' style="%s"' % ";".join(p_css)) if p_css else ""
            return "<p%s%s>%s%s</p>" % (cls, style_attr, list_prefix, inner_html or "&nbsp;")

        def render_table(tbl_el):
            stats["tables"] += 1
            # Column widths from tblGrid
            col_ws = []
            tbl_grid = tbl_el.find(qn("w", "tblGrid"))
            if tbl_grid is not None:
                for gc in tbl_grid.findall(qn("w", "gridCol")):
                    w = gc.attrib.get(qn("w", "w"), "")
                    if w.isdigit():
                        col_ws.append(int(w))
            total_w = sum(col_ws) if col_ws else 0

            # Detect table-level borders (for callout boxes vs data tables)
            tbl_pr = tbl_el.find(qn("w", "tblPr"))
            tbl_bdr_css = []
            if tbl_pr is not None:
                tb = tbl_pr.find(qn("w", "tblBorders"))
                if tb is not None:
                    for side in ("top", "bottom", "left", "right"):
                        s_el = tb.find(qn("w", side))
                        if s_el is not None:
                            val = s_el.attrib.get(qn("w", "val"), "")
                            if val == "none" or val == "nil":
                                tbl_bdr_css.append("border-%s:none" % side)
                            elif val:
                                col = safe_color(s_el.attrib.get(qn("w", "color"))) or "#CBD5E0"
                                sz = s_el.attrib.get(qn("w", "sz"), "4")
                                px = max(1, min(5, int(sz) // 6)) if sz.isdigit() else 1
                                tbl_bdr_css.append("border-%s:%dpx solid %s" % (side, px, col))

            # First pass: collect rows & resolve vMerge rowspans
            trs = tbl_el.findall(qn("w", "tr"))
            grid_rows = []
            for tr in trs:
                row_cells = []
                col_cursor = 0
                for tc in tr.findall(qn("w", "tc")):
                    tcpr = tc.find(qn("w", "tcPr"))
                    colspan = 1
                    vmerge = None
                    bg = None
                    valign = "middle"
                    tc_w = None
                    if tcpr is not None:
                        gs = tcpr.find(qn("w", "gridSpan"))
                        if gs is not None and gs.attrib.get(qn("w", "val"), "").isdigit():
                            colspan = int(gs.attrib.get(qn("w", "val")))
                        vm = tcpr.find(qn("w", "vMerge"))
                        if vm is not None:
                            vmerge = vm.attrib.get(qn("w", "val"), "continue")
                        shd = tcpr.find(qn("w", "shd"))
                        if shd is not None:
                            bg = safe_color(shd.attrib.get(qn("w", "fill")))
                        va = tcpr.find(qn("w", "vAlign"))
                        if va is not None:
                            vmap = {"top": "top", "center": "middle", "bottom": "bottom"}
                            valign = vmap.get(va.attrib.get(qn("w", "val"), ""), "middle")
                        tcw_el = tcpr.find(qn("w", "tcW"))
                        if tcw_el is not None and tcw_el.attrib.get(qn("w", "w"), "").isdigit():
                            tc_w = int(tcw_el.attrib.get(qn("w", "w")))
                    cell_html = []
                    for ch in tc:
                        if ch.tag == qn("w", "p"):
                            cell_html.append(render_paragraph(ch, in_table=True))
                        elif ch.tag == qn("w", "tbl"):
                            cell_html.append(render_table(ch))
                    row_cells.append({
                        "col": col_cursor,
                        "colspan": colspan,
                        "rowspan": 1,
                        "vmerge": vmerge,
                        "skip": False,
                        "bg": bg,
                        "valign": valign,
                        "w": tc_w,
                        "html": "".join(cell_html) or "&nbsp;",
                    })
                    col_cursor += colspan
                grid_rows.append(row_cells)

            # Resolve vertical merges
            for r_i, row in enumerate(grid_rows):
                for cell in row:
                    if cell["vmerge"] == "continue":
                        cell["skip"] = True
                        for up in range(r_i - 1, -1, -1):
                            matched = False
                            for ucell in grid_rows[up]:
                                if ucell["col"] == cell["col"] and not ucell["skip"]:
                                    ucell["rowspan"] += 1
                                    matched = True
                                    break
                            if matched:
                                break

            is_single_box = (len(grid_rows) == 1 and len(grid_rows[0]) == 1)
            tbl_cls = "ov-callout-tbl" if is_single_box else "ov-table"
            style_str = (' style="%s"' % ";".join(tbl_bdr_css)) if (is_single_box and tbl_bdr_css) else ""
            out = ['<div class="ov-table-wrap"><table class="%s"%s>' % (tbl_cls, style_str)]
            if total_w > 0 and not is_single_box:
                out.append("<colgroup>")
                for cw in col_ws:
                    out.append('<col style="width:%.2f%%"/>' % (cw * 100.0 / total_w))
                out.append("</colgroup>")
            out.append("<tbody>")
            for row in grid_rows:
                out.append("<tr>")
                for cell in row:
                    if cell["skip"]:
                        continue
                    attrs = []
                    if cell["colspan"] > 1:
                        attrs.append('colspan="%d"' % cell["colspan"])
                    if cell["rowspan"] > 1:
                        attrs.append('rowspan="%d"' % cell["rowspan"])
                    c_css = ["vertical-align:" + cell["valign"]]
                    if cell["bg"]:
                        c_css.append("background-color:" + cell["bg"])
                    attrs.append('style="%s"' % ";".join(c_css))
                    out.append("<td %s>%s</td>" % (" ".join(attrs), cell["html"]))
                out.append("</tr>")
            out.append("</tbody></table></div>")
            return "".join(out)

        body_parts = []
        for child in body:
            if child.tag == qn("w", "p"):
                body_parts.append(render_paragraph(child, in_table=False))
            elif child.tag == qn("w", "tbl"):
                body_parts.append(render_table(child))

        return {
            "kind": "docx",
            "headerText": header_text,
            "html": "".join(body_parts),
            "outline": outline,
            "stats": stats,
            "skippedImages": img_budget["skipped"],
        }


# =============================================================================
# XLSX PARSER
# =============================================================================
def col_letters_to_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n - 1


def idx_to_col_letters(idx):
    res = []
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res.append(chr(ord("A") + rem))
    return "".join(reversed(res))


def excel_serial_to_date(val):
    try:
        f = float(val)
        if f < 1 or f > 2958465:
            return str(val)
        base = datetime.datetime(1899, 12, 30)
        dt = base + datetime.timedelta(days=f)
        if abs(f - int(f)) < 1e-6:
            return dt.strftime("%Y-%m-%d")
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(val)


def parse_xlsx(file_path):
    with zipfile.ZipFile(file_path, "r") as zf:
        if "xl/workbook.xml" not in zf.namelist():
            raise ValueError("Invalid .xlsx file: missing xl/workbook.xml")

        # Shared strings
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            try:
                ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in ss_root.findall(qn("x", "si")):
                    txt = "".join(t.text or "" for t in si.iter(qn("x", "t")))
                    shared_strings.append(txt)
            except Exception:
                pass

        # Styles: detect date formats, bold fonts, and fill colors
        date_xf = set()
        bold_xf = set()
        fill_xf = {}
        if "xl/styles.xml" in zf.namelist():
            try:
                st_root = ET.fromstring(zf.read("xl/styles.xml"))
                custom_date_fmts = set()
                num_fmts = st_root.find(qn("x", "numFmts"))
                if num_fmts is not None:
                    for nf in num_fmts.findall(qn("x", "numFmt")):
                        fid = int(nf.attrib.get("numFmtId", "0"))
                        code = nf.attrib.get("formatCode", "").lower()
                        if any(k in code for k in ("yy", "mm", "dd")) and "0" not in code:
                            custom_date_fmts.add(fid)
                builtin_date = set(range(14, 23)) | set(range(45, 48))

                bold_fonts = set()
                fonts = st_root.find(qn("x", "fonts"))
                if fonts is not None:
                    for i, f_el in enumerate(fonts.findall(qn("x", "font"))):
                        if f_el.find(qn("x", "b")) is not None:
                            bold_fonts.add(i)

                fill_colors = {}
                fills = st_root.find(qn("x", "fills"))
                if fills is not None:
                    for i, fl in enumerate(fills.findall(qn("x", "fill"))):
                        pf = fl.find(qn("x", "patternFill"))
                        if pf is not None and pf.attrib.get("patternType") == "solid":
                            fg = pf.find(qn("x", "fgColor"))
                            if fg is not None:
                                c = safe_color(fg.attrib.get("rgb"))
                                if c and c.upper() != "#000000":
                                    fill_colors[i] = c

                cxfs = st_root.find(qn("x", "cellXfs"))
                if cxfs is not None:
                    for idx, xf in enumerate(cxfs.findall(qn("x", "xf"))):
                        nfid = int(xf.attrib.get("numFmtId", "0"))
                        if nfid in builtin_date or nfid in custom_date_fmts:
                            date_xf.add(idx)
                        fid = int(xf.attrib.get("fontId", "0"))
                        if fid in bold_fonts:
                            bold_xf.add(idx)
                        flid = int(xf.attrib.get("fillId", "0"))
                        if flid in fill_colors:
                            fill_xf[idx] = fill_colors[flid]
            except Exception:
                pass

        wb_rels = parse_rels(zf, "xl/_rels/workbook.xml.rels")
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets_out = []
        sheets_el = wb_root.find(qn("x", "sheets"))
        if sheets_el is None:
            raise ValueError("No worksheets found")

        for sh in sheets_el.findall(qn("x", "sheet")):
            name = sh.attrib.get("name", "Sheet")
            rid = sh.attrib.get(qn("r", "id"), "")
            rel = wb_rels.get(rid)
            if not rel:
                continue
            spath = resolve_zip_path("xl", rel["target"])
            if spath not in zf.namelist():
                continue

            sroot = ET.fromstring(zf.read(spath))
            # Parse merges
            merges = {}
            skip_coords = set()
            mc_el = sroot.find(qn("x", "mergeCells"))
            if mc_el is not None:
                for mc in mc_el.findall(qn("x", "mergeCell")):
                    ref = mc.attrib.get("ref", "")
                    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
                    if m:
                        c1, r1 = col_letters_to_idx(m.group(1)), int(m.group(2)) - 1
                        c2, r2 = col_letters_to_idx(m.group(3)), int(m.group(4)) - 1
                        merges[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
                        for rr in range(r1, r2 + 1):
                            for cc in range(c1, c2 + 1):
                                if (rr, cc) != (r1, c1):
                                    skip_coords.add((rr, cc))

            sheet_data = sroot.find(qn("x", "sheetData"))
            cell_map = {}
            max_r = 0
            max_c = 0
            truncated_rows = False
            truncated_cols = False

            if sheet_data is not None:
                for row_el in sheet_data.findall(qn("x", "row")):
                    r_attr = row_el.attrib.get("r")
                    r_idx = (int(r_attr) - 1) if (r_attr and r_attr.isdigit()) else max_r
                    if r_idx >= MAX_XLSX_ROWS:
                        truncated_rows = True
                        continue
                    for c_el in row_el.findall(qn("x", "c")):
                        cref = c_el.attrib.get("r", "")
                        cm = re.match(r"([A-Z]+)(\d+)", cref)
                        if not cm:
                            continue
                        c_idx = col_letters_to_idx(cm.group(1))
                        rr_idx = int(cm.group(2)) - 1
                        if rr_idx >= MAX_XLSX_ROWS:
                            truncated_rows = True
                            continue
                        if c_idx >= MAX_XLSX_COLS:
                            truncated_cols = True
                            continue
                        ctype = c_el.attrib.get("t", "")
                        s_idx = int(c_el.attrib.get("s", "-1")) if c_el.attrib.get("s", "").isdigit() else -1
                        val = ""
                        if ctype == "s":
                            v_el = c_el.find(qn("x", "v"))
                            if v_el is not None and v_el.text and v_el.text.isdigit():
                                idx_ss = int(v_el.text)
                                val = shared_strings[idx_ss] if idx_ss < len(shared_strings) else ""
                        elif ctype == "inlineStr":
                            is_el = c_el.find(qn("x", "is"))
                            if is_el is not None:
                                val = "".join(t.text or "" for t in is_el.iter(qn("x", "t")))
                        elif ctype == "b":
                            v_el = c_el.find(qn("x", "v"))
                            val = "TRUE" if (v_el is not None and v_el.text == "1") else "FALSE"
                        else:
                            v_el = c_el.find(qn("x", "v"))
                            if v_el is not None and v_el.text is not None:
                                val = v_el.text
                                if ctype != "str" and s_idx in date_xf:
                                    val = excel_serial_to_date(val)
                        is_num = (ctype in ("", "n") and val != "" and s_idx not in date_xf and re.fullmatch(r"-?\d+(\.\d+)?([eE][+-]?\d+)?", val))
                        if is_num and "." in val:
                            try:
                                fv = float(val)
                                val = ("%g" % fv) if abs(fv) < 1e12 else val
                            except Exception:
                                pass
                        cell_map[(rr_idx, c_idx)] = {
                            "v": val,
                            "b": s_idx in bold_xf,
                            "bg": fill_xf.get(s_idx),
                            "num": bool(is_num),
                        }
                        if val != "" or (rr_idx, c_idx) in merges:
                            max_r = max(max_r, rr_idx + 1)
                            max_c = max(max_c, c_idx + 1)

            # Build clean HTML table for this sheet
            rows_html = []
            if max_r == 0 or max_c == 0:
                table_html = '<div class="ov-empty">此工作表为空</div>'
            else:
                hdr_cells = ['<th class="ov-xl-rn"></th>']
                for c in range(max_c):
                    hdr_cells.append("<th>%s</th>" % idx_to_col_letters(c))
                rows_html.append('<table class="ov-xl-table"><thead><tr>%s</tr></thead><tbody>' % "".join(hdr_cells))
                for r in range(max_r):
                    tr_cells = ['<td class="ov-xl-rn">%d</td>' % (r + 1)]
                    for c in range(max_c):
                        if (r, c) in skip_coords:
                            continue
                        cell = cell_map.get((r, c))
                        attrs = []
                        if (r, c) in merges:
                            rs, cs = merges[(r, c)]
                            if rs > 1:
                                attrs.append('rowspan="%d"' % rs)
                            if cs > 1:
                                attrs.append('colspan="%d"' % cs)
                        css = []
                        if cell:
                            if cell["b"]:
                                css.append("font-weight:700")
                            if cell["bg"]:
                                css.append("background-color:" + cell["bg"])
                            if cell["num"]:
                                css.append("text-align:right;font-variant-numeric:tabular-nums")
                        if css:
                            attrs.append('style="%s"' % ";".join(css))
                        attr_s = (" " + " ".join(attrs)) if attrs else ""
                        txt = html_lib.escape(cell["v"]).replace("\n", "<br/>") if cell else ""
                        tr_cells.append("<td%s>%s</td>" % (attr_s, txt))
                    rows_html.append("<tr>%s</tr>" % "".join(tr_cells))
                rows_html.append("</tbody></table>")
                table_html = "".join(rows_html)

            sheets_out.append({
                "name": name,
                "rows": max_r,
                "cols": max_c,
                "truncated": truncated_rows or truncated_cols,
                "html": table_html,
            })

        return {
            "kind": "xlsx",
            "sheets": sheets_out,
            "stats": {
                "sheets": len(sheets_out),
                "totalRows": sum(s["rows"] for s in sheets_out),
            },
        }


# =============================================================================
# PPTX PARSER
# =============================================================================
def parse_pptx(file_path):
    with zipfile.ZipFile(file_path, "r") as zf:
        if "ppt/presentation.xml" not in zf.namelist():
            raise ValueError("Invalid .pptx file: missing ppt/presentation.xml")

        pres_rels = parse_rels(zf, "ppt/_rels/presentation.xml.rels")
        pres_root = ET.fromstring(zf.read("ppt/presentation.xml"))
        sld_id_lst = pres_root.find(qn("p", "sldIdLst"))
        slide_paths = []
        if sld_id_lst is not None:
            for sid in sld_id_lst.findall(qn("p", "sldId")):
                rid = sid.attrib.get(qn("r", "id"))
                if rid and rid in pres_rels:
                    sp = resolve_zip_path("ppt", pres_rels[rid]["target"])
                    if sp in zf.namelist():
                        slide_paths.append(sp)
        if not slide_paths:
            slide_paths = sorted(
                [n for n in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
                key=lambda s: int(re.search(r"(\d+)\.xml$", s).group(1)),
            )

        slides_out = []
        img_budget = {"used": 0, "skipped": 0}

        for idx, spath in enumerate(slide_paths, 1):
            sroot = ET.fromstring(zf.read(spath))
            s_dir = os.path.dirname(spath)
            s_base = os.path.basename(spath)
            s_rels = parse_rels(zf, "%s/_rels/%s.rels" % (s_dir, s_base))

            title = ""
            blocks_html = []

            # Extract speaker notes if linked
            notes_text = ""
            for rel in s_rels.values():
                if "notesSlide" in rel.get("type", ""):
                    npath = resolve_zip_path(s_dir, rel["target"])
                    if npath in zf.namelist():
                        try:
                            nroot = ET.fromstring(zf.read(npath))
                            nt = "\n".join(
                                "".join(t.text or "" for t in p.iter(qn("a", "t"))).strip()
                                for p in nroot.iter(qn("a", "p"))
                            ).strip()
                            notes_text = re.sub(r"\n{3,}", "\n\n", nt)
                        except Exception:
                            pass

            # Walk shapes in spTree
            sp_tree = sroot.find(".//" + qn("p", "spTree"))
            if sp_tree is not None:
                for elem in sp_tree:
                    if elem.tag == qn("p", "sp"):
                        ph = elem.find(".//" + qn("p", "ph"))
                        ph_type = ph.attrib.get("type", "") if ph is not None else ""
                        tx = elem.find(qn("p", "txBody"))
                        if tx is None:
                            continue
                        paras = []
                        for p_el in tx.findall(qn("a", "p")):
                            ppr = p_el.find(qn("a", "pPr"))
                            lvl = int(ppr.attrib.get("lvl", "0")) if (ppr is not None and ppr.attrib.get("lvl", "").isdigit()) else 0
                            runs = []
                            for r_el in p_el.findall(qn("a", "r")):
                                t_el = r_el.find(qn("a", "t"))
                                if t_el is None or not t_el.text:
                                    continue
                                rpr = r_el.find(qn("a", "rPr"))
                                css = []
                                if rpr is not None:
                                    if rpr.attrib.get("b") == "1":
                                        css.append("font-weight:700")
                                    if rpr.attrib.get("i") == "1":
                                        css.append("font-style:italic")
                                    sfill = rpr.find(qn("a", "solidFill"))
                                    if sfill is not None:
                                        srgb = sfill.find(qn("a", "srgbClr"))
                                        if srgb is not None:
                                            c = safe_color(srgb.attrib.get("val"))
                                            if c:
                                                css.append("color:" + c)
                                esc_t = html_lib.escape(t_el.text)
                                runs.append(('<span style="%s">%s</span>' % (";".join(css), esc_t)) if css else esc_t)
                            line_html = "".join(runs).strip()
                            if line_html:
                                paras.append((lvl, line_html))
                        if not paras:
                            continue
                        if ph_type in ("title", "ctrTitle") and not title:
                            title = re.sub(r"<[^>]+>", "", paras[0][1])
                            blocks_html.append('<h2 class="ov-ppt-title">%s</h2>' % "<br/>".join(x[1] for x in paras))
                        elif ph_type == "subTitle":
                            blocks_html.append('<p class="ov-ppt-subtitle">%s</p>' % "<br/>".join(x[1] for x in paras))
                        else:
                            for lvl, lh in paras:
                                pad = 16 + lvl * 18
                                bullet = "• " if (len(paras) > 1 or lvl > 0) else ""
                                blocks_html.append('<p class="ov-ppt-p" style="padding-left:%dpx">%s%s</p>' % (pad, bullet, lh))

                    elif elem.tag == qn("p", "pic"):
                        for blip in elem.iter(qn("a", "blip")):
                            embed = blip.attrib.get(qn("r", "embed"))
                            if embed and embed in s_rels:
                                ipath = resolve_zip_path(s_dir, s_rels[embed]["target"])
                                uri = read_image_data_uri(zf, ipath, img_budget)
                                if uri:
                                    blocks_html.append('<div class="ov-ppt-img-wrap"><img class="ov-img" src="%s" alt="slide image"/></div>' % uri)

                    elif elem.tag == qn("p", "graphicFrame"):
                        tbl = elem.find(".//" + qn("a", "tbl"))
                        if tbl is not None:
                            t_rows = ['<div class="ov-table-wrap"><table class="ov-table"><tbody>']
                            for tr in tbl.findall(qn("a", "tr")):
                                t_rows.append("<tr>")
                                for tc in tr.findall(qn("a", "tc")):
                                    txt = "<br/>".join(
                                        html_lib.escape("".join(t.text or "" for t in p.iter(qn("a", "t"))).strip())
                                        for p in tc.iter(qn("a", "p"))
                                    )
                                    t_rows.append("<td>%s</td>" % (txt or "&nbsp;"))
                                t_rows.append("</tr>")
                            t_rows.append("</tbody></table></div>")
                            blocks_html.append("".join(t_rows))

            slides_out.append({
                "index": idx,
                "title": title or ("Slide %d" % idx),
                "html": "".join(blocks_html) or '<div class="ov-empty">（本页无文本或图片元素）</div>',
                "notes": notes_text,
            })

        return {
            "kind": "pptx",
            "slides": slides_out,
            "stats": {"slides": len(slides_out)},
            "skippedImages": img_budget["skipped"],
        }


# =============================================================================
# LOCAL OFFICE APP AUTO-DETECTION & EXTERNAL LAUNCHER (macOS)
# =============================================================================
OFFICE_APP_CATALOG = [
    # (id, display_name, short_badge, badge_color, bundle_candidates, supported_exts, primary_exts)
    ("ms-word", "Microsoft Word", "W", "#2563EB", ["Microsoft Word.app"], {".docx"}, {".docx"}),
    ("ms-excel", "Microsoft Excel", "X", "#059669", ["Microsoft Excel.app"], {".xlsx"}, {".xlsx"}),
    ("ms-ppt", "Microsoft PowerPoint", "P", "#EA580C", ["Microsoft PowerPoint.app"], {".pptx"}, {".pptx"}),
    ("wps", "WPS Office", "WPS", "#E11D48", ["wpsoffice.app", "WPS Office.app"], {".docx", ".xlsx", ".pptx"}, set()),
    ("pages", "Pages 文稿", "Pg", "#D97706", ["Pages.app"], {".docx"}, set()),
    ("numbers", "Numbers 表格", "Nb", "#16A34A", ["Numbers.app"], {".xlsx"}, set()),
    ("keynote", "Keynote 讲演", "Kn", "#0284C7", ["Keynote.app"], {".pptx"}, set()),
    ("libreoffice", "LibreOffice", "LO", "#475569", ["LibreOffice.app"], {".docx", ".xlsx", ".pptx"}, set()),
    ("onlyoffice", "ONLYOFFICE", "OO", "#0F766E", ["ONLYOFFICE.app", "OnlyOffice.app"], {".docx", ".xlsx", ".pptx"}, set()),
]


def detect_office_apps(ext=""):
    """Auto-detect installed desktop Office apps across macOS, Windows, and Linux."""
    ext_norm = ("." + ext.lstrip(".").lower()) if ext else ""
    found = []
    seen_ids = set()

    if sys.platform == "darwin":
        search_roots = [
            "/Applications",
            os.path.expanduser("~/Applications"),
            "/System/Applications",
        ]
        for app_id, display_name, badge, color, bundles, exts, primary_exts in OFFICE_APP_CATALOG:
            if ext_norm and ext_norm not in exts:
                continue
            matched_path = None
            for root in search_roots:
                for bname in bundles:
                    cand = os.path.join(root, bname)
                    if os.path.isdir(cand):
                        matched_path = cand
                        break
                if matched_path:
                    break
            if matched_path and app_id not in seen_ids:
                seen_ids.add(app_id)
                found.append({
                    "id": app_id,
                    "name": display_name,
                    "badge": badge,
                    "color": color,
                    "appPath": matched_path,
                    "primary": bool(ext_norm and ext_norm in primary_exts),
                })
    elif sys.platform == "win32":
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        lad = os.environ.get("LOCALAPPDATA", "")
        win_catalog = [
            ("ms-word", "Microsoft Word", "W", "#2563EB",
             [os.path.join(r, "Microsoft Office", "root", "Office16", "WINWORD.EXE") for r in (pf, pf86)], {".docx"}, {".docx"}),
            ("ms-excel", "Microsoft Excel", "X", "#059669",
             [os.path.join(r, "Microsoft Office", "root", "Office16", "EXCEL.EXE") for r in (pf, pf86)], {".xlsx"}, {".xlsx"}),
            ("ms-ppt", "Microsoft PowerPoint", "P", "#EA580C",
             [os.path.join(r, "Microsoft Office", "root", "Office16", "POWERPNT.EXE") for r in (pf, pf86)], {".pptx"}, {".pptx"}),
            ("wps", "WPS Office", "WPS", "#E11D48",
             [os.path.join(lad, "Kingsoft", "WPS Office", "ksolaunch.exe")] if lad else [], {".docx", ".xlsx", ".pptx"}, set()),
            ("libreoffice", "LibreOffice", "LO", "#475569",
             [os.path.join(r, "LibreOffice", "program", "soffice.exe") for r in (pf, pf86)], {".docx", ".xlsx", ".pptx"}, set()),
        ]
        for app_id, display_name, badge, color, cands, exts, primary_exts in win_catalog:
            if ext_norm and ext_norm not in exts:
                continue
            for cand in cands:
                if os.path.isfile(cand) and app_id not in seen_ids:
                    seen_ids.add(app_id)
                    found.append({
                        "id": app_id,
                        "name": display_name,
                        "badge": badge,
                        "color": color,
                        "appPath": cand,
                        "primary": bool(ext_norm and ext_norm in primary_exts),
                    })
                    break
    else:
        linux_catalog = [
            ("libreoffice", "LibreOffice", "LO", "#475569", ["libreoffice", "soffice"], {".docx", ".xlsx", ".pptx"}, {".docx", ".xlsx", ".pptx"}),
            ("wps", "WPS Office", "WPS", "#E11D48", ["wps", "et", "wpp"], {".docx", ".xlsx", ".pptx"}, set()),
            ("onlyoffice", "ONLYOFFICE", "OO", "#0F766E", ["onlyoffice-desktopeditors"], {".docx", ".xlsx", ".pptx"}, set()),
        ]
        for app_id, display_name, badge, color, bins, exts, primary_exts in linux_catalog:
            if ext_norm and ext_norm not in exts:
                continue
            for b in bins:
                exe = shutil.which(b)
                if exe and app_id not in seen_ids:
                    seen_ids.add(app_id)
                    found.append({
                        "id": app_id,
                        "name": display_name,
                        "badge": badge,
                        "color": color,
                        "appPath": exe,
                        "primary": bool(ext_norm and ext_norm in primary_exts),
                    })
                    break
    return found


def open_external_app(file_path, mode="default", app_path=""):
    """Safely launch local desktop Office software or reveal file in file manager."""
    if not file_path:
        raise ValueError("No active document open")
    fp = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.isfile(fp):
        raise ValueError("File not found: %s" % fp)
    ext = os.path.splitext(fp)[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise ValueError("Only Office files (.docx, .xlsx, .pptx) are allowed")

    if mode == "reveal":
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        elif sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(fp)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return "Revealed in file manager: %s" % os.path.basename(fp)

    if mode == "app" and app_path:
        detected = detect_office_apps(ext)
        allowed = {a["appPath"]: a["name"] for a in detected}
        if app_path not in allowed:
            raise ValueError("Detected app not found: %s" % app_path)
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-a", app_path, fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            subprocess.Popen([app_path, fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return "Opened in %s: %s" % (allowed[app_path], os.path.basename(fp))

    if sys.platform == "darwin":
        subprocess.Popen(["open", fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    elif sys.platform == "win32":
        os.startfile(fp)
    else:
        subprocess.Popen(["xdg-open", fp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return "Opened with default app: %s" % os.path.basename(fp)


# =============================================================================
# FILE SCANNER & PERSISTENT STATE
# =============================================================================
def scan_office_files(directory, max_depth=3):
    results = []
    if not directory or not os.path.isdir(directory):
        return results
    directory = os.path.abspath(os.path.expanduser(directory))
    base_depth = directory.rstrip(os.sep).count(os.sep)
    skip_dirs = {".git", "node_modules", "__pycache__", ".venv", "venv", ".Trash"}
    for root, dirs, files in os.walk(directory):
        depth = root.rstrip(os.sep).count(os.sep) - base_depth
        if depth >= max_depth:
            dirs[:] = []
        else:
            dirs[:] = sorted([d for d in dirs if d not in skip_dirs and not d.startswith(".")])
        for fn in sorted(files):
            if fn.startswith("~$") or fn.startswith("."):
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext in SUPPORTED_EXTS:
                full = os.path.join(root, fn)
                try:
                    st = os.stat(full)
                    results.append({
                        "name": fn,
                        "path": full,
                        "relPath": os.path.relpath(full, directory),
                        "ext": ext.lstrip("."),
                        "size": st.st_size,
                        "sizeHuman": human_size(st.st_size),
                        "mtime": st.st_mtime,
                        "mtimeHuman": format_mtime(st.st_mtime),
                    })
                except OSError:
                    continue
    results.sort(key=lambda x: x["mtime"], reverse=True)
    return results[:200]


class StateStore:
    def __init__(self, data_dir):
        self.data_dir = data_dir
        self.state_path = os.path.join(data_dir, "state.json") if data_dir else None
        self.data = {"lastFile": "", "lastDirectory": "", "recentFiles": []}
        self._load()

    def _load(self):
        if self.state_path and os.path.isfile(self.state_path):
            try:
                with open(self.state_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    self.data.update(loaded)
            except Exception:
                pass

    def save(self):
        if not self.state_path:
            return
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            with open(self.state_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def record(self, file_path=None, directory=None):
        if file_path:
            fp = os.path.abspath(os.path.expanduser(file_path))
            self.data["lastFile"] = fp
            if not directory:
                directory = os.path.dirname(fp)
            rec = [x for x in self.data.get("recentFiles", []) if x != fp]
            rec.insert(0, fp)
            self.data["recentFiles"] = rec[:15]
        if directory and os.path.isdir(os.path.expanduser(directory)):
            self.data["lastDirectory"] = os.path.abspath(os.path.expanduser(directory))
        self.save()


def load_office_payload(file_path, directory, store, allow_large=False, bump_version=True):
    resolved_dir = directory or store.data.get("lastDirectory") or ""
    if not file_path:
        # Fallback to last opened file if it still exists
        lf = store.data.get("lastFile", "")
        if lf and os.path.isfile(lf):
            file_path = lf

    if file_path:
        file_path = os.path.abspath(os.path.expanduser(file_path))
        if not resolved_dir:
            resolved_dir = os.path.dirname(file_path)

    files_in_dir = scan_office_files(resolved_dir, max_depth=3) if resolved_dir else []
    recent_meta = []
    for rp in store.data.get("recentFiles", []):
        if os.path.isfile(rp):
            try:
                st = os.stat(rp)
                recent_meta.append({
                    "name": os.path.basename(rp),
                    "path": rp,
                    "ext": os.path.splitext(rp)[1].lstrip(".").lower(),
                    "sizeHuman": human_size(st.st_size),
                    "mtimeHuman": format_mtime(st.st_mtime),
                })
            except OSError:
                pass

    # If still no file, auto-select newest file in directory if available
    if not file_path and files_in_dir:
        file_path = files_in_dir[0]["path"]

    if not file_path:
        return {
            "mode": "empty",
            "directory": resolved_dir,
            "files": files_in_dir,
            "recentFiles": recent_meta,
            "availableApps": detect_office_apps(""),
            "message": "请选择或输入本地 .docx / .xlsx / .pptx 文件路径进行预览。",
        }, "Office Viewer 已就绪，请在右侧窗口选择或输入本地 Office 文件路径。"

    if not os.path.isfile(file_path):
        return {
            "mode": "error",
            "error": "文件不存在：%s" % file_path,
            "directory": resolved_dir,
            "files": files_in_dir,
            "recentFiles": recent_meta,
        }, "无法预览，文件不存在：%s" % file_path

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        return {
            "mode": "error",
            "error": "暂不支持该文件格式（%s）。目前支持 .docx、.xlsx、.pptx。" % (ext or "未知"),
            "directory": resolved_dir,
            "files": files_in_dir,
            "recentFiles": recent_meta,
        }, "不支持的文件格式：%s" % ext

    st = os.stat(file_path)
    if bump_version:
        store.record(file_path=file_path, directory=resolved_dir)
    else:
        store.record_silent(file_path=file_path, directory=resolved_dir)
    if not files_in_dir and resolved_dir:
        files_in_dir = scan_office_files(resolved_dir, max_depth=3)

    if ext == ".docx":
        doc = parse_docx(file_path)
        summary = "已在右侧 Office Viewer 预览 Word 文档：%s（%d 段落，%d 表格，%d 图片，约 %d 字）" % (
            os.path.basename(file_path),
            doc["stats"]["paragraphs"],
            doc["stats"]["tables"],
            doc["stats"]["images"],
            doc["stats"]["words"],
        )
    elif ext == ".xlsx":
        doc = parse_xlsx(file_path)
        summary = "已在右侧 Office Viewer 预览 Excel 工作簿：%s（共 %d 个工作表，%d 行数据）" % (
            os.path.basename(file_path),
            doc["stats"]["sheets"],
            doc["stats"]["totalRows"],
        )
    else:
        doc = parse_pptx(file_path)
        summary = "已在右侧 Office Viewer 预览 PowerPoint 演示文稿：%s（共 %d 页幻灯片）" % (
            os.path.basename(file_path),
            doc["stats"]["slides"],
        )

    payload = {
        "mode": "preview",
        "file": {
            "name": os.path.basename(file_path),
            "path": file_path,
            "ext": ext.lstrip("."),
            "size": st.st_size,
            "sizeHuman": human_size(st.st_size),
            "mtimeHuman": format_mtime(st.st_mtime),
        },
        "directory": resolved_dir,
        "files": files_in_dir,
        "recentFiles": recent_meta,
        "availableApps": detect_office_apps(ext),
        "document": doc,
        "deferred": False,
    }

    # Respect Antigravity 1 MiB trajectory cap on model-visible tool results
    if not allow_large:
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        if len(encoded) > MAX_TRAJECTORY_BYTES:
            payload["document"] = {
                "kind": doc["kind"],
                "stats": doc.get("stats", {}),
            }
            payload["deferred"] = True

    return payload, summary




def default_shared_state_dir():
    env_dir = os.environ.get("OFFICE_VIEWER_DATA_DIR", "").strip()
    if env_dir:
        return env_dir
    app_data = os.environ.get("ANTIGRAVITY_APP_DATA_DIR", "").strip()
    if not app_data:
        app_data = os.path.join(os.path.expanduser("~"), ".gemini", "antigravity")
    return os.path.join(app_data, "sidecar_data", "office-viewer", "office_viewer", "data")


class SharedStateStore(StateStore):
    """Atomic shared state.json synced between MCP server and Sidecar HTTP server."""
    def __init__(self, data_dir=None):
        resolved = data_dir or default_shared_state_dir()
        super().__init__(resolved)
        self.data.setdefault("version", 0)

    def reload(self):
        self._load()
        self.data.setdefault("version", 0)
        return self.data

    def save(self):
        if not self.state_path:
            return
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            tmp = "%s.tmp.%d" % (self.state_path, os.getpid())
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.state_path)
        except Exception:
            pass

    def record(self, file_path=None, directory=None):
        self.reload()
        self.data["version"] = int(self.data.get("version", 0)) + 1
        self.data["updatedAt"] = datetime.datetime.now().isoformat(timespec="seconds")
        super().record(file_path=file_path, directory=directory)

    def record_silent(self, file_path=None, directory=None):
        """Persist last opened file/dir for UI navigation WITHOUT bumping SSE version."""
        self.reload()
        fp = os.path.abspath(os.path.expanduser(file_path)) if file_path else None
        dp = os.path.abspath(os.path.expanduser(directory)) if (directory and os.path.isdir(os.path.expanduser(directory))) else None
        changed = False
        if fp and self.data.get("lastFile") != fp:
            changed = True
        if dp and self.data.get("lastDirectory") != dp:
            changed = True
        if changed:
            super().record(file_path=fp, directory=dp)
