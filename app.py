import io
import copy
import os
import re
import sys
import json
import time
import uuid
import math
import tempfile
import threading
import zipfile
import statistics
from collections import Counter
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import pdfplumber
from pdfminer.pdftypes import resolve1
from PIL import Image as PILImage
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.section import _Header, _Footer
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break


def resource_path(rel):
    """Bundled files (templates/static/icon). EXE me ye _MEIPASS se aati hain."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


BASE_DIR = Path(__file__).resolve().parent
if getattr(sys, "frozen", False):
    APP_DATA = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "PDFConverter"
else:
    APP_DATA = BASE_DIR
STORAGE = APP_DATA / "storage"
STORAGE.mkdir(parents=True, exist_ok=True)
app = Flask(
    __name__,
    template_folder=str(resource_path("templates")),
    static_folder=str(resource_path("static")),)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
BULLET_RE = re.compile(r"^(\u2022|\u25CF|\u25AA|-|\*|\d+[\.\)]|\(?[a-zA-Z]\)|[a-zA-Z]\.)\s+")
FID_RE = re.compile(r"^[0-9a-f]{32}$")
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---- Image settings ----
MIN_IMG_PT = 6          # isse chhoti images (lines/dots) skip hongi
MAX_IMG_PX = 1600       # bohat badi image ko is size tak chhota karte hain (file size control)
MAX_IMAGES = 300        # ek PDF se zyada se zyada itni images
MAX_DOCX_WIDTH_IN = 6.5  # Word page ki usable width (image ke liye)
WORD_USABLE_IN = 6.7     # Word page ki text width (8.5in - 2 x 0.9in margin)


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def clean_text(s):
    s = CTRL_RE.sub("", s or "")
    return re.sub(r"\s+", " ", s).strip()


def clean_cell(s):
    s = CTRL_RE.sub("", s or "")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.splitlines()]
    return "\n".join(ln for ln in lines if ln)


def normalize_with_map(rows, keep=None):
    """normalize_rows jaisa hi kaam, plus ye bhi batata hai ke kaun si original rows/cols bachi.
    keep = {(row, col)} wo cells jo khali hone par bhi rakhne hain (jin me image hai)."""
    keep = keep or set()
    cleaned = [(ri, [clean_cell(c) for c in r]) for ri, r in enumerate(rows or []) if r]
    cleaned = [(ri, r) for ri, r in cleaned
               if any(r) or any((ri, ci) in keep for ci in range(len(r)))]
    if not cleaned:
        return [], [], []
    ncols = max(len(r) for _, r in cleaned)
    cleaned = [(ri, r + [""] * (ncols - len(r))) for ri, r in cleaned]
    cols = [i for i in range(ncols)
            if any(r[i] for _, r in cleaned) or any((ri, i) in keep for ri, _ in cleaned)]
    return [[r[i] for i in cols] for _, r in cleaned], [ri for ri, _ in cleaned], cols


def normalize_rows(rows):
    return normalize_with_map(rows)[0]


def looks_like_table(rows):
    if len(rows) < 3:
        return False
    ncols = len(rows[0])
    if ncols < 2 or ncols > 12:
        return False
    total = len(rows) * ncols
    filled = sum(1 for r in rows for c in r if c)
    multi = sum(1 for r in rows if sum(1 for c in r if c) >= 2)
    return filled / total >= 0.5 and multi / len(rows) >= 0.6


def _cell_of(grid, x, y):
    """Table grid me wo (row, col) dhoondta hai jisme point (x, y) aata hai."""
    for ri, row in enumerate(grid):
        for ci, c in enumerate(row):
            if c and c[0] <= x <= c[2] and c[1] <= y <= c[3]:
                return ri, ci
    return None


def _build_table(t, img_pts):
    """Ek pdfplumber table se (bbox, rows, meta) banata hai.
    meta = {'img_cells': {image_index: (row, col)}, 'col_w': [column widths in pt] ya None}"""
    raw = t.extract()
    try:
        grid = [list(r.cells) for r in t.rows]
    except Exception:
        grid = []
    keep, found = set(), {}
    for i, (cx, cy) in enumerate(img_pts or []):
        b = t.bbox
        if b[0] <= cx <= b[2] and b[1] <= cy <= b[3]:
            pos = _cell_of(grid, cx, cy)
            if pos:
                keep.add(pos)
                found[i] = pos
    rows, rmap, cmap = normalize_with_map(raw, keep)
    meta = {"img_cells": {}, "col_w": None, "row_h": None}
    for i, (rr, cc) in found.items():
        if rr in rmap and cc in cmap:
            meta["img_cells"][i] = (rmap.index(rr), cmap.index(cc))
    col_w = []
    for j in cmap:
        w = None
        for row in grid:
            if j < len(row) and row[j]:
                w = row[j][2] - row[j][0]
                break
        col_w.append(w)
    if col_w and all(col_w):
        meta["col_w"] = [float(w) for w in col_w]
    row_h = []
    for ri in rmap:
        h = None
        if ri < len(grid):
            for c in grid[ri]:
                if c:
                    h = c[3] - c[1]
                    break
        row_h.append(h)
    if row_h and all(row_h):
        meta["row_h"] = [float(h) for h in row_h]
    return t.bbox, rows, meta


def find_tables(page, img_pts=None):
    """Returns [(bbox, rows, meta)]. Pehle bordered tables, phir borderless fallback."""
    found = []
    for t in page.find_tables():
        item = _build_table(t, img_pts)
        if item[1] and len(item[1][0]) >= 2:
            found.append(item)
    if found:
        return found
    try:
        settings = {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "intersection_tolerance": 5,
            "min_words_vertical": 3,
            "min_words_horizontal": 2,
        }
        for t in page.find_tables(settings):
            item = _build_table(t, img_pts)
            if looks_like_table(item[1]):
                found.append(item)
    except Exception:
        pass
    return found


def make_filter(bboxes):
    def outside(obj):
        if obj.get("object_type") != "char":
            return True
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in bboxes)
    return outside


def line_info(ln):
    text = clean_text(ln.get("text", ""))
    if not text:
        return None
    chars = [c for c in ln.get("chars", []) if c.get("text", "").strip()]
    sizes = [c.get("size", 0) for c in chars] or [0]
    size = round(statistics.median(sizes) * 2) / 2
    bold_n = sum(1 for c in chars if re.search(r"bold|black|heavy", c.get("fontname", ""), re.I))
    bold = bool(chars) and bold_n / len(chars) > 0.8
    fonts = Counter(c.get("fontname", "") for c in chars)
    return {
        "text": text, "size": size, "bold": bold,
        "top": float(ln["top"]), "bottom": float(ln["bottom"]),
        "x0": float(ln.get("x0", 0)), "x1": float(ln.get("x1", 0)),
        "font": fonts.most_common(1)[0][0] if fonts else "",
    }


def build_paragraphs(lines, body):
    """Lines ko paragraphs aur headings me convert karta hai."""
    paras, cur = [], None
    for ln in lines:
        ratio = ln["size"] / body if body else 1
        level = 0
        if ratio >= 1.2:
            level = 1 if ratio >= 1.7 else (2 if ratio >= 1.4 else 3)
        split = True
        if cur:
            gap = ln["top"] - cur["bottom"]
            limit = max(ln["size"], cur["size"], 1) * 0.6
            split = bool(
                gap > limit or level or cur["heading"]
                or BULLET_RE.match(ln["text"])
                or abs(ln["size"] - cur["size"]) > 1
                or ln["bold"] != cur["bold"]
            )
        if cur and not split:
            if cur["text"].endswith("-") and ln["text"][:1].islower():
                cur["text"] = cur["text"][:-1] + ln["text"]
            else:
                cur["text"] += " " + ln["text"]
            cur["bottom"] = ln["bottom"]
            cur["x0"] = min(cur["x0"], ln["x0"])
            cur["x1"] = max(cur["x1"], ln["x1"])
            cur["n"] += 1
            cur["last_top"] = ln["top"]
            cur["lines"].append(ln["text"])
        else:
            if cur:
                paras.append(cur)
            cur = {"text": ln["text"], "size": ln["size"], "bold": ln["bold"],
                   "heading": level, "bottom": ln["bottom"], "top": ln["top"],
                   "x0": ln["x0"], "x1": ln["x1"], "font": ln.get("font", ""),
                   "n": 1, "last_top": ln["top"], "lines": [ln["text"]]}
    if cur:
        paras.append(cur)
    out = []
    for p in paras:
        # line ki doori (baseline se baseline): single line ho to uski apni height
        pitch = (p["last_top"] - p["top"]) / (p["n"] - 1) if p["n"] > 1 else (p["bottom"] - p["top"])
        out.append({"type": "para", "text": p["text"], "size": p["size"],
                    "bold": p["bold"], "heading": p["heading"],
                    "x0": p["x0"], "x1": p["x1"], "top": p["top"], "bottom": p["bottom"],
                    "font": p["font"], "nlines": p["n"], "pitch": round(float(pitch), 2),
                    "lines": p["lines"]})
    return out


# ------------------------------------------------------------------
# Image extraction
# ------------------------------------------------------------------
def _names(value):
    """PDF ke filter/colorspace values ko simple string names me badalta hai."""
    value = resolve1(value)
    if value is None:
        return []
    if not isinstance(value, list):
        value = [value]
    out = []
    for x in value:
        x = resolve1(x)
        out.append(getattr(x, "name", None) or (x if isinstance(x, str) else ""))
    return out


def _channels(cs):
    """Colorspace se channel count nikalta hai (3=RGB, 1=Gray). Unknown ho to 0."""
    names = _names(cs)
    if not names:
        return 0
    first = names[0]
    if first in ("DeviceRGB", "RGB", "CalRGB"):
        return 3
    if first in ("DeviceGray", "G", "CalGray"):
        return 1
    if first == "ICCBased":
        try:
            raw = resolve1(cs)
            n = resolve1(resolve1(raw[1]).attrs.get("N"))
            return n if n in (1, 3) else 0
        except Exception:
            return 0
    return 0


def _decode_embedded(img):
    """PDF ke andar saved image ko seedha decode karta hai. Fail ho to None."""
    stream = img.get("stream")
    if stream is None:
        return None
    filters = _names(stream.attrs.get("Filter"))
    # JPEG / JPEG2000 -> bytes seedhe PIL me khul jate hain
    if filters and filters[-1] in ("DCTDecode", "DCT", "JPXDecode"):
        return PILImage.open(io.BytesIO(stream.get_data()))
    # Raw pixels (Flate etc.) -> sirf simple RGB/Gray 8-bit handle karte hain
    unsupported = {"CCITTFaxDecode", "CCF", "JBIG2Decode"}
    if any(f in unsupported for f in filters):
        return None
    w, h = img.get("srcsize") or (0, 0)
    bits = img.get("bits") or 8
    ch = _channels(img.get("colorspace"))
    if not (w and h and ch and bits == 8):
        return None
    data = stream.get_data()
    need = w * h * ch
    if len(data) < need:
        return None
    return PILImage.frombytes("RGB" if ch == 3 else "L", (w, h), data[:need])


def _render_region(page, img):
    """Fallback: PDF page ka wo hissa (image ki jagah) picture bana leta hai."""
    x0 = max(page.bbox[0], img["x0"])
    top = max(page.bbox[1], img["top"])
    x1 = min(page.bbox[2], img["x1"])
    bottom = min(page.bbox[3], img["bottom"])
    if x1 - x0 < 1 or bottom - top < 1:
        return None
    return page.crop((x0, top, x1, bottom)).to_image(resolution=150).original


def _save_image(pil, fid, n):
    """PIL image ko clean PNG bana kar storage me save karta hai. File name return karta hai."""
    if pil.mode not in ("RGB", "RGBA", "L"):
        pil = pil.convert("RGBA" if "A" in pil.mode or "transparency" in pil.info else "RGB")
    if max(pil.size) > MAX_IMG_PX:
        pil.thumbnail((MAX_IMG_PX, MAX_IMG_PX))
    name = f"{fid}_{n}.png"
    pil.save(STORAGE / name, format="PNG", optimize=True)
    return name


def collect_image_candidates(page):
    """Page ki saari kaam ki images (size/position ke sath). Abhi save nahi karta."""
    out = []
    page_area = float(page.width) * float(page.height) or 1.0
    for img in page.images:
        try:
            w = float(img["x1"]) - float(img["x0"])
            h = float(img["bottom"]) - float(img["top"])
            if w < MIN_IMG_PT or h < MIN_IMG_PT:
                continue  # chhoti decoration / lines
            out.append({
                "img": img, "w": w, "h": h,
                "x0": float(img["x0"]), "x1": float(img["x1"]),
                "top": float(img["top"]), "bottom": float(img["bottom"]),
                "cx": (float(img["x0"]) + float(img["x1"])) / 2,
                "cy": (float(img["top"]) + float(img["bottom"])) / 2,
                "full": (w * h) / page_area > 0.9,  # poore page ki image (background)
                "mi": None,
            })
        except Exception:
            continue
    return out


def save_candidate(page, c, fid, st):
    """Image ko PNG bana kar save karta hai. File name ya None. Kabhi crash nahi karta."""
    if st["n"] >= MAX_IMAGES:
        return None
    img = c["img"]
    key = None
    try:
        oid = getattr(img.get("stream"), "objid", None)
        if oid is not None:
            key = oid  # same image dobara aaye (jaise har page ka logo) to wohi file use hogi
    except Exception:
        key = None
    if key is not None and key in st["cache"]:
        return st["cache"][key]
    pil = None
    try:
        pil = _decode_embedded(img)
    except Exception:
        pil = None
    if pil is None:
        try:
            pil = _render_region(page, img)
        except Exception:
            pil = None
    if pil is None:
        return None
    try:
        st["n"] += 1
        fname = _save_image(pil, fid, st["n"])
    except Exception:
        return None
    if key is not None:
        st["cache"][key] = fname
    return fname


def remove_images(fid):
    """Is file id ki saari saved images delete karta hai."""
    if not FID_RE.match(fid or ""):
        return
    for p in STORAGE.glob(f"{fid}_*.png"):
        try:
            p.unlink()
        except OSError:
            pass


# ------------------------------------------------------------------
# Layout: side-by-side blocks (logo + text, image + paragraph, image + image)
# ------------------------------------------------------------------
def group_side_by_side(blocks):
    """Jo blocks PDF me image ke bagal (left/right) me hain unhe ek 'row group' banata hai.
    Returns (blocks, groups). Har grouped block me 'group' (id) aur 'col' (column number) aata hai."""
    n = len(blocks)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, bi in enumerate(blocks):
        if bi["type"] != "image":
            continue
        hi = bi["bottom"] - bi["top"]
        for j, bj in enumerate(blocks):
            if j == i:
                continue
            ov = min(bi["bottom"], bj["bottom"]) - max(bi["top"], bj["top"])
            if ov <= 0 or ov < 0.3 * min(hi, bj["bottom"] - bj["top"]):
                continue  # upar-neeche hain, bagal me nahi
            if bj["x1"] <= bi["x0"] + 2 or bj["x0"] >= bi["x1"] - 2:
                ra, rb = find(i), find(j)
                if ra != rb:
                    parent[rb] = ra

    comps = {}
    for i in range(n):
        comps.setdefault(find(i), []).append(i)

    groups, assigned = [], {}
    for idxs in comps.values():
        if len(idxs) < 2:
            continue
        members = sorted(idxs, key=lambda k: (blocks[k]["x0"], blocks[k]["top"]))
        cols = []
        for k in members:
            b = blocks[k]
            if cols and b["x0"] < cols[-1]["x1"] - 2:
                cols[-1]["items"].append(k)
                cols[-1]["x1"] = max(cols[-1]["x1"], b["x1"])
            else:
                cols.append({"x0": b["x0"], "x1": b["x1"], "items": [k]})
        if len(cols) < 2:
            continue
        gid = len(groups)
        g_cols = []
        for ci, col in enumerate(cols):
            s1 = cols[ci + 1]["x0"] if ci + 1 < len(cols) else col["x1"]
            g_cols.append({"s0": col["x0"], "s1": s1})
            for k in col["items"]:
                assigned[k] = (gid, ci)
        groups.append({"x0": cols[0]["x0"], "x1": max(c["x1"] for c in cols), "cols": g_cols})

    if not groups:
        return blocks, []

    out, done = [], set()
    for i, b in enumerate(blocks):
        if i in assigned:
            gid = assigned[i][0]
            if gid in done:
                continue
            done.add(gid)
            mem = [k for k, (g, _) in assigned.items() if g == gid]
            mem.sort(key=lambda k: (assigned[k][1], blocks[k]["top"]))
            for k in mem:
                blocks[k]["group"] = gid
                blocks[k]["col"] = assigned[k][1]
                out.append(blocks[k])
        else:
            out.append(b)
    return out, groups


# ------------------------------------------------------------------
# Header / Footer detection
# ------------------------------------------------------------------
HF_TOP = 0.20          # page ke upar ka itna hissa header zone
HF_BOT = 0.80          # page ke neeche ka itna hissa footer zone
HF_SINGLE_FOOT = 0.93  # 1-page PDF me footer sirf bilkul neeche ke text ko mante hain


def page_rules(page, bboxes):
    """Page ki chaurri horizontal lines (header/footer ko body se alag karne wali). Tables ki lines ignore."""
    out = []
    pw = float(page.width)
    try:
        cands = list(page.lines) + list(page.rects)
    except Exception:
        return out
    for o in cands:
        try:
            x0, x1 = float(o["x0"]), float(o["x1"])
            top, bot = float(o["top"]), float(o["bottom"])
        except Exception:
            continue
        if bot - top > 2.5 or (x1 - x0) < 0.5 * pw:
            continue
        cx, cy = (x0 + x1) / 2, (top + bot) / 2
        if any(b[0] - 2 <= cx <= b[2] + 2 and b[1] - 2 <= cy <= b[3] + 2 for b in bboxes):
            continue
        out.append(cy)
    return sorted(out)


def _hf_key(b):
    """Block ki pehchan jo page number badalne par bhi same rahe (digits ko # bana dete hain)."""
    t = b["type"]
    if t == "para":
        return ("p", re.sub(r"\d+", "#", re.sub(r"\s+", " ", b["text"].lower())).strip())
    if t == "image":
        return ("i", round(b["width"] / 3.0), round(b["height"] / 3.0))
    txt = " ".join(" ".join(r) for r in b["rows"]).lower()
    return ("t", re.sub(r"\d+", "#", re.sub(r"\s+", " ", txt)).strip())


def _hf_region(b, ph):
    if b["bottom"] <= HF_TOP * ph:
        return "header"
    if b["top"] >= HF_BOT * ph:
        return "footer"
    return None


def _marks_repeated(pages):
    """Multi-page PDF: jo block (text/image) kai pages par same jagah repeat ho wo header/footer hai."""
    n = len(pages)
    need = max(2, math.ceil(n * 0.5))
    occ = {}
    for pi, p in enumerate(pages):
        ph = p["height"]
        for bi, b in enumerate(p["blocks"]):
            reg = _hf_region(b, ph)
            if reg:
                occ.setdefault((reg, _hf_key(b)), []).append((pi, bi, b))
    marked, srcs = {}, {}
    for (reg, key), lst in occ.items():
        anchor = lst[0][2]
        match = [(pi, bi, b) for pi, bi, b in lst
                 if abs(b["top"] - anchor["top"]) <= 10 and abs(b["x0"] - anchor["x0"]) <= 10]
        if len({pi for pi, _, _ in match}) >= need:
            per_page = {}
            for pi, bi, b in match:
                marked[(pi, bi)] = reg
                per_page.setdefault(pi, b)
            srcs[(reg, key)] = per_page
    return marked, srcs


def _marks_single(pages):
    """1-page PDF: repeat ka muqabla nahi ho sakta, is liye horizontal line ya bilkul neeche ke text se."""
    p = pages[0]
    ph = p["height"]
    blocks = p["blocks"]
    rules = p.get("rules", [])
    marked = {}
    for y in sorted(r for r in rules if r <= HF_TOP * ph):
        idx = [bi for bi, b in enumerate(blocks) if b["bottom"] <= y + 2]
        if idx:
            for bi in idx:
                marked[(0, bi)] = "header"
            break
    for y in sorted((r for r in rules if r >= HF_BOT * ph), reverse=True):
        idx = [bi for bi, b in enumerate(blocks) if b["top"] >= y - 2]
        if idx:
            for bi in idx:
                marked[(0, bi)] = "footer"
            break
    for bi, b in enumerate(blocks):
        if (0, bi) not in marked and b["type"] == "para" and b["top"] >= HF_SINGLE_FOOT * ph:
            marked[(0, bi)] = "footer"
    return marked, {}


def _renumber_groups(blocks, groups):
    gmap, new_groups = {}, []
    for b in blocks:
        gid = b.get("group")
        if gid is not None:
            if gid not in gmap:
                gmap[gid] = len(new_groups)
                new_groups.append(groups[gid])
            b["group"] = gmap[gid]
    return new_groups


def _page_fields(b, key, reg, srcs, pages):
    """'Page 1 of 2' jaisi text me page number ko Word/Excel ke PAGE / NUMPAGES field me badalta hai."""
    per_page = srcs.get((reg, key))
    n = len(pages)
    if not per_page or len(per_page) != n:
        return None
    nums = []
    for pi in range(n):
        nums.append(re.findall(r"\d+", per_page[pi]["text"]))
    k = len(nums[0])
    if k == 0 or any(len(x) != k for x in nums):
        return None
    kinds = []
    for t in range(k):
        vals = [int(nums[pi][t]) for pi in range(n)]
        if vals == list(range(1, n + 1)):
            kinds.append("PAGE")
        elif all(v == n for v in vals):
            kinds.append("NUMPAGES")
        else:
            kinds.append(None)
    if not any(kinds):
        return None
    parts = re.split(r"\d+", b["text"])
    seg = []
    for t in range(k):
        if parts[t]:
            seg.append(["t", parts[t]])
        seg.append([kinds[t], ""] if kinds[t] else ["t", re.findall(r"\d+", b["text"])[t]])
    if parts[k]:
        seg.append(["t", parts[k]])
    return seg


def split_header_footer(pages):
    """Header/footer blocks ko pages ki body se nikal kar alag karta hai.
    Returns (header, footer) -- har ek {'blocks','groups','width'} ya None."""
    if not pages:
        return None, None
    try:
        marked, srcs = _marks_repeated(pages) if len(pages) >= 2 else _marks_single(pages)
    except Exception:
        return None, None
    # group ke saare members ek hi region me hon, warna group ko body me rehne do
    for pi, p in enumerate(pages):
        by_gid = {}
        for bi, b in enumerate(p["blocks"]):
            if b.get("group") is not None:
                by_gid.setdefault(b["group"], []).append(bi)
        for gid, idxs in by_gid.items():
            regs = {marked.get((pi, bi)) for bi in idxs}
            if len(regs) != 1 or None in regs:
                for bi in idxs:
                    marked.pop((pi, bi), None)
    if not marked:
        return None, None

    def take(region):
        pis = [pi for (pi, _), r in marked.items() if r == region]
        if not pis:
            return None
        p0 = pages[min(pis)]
        idxs = sorted(bi for (pi, bi), r in marked.items() if pi == min(pis) and r == region)
        blocks = [copy.deepcopy(p0["blocks"][bi]) for bi in idxs]
        groups = _renumber_groups(blocks, copy.deepcopy(p0["groups"]))
        for b in blocks:
            if b["type"] == "para":
                seg = _page_fields(b, _hf_key(b), region, srcs, pages)
                if seg:
                    b["segments"] = seg
        return {"blocks": blocks, "groups": groups, "width": p0["width"]}

    header, footer = take("header"), take("footer")
    for pi, p in enumerate(pages):
        drop = {bi for (pj, bi) in marked if pj == pi}
        if drop:
            p["blocks"] = [b for bi, b in enumerate(p["blocks"]) if bi not in drop]
            p["groups"] = _renumber_groups(p["blocks"], p["groups"])
    return header, footer


# ------------------------------------------------------------------
# PDF -> structured data
# ------------------------------------------------------------------
def extract_pdf(path, original_name, fid=None):
    raw_pages, sizes = [], Counter()
    img_state = {"n": 0, "cache": {}}
    with pdfplumber.open(path) as pdf:
        for idx, page in enumerate(pdf.pages, 1):
            items = []
            # images ki list pehle (tables ke andar wali images pehchanne ke liye)
            cands = collect_image_candidates(page) if fid else []
            mappable = [c for c in cands if not c["full"]]
            for mi, c in enumerate(mappable):
                c["mi"] = mi
            img_pts = [(c["cx"], c["cy"]) for c in mappable]
            try:
                tables = find_tables(page, img_pts)
            except Exception:
                tables = []
            bboxes = [t[0] for t in tables]
            try:
                src = page.filter(make_filter(bboxes)) if bboxes else page
                raw_lines = src.extract_text_lines(strip=True, return_chars=True)
            except Exception:
                raw_lines = []
            for ln in raw_lines:
                d = line_info(ln)
                if d:
                    items.append((d["top"], "line", d))
                    sizes[d["size"]] += len(d["text"])
            used = set()
            for bbox, rows, meta in tables:
                cell_imgs = []
                for mi, (rr, cc) in meta["img_cells"].items():
                    if mi in used:
                        continue
                    used.add(mi)
                    c = mappable[mi]
                    fname = save_candidate(page, c, fid, img_state)
                    if fname:
                        cell_imgs.append({"r": rr, "c": cc, "file": fname,
                                          "width": c["w"], "height": c["h"]})
                fsize, ffont = None, ""
                try:  # table ke text ka size/font
                    tc = [c for c in page.chars if c.get("text", "").strip()
                          and bbox[0] <= (c["x0"] + c["x1"]) / 2 <= bbox[2]
                          and bbox[1] <= (c["top"] + c["bottom"]) / 2 <= bbox[3]]
                    if tc:
                        fsize = round(statistics.median(c.get("size", 0) for c in tc) * 2) / 2
                        ffont = Counter(c.get("fontname", "") for c in tc).most_common(1)[0][0]
                except Exception:
                    pass
                items.append((bbox[1], "table", {
                    "type": "table", "rows": rows,
                    "x0": float(bbox[0]), "x1": float(bbox[2]),
                    "top": float(bbox[1]), "bottom": float(bbox[3]),
                    "col_w": meta["col_w"], "row_h": meta["row_h"],
                    "font_size": fsize, "font": ffont, "cell_images": cell_imgs}))
            # baaki images (logo etc.) jo kisi table ke andar nahi hain
            has_text = any(k == "line" for _, k, _ in items)
            pw = float(page.width)
            for c in cands:
                if c["mi"] is not None and c["mi"] in used:
                    continue
                if c["full"] and has_text:
                    continue  # poore page ka background image (text ke peeche) skip
                fname = save_candidate(page, c, fid, img_state)
                if not fname:
                    continue
                align = "left" if c["cx"] < pw * 0.38 else ("right" if c["cx"] > pw * 0.62 else "center")
                items.append((c["top"], "image", {
                    "type": "image", "file": fname, "width": c["w"], "height": c["h"],
                    "align": align, "x0": c["x0"], "x1": c["x1"],
                    "top": c["top"], "bottom": c["bottom"]}))
            items.sort(key=lambda x: x[0])
            try:
                rules = page_rules(page, bboxes)
            except Exception:
                rules = []
            raw_pages.append({"number": idx, "items": items, "rules": rules,
                              "width": pw, "height": float(page.height)})
            try:
                page.flush_cache()
            except Exception:
                pass

    body = sizes.most_common(1)[0][0] if sizes else 11
    pages = []
    for pg in raw_pages:
        flow, buf, pics = [], [], []
        for _, kind, payload in pg["items"]:
            if kind == "line":
                buf.append(payload)
            elif kind == "table":
                flow += build_paragraphs(buf, body)
                buf = []
                flow.append(payload)
            else:  # image (text ko tod'ti nahi, alag se position ke hisaab se lagti hai)
                pics.append(payload)
        flow += build_paragraphs(buf, body)
        blocks = flow + pics
        blocks.sort(key=lambda b: b["top"])
        try:
            blocks, groups = group_side_by_side(blocks)
        except Exception:
            groups = []
            for b in blocks:
                b.pop("group", None)
                b.pop("col", None)
        pages.append({"number": pg["number"], "width": pg["width"], "height": pg["height"],
                      "blocks": blocks, "groups": groups, "rules": pg["rules"]})

    # header / footer ko body se alag karo (Word/Excel ke real header-footer section ke liye)
    header, footer = split_header_footer(pages)
    for p in pages:
        p.pop("rules", None)
    extra = [r["blocks"] for r in (header, footer) if r]

    all_blocks = [b for p in pages for b in p["blocks"]] + [b for bl in extra for b in bl]
    paras = sum(1 for b in all_blocks if b["type"] == "para")
    tbls = sum(1 for b in all_blocks if b["type"] == "table")
    imgs = sum(1 for b in all_blocks if b["type"] == "image")
    imgs += sum(len(b.get("cell_images", [])) for b in all_blocks if b["type"] == "table")
    words = sum(len(b["text"].split()) for b in all_blocks if b["type"] == "para")
    return {
        "filename": original_name,
        "header": header,
        "footer": footer,
        "pages": pages,
        "stats": {"pages": len(pages), "paragraphs": paras, "tables": tbls,
                  "words": words, "images": imgs},
    }


# ------------------------------------------------------------------
# Word builder
# ------------------------------------------------------------------
ALIGN_MAP = {"left": WD_ALIGN_PARAGRAPH.LEFT,
             "center": WD_ALIGN_PARAGRAPH.CENTER,
             "right": WD_ALIGN_PARAGRAPH.RIGHT}
_ctx = threading.local()  
def _usable():
    return getattr(_ctx, "usable_in", WORD_USABLE_IN)


def _left_pt():
    return getattr(_ctx, "left_pt", 0.0)


def _right_edge():
    return getattr(_ctx, "right_edge_pt", None)


def map_font(name):
    """PDF ke font ka naam -> Word ka font (Helvetica -> Arial, wagera). Na samjhe to None (default font)."""
    n = re.sub(r"^[A-Z]{6}\+", "", name or "")
    base = re.split(r"[-,]", n)[0].strip()
    low = base.lower()
    if low in ("helvetica", "arial", "arialmt"):
        return "Arial"
    if low.startswith("times"):
        return "Times New Roman"
    if low.startswith("courier"):
        return "Courier New"
    if re.fullmatch(r"[A-Za-z][A-Za-z ]{2,28}", base):
        return base
    return None


def slab_align(b, s0, s1):
    """Block apne column ke andar left/center/right kahan hai (PDF position ke hisaab se)."""
    left = b["x0"] - s0
    right = s1 - b["x1"]
    if right <= 4 and left > 12:
        return "right"
    if left > 12 and right > 12 and abs(left - right) <= 8:
        return "center"
    return "left"


def set_table_borders(table, sz=4, color="000000"):
    """Table ki lines seedha XML me likhta hai (style par depend nahi) taake har viewer me dikhein."""
    tblPr = table._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar",
                                    "w:tblLook", "w:tblCaption", "w:tblDescription")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
    for row in table.rows:  # har cell par bhi (kuch viewers sirf cell borders maante hain)
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            cb = tcPr.find(qn("w:tcBorders"))
            if cb is None:
                cb = OxmlElement("w:tcBorders")
                tcPr.insert_element_before(cb, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                                           "w:tcFitText", "w:vAlign", "w:hideMark")
            for edge in ("top", "left", "bottom", "right"):
                el = cb.find(qn("w:" + edge))
                if el is None:
                    el = OxmlElement("w:" + edge)
                    cb.append(el)
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(sz))
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), color)


def _set_tbl_indent(table, pt):
    """Table ko margin se pt points andar se shuru karta hai (PDF jaisi x-position)."""
    if pt < 0.5:
        return
    tblPr = table._tbl.tblPr
    el = tblPr.find(qn("w:tblInd"))
    if el is None:
        el = OxmlElement("w:tblInd")
        tblPr.insert_element_before(el, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
                                    "w:tblLook", "w:tblCaption", "w:tblDescription")
    el.set(qn("w:w"), str(int(pt * 20)))
    el.set(qn("w:type"), "dxa")


def _zero_cell_margins(table):
    """Layout table me cell ke andar ka gap 0 (taake content bilkul PDF wali jagah par aaye)."""
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + side)
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.insert_element_before(mar, "w:tblLook", "w:tblCaption", "w:tblDescription")


def _mark_size(p, pt):
    """Paragraph ke end-mark ka font size (khali/image paragraph ki height kam rakhne ke liye)."""
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(round(pt * 2))))
        rPr.append(el)


def _fmt_par(p, line_pt=None, page_break=False):
    """Paragraph ki spacing: upar-neeche gap 0, line height PDF jaisi (exact), widow control band."""
    f = p.paragraph_format
    f.space_before = Pt(0)
    f.space_after = Pt(0)
    f.widow_control = False
    f.keep_with_next = False
    f.keep_together = False
    if line_pt:
        f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        f.line_spacing = Pt(line_pt)
    else:
        f.line_spacing = 1.0
    if page_break:
        f.page_break_before = True


def _new_table(container, rows, cols):
    """Document / cell / header-footer teeno me table banane ka sahi tareeqa."""
    if isinstance(container, (_Header, _Footer)):
        return container.add_table(rows, cols, Inches(_usable()))  # header-footer me width zaroori hai
    return container.add_table(rows=rows, cols=cols)


def _add_field(par, kind, fmt):
    """Page number (PAGE) ya total pages (NUMPAGES) ka live field."""
    runs = []
    for step in ("begin", "instr", "separate", "text", "end"):
        r = par.add_run()
        if step == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = " %s " % kind
            r._r.append(el)
        elif step == "text":
            r.text = "1"
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), step)
            r._r.append(el)
        runs.append(r)
    for r in runs:
        fmt(r)


def _word_par(container, state):
    """Cell ka pehla khali paragraph reuse karta hai, warna naya banata hai."""
    if state.get("fresh"):
        state["fresh"] = False
        return container.paragraphs[0]
    return container.add_paragraph()


def _spacer(container, state, height, page_break=False):
    """Khali paragraph jis ki height exact 'height' points hoti hai (PDF ke blocks ke beech ka gap)."""
    p = _word_par(container, state)
    _fmt_par(p, min(1500.0, max(1.0, height)), page_break)
    _mark_size(p, 1)
    return p


def add_docx_image(container, b, state=None, max_in=None, align=None):
    """Image lagata hai. Returns Word me uski height (points)."""
    path = STORAGE / Path(b["file"]).name
    if not path.exists():
        return 0.0
    max_in = max_in or _usable()
    nat = max(b["width"] / 72.0, 0.01)
    w_in = min(max(nat, 0.2), max_in)
    scale = w_in / nat
    p = _word_par(container, state if state is not None else {})
    _fmt_par(p)
    _mark_size(p, 1)
    p.alignment = ALIGN_MAP.get(align or b.get("align"), WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run()
    run.font.size = Pt(1)
    run.add_picture(str(path), width=Inches(w_in))
    return b["height"] * scale + 1.5


def add_docx_table(container, b, state, max_in, in_cell=False):
    """Table banata hai (PDF jaisi width, row height, font). Returns Word me uski height (points)."""
    rows = b["rows"]
    ncols = max(len(r) for r in rows)
    fresh_par = container.paragraphs[0] if state.get("fresh") else None
    table = _new_table(container, len(rows), ncols)
    if fresh_par is not None:  # cell/header ke upar ka khali paragraph hata do
        state["fresh"] = False
        fresh_par._p.getparent().remove(fresh_par._p)
    table.style = "Table Grid"
    set_table_borders(table)
    if not in_cell and "x0" in b:
        _set_tbl_indent(table, b["x0"] - _left_pt())
    col_w = b.get("col_w")
    widths_in = None
    if col_w and len(col_w) == ncols and sum(col_w) > 0:
        tw = min(max_in, max(0.5, (b["x1"] - b["x0"]) / 72.0))
        widths_in = [w / sum(col_w) * tw for w in col_w]
        table.autofit = False
        for ci, w in enumerate(widths_in):
            table.columns[ci].width = Inches(w)
            for ri in range(len(rows)):
                table.cell(ri, ci).width = Inches(w)
    fsize = b.get("font_size")
    fam = map_font(b.get("font"))
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = table.cell(ri, ci)
            cell.text = row[ci] if ci < len(row) else ""
            for par in cell.paragraphs:
                _fmt_par(par)
                for r in par.runs:
                    if fsize:
                        r.font.size = Pt(fsize)
                    if fam:
                        r.font.name = fam
                    if ri == 0 and len(rows) > 1:
                        r.bold = True
    row_h = b.get("row_h")
    if row_h and len(row_h) == len(rows):
        for ri, h in enumerate(row_h):
            table.rows[ri].height = Pt(h)
            table.rows[ri].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    for im in b.get("cell_images", []):  
        try:
            path = STORAGE / Path(im["file"]).name
            if not path.exists():
                continue
            cell = table.cell(im["r"], im["c"])
            cap = (widths_in[im["c"]] - 0.2) if widths_in else 1.5
            w_in = min(max(im["width"] / 72.0, 0.2), max(0.3, cap))
            par = cell.add_paragraph() if cell.text.strip() else cell.paragraphs[0]
            _fmt_par(par)
            par.add_run().add_picture(str(path), width=Inches(w_in))
        except Exception:
            pass
    if in_cell:  # cell ke andar table ke baad wala zaroori paragraph chhota kar do
        tail = container.paragraphs[-1]
        _fmt_par(tail, 1.0)
        _mark_size(tail, 1)
    return (b.get("bottom", 0) - b.get("top", 0)) + 0.5 * len(rows)


def add_docx_block(container, b, state, max_in, slab=None, in_cell=False, hf=False):
    """Ek block (paragraph / image / table) likhta hai. Returns Word me uski height (points)."""
    t = b["type"]
    if t == "para":
        p = _word_par(container, state)
        size = max(5.0, min(float(b.get("size") or 11), 72.0))
        fam = map_font(b.get("font"))
        heading = bool(b["heading"]) and not hf  # header/footer me heading style nahi (navigation/TOC me aa jata hai)
        if heading:
            p.style = "Heading %d" % min(b["heading"], 3)

        def fmt(run):
            run.font.size = Pt(size)
            if fam:
                run.font.name = fam
            if hf:
                if b["bold"] or b["heading"]:
                    run.bold = True
            elif not heading:
                run.bold = b["bold"]

        if b.get("segments"):
            for kind, txt in b["segments"]:
                if kind == "t":
                    fmt(p.add_run(txt))
                else:
                    _add_field(p, kind, fmt)
        else:
            fmt(p.add_run(b["text"]))
        pitch = b.get("pitch")
        line = max(float(pitch), size) if pitch else None
        _fmt_par(p, line)
        if slab:
            p.alignment = ALIGN_MAP[slab_align(b, slab[0], slab[1])]
        elif _right_edge() and "x0" in b:
            p.alignment = ALIGN_MAP[slab_align(b, _left_pt(), _right_edge())]
        if line:
            return b.get("nlines", 1) * line
        return b.get("bottom", 0) - b.get("top", 0)
    if t == "image":
        try:
            al = slab_align(b, slab[0], slab[1]) if slab else None
            return add_docx_image(container, b, state, max_in, al)
        except Exception:
            return 0.0  # ek image kharab ho to baaki document na ruke
    return add_docx_table(container, b, state, max_in, in_cell)


def docx_group_widths(cols, members, width_in):
    """Group ke columns ki width (inch). Image wale column me image fit hogi, baaki text ko milega."""
    n = len(cols)
    span = [max(1.0, c["s1"] - c["s0"]) for c in cols]
    widths = [max(0.4, sp / sum(span) * width_in) for sp in span]
    nat = [0.0] * n
    has_text = [False] * n
    for b in members:
        ci = b.get("col", 0)
        if b["type"] == "image":
            nat[ci] = max(nat[ci], b["width"] / 72.0)
        else:
            has_text[ci] = True
    pad = 0.05
    for ci in range(n):
        if nat[ci]:
            widths[ci] = max(widths[ci], nat[ci] + pad)
    excess = sum(widths) - width_in
    if excess > 0:  # pehle text columns se kam karo (kam az kam 1 inch rakho)
        for ci in sorted([k for k in range(n) if has_text[k]], key=lambda k: -widths[k]):
            cut = min(excess, max(0.0, widths[ci] - 1.0))
            widths[ci] -= cut
            excess -= cut
            if excess <= 0:
                break
    if excess > 0.01:  # ab bhi jagah kam ho (sirf images) to images ko barabar chhota karo
        img_cols = [k for k in range(n) if nat[k]]
        other = sum(widths[k] for k in range(n) if k not in img_cols)
        avail = width_in - other - pad * len(img_cols)
        tot = sum(nat[k] for k in img_cols)
        if avail > 0.5 and tot > 0:
            f = min(1.0, avail / tot)
            for k in img_cols:
                widths[k] = nat[k] * f + pad
    return widths


def add_docx_group(doc, members, g, page_w=None, state=None, hf=False):
    """Side-by-side blocks ke liye bina lines wali layout table (jaise PDF me left logo / right text).
    Returns Word me group ka neeche wala y (PDF points me)."""
    cols = g["cols"]
    right = g["x1"]
    edge = _right_edge()
    if any(b["type"] != "image" and b.get("col") == len(cols) - 1 for b in members):
        if edge:
            right = max(right, edge)  # text wale column ko right margin tak failne do
        elif page_w:
            right = max(right, page_w - g["x0"])
    width_in = min(_usable(), max(1.0, (right - g["x0"]) / 72.0))
    cols_eff = [dict(c) for c in cols]
    cols_eff[-1]["s1"] = max(cols_eff[-1]["s1"], right)
    widths = docx_group_widths(cols_eff, members, width_in)
    fresh_par = doc.paragraphs[0] if (state and state.get("fresh")) else None
    tbl = _new_table(doc, 1, len(cols))
    if fresh_par is not None:  # header/footer ka upar wala khali paragraph hata do
        state["fresh"] = False
        fresh_par._p.getparent().remove(fresh_par._p)
    tbl.autofit = False
    _set_tbl_indent(tbl, g["x0"] - _left_pt())
    _zero_cell_margins(tbl)
    for ci, w in enumerate(widths):
        tbl.columns[ci].width = Inches(w)
        tbl.cell(0, ci).width = Inches(w)
    top = min(b.get("top", 0.0) for b in members)
    bottoms = [top]
    for ci, c in enumerate(cols):
        items = [b for b in members if b.get("col") == ci]
        bottoms.append(write_blocks(tbl.cell(0, ci), items, [], top, {"fresh": True}, hf=hf,
                                    max_in=max(0.3, widths[ci] - 0.02), in_cell=True,
                                    slab=(c["s0"], c["s1"]))[0])
    return max(bottoms)


def write_blocks(container, blocks, groups, start_y, state, hf=False, max_in=None,
                 page_w=None, page_break=False, in_cell=False, slab=None, end_para=False):
    """Blocks ko PDF ki tarah likhta hai: har block se pehle PDF wala gap (exact-height khali paragraph).
    Returns (neeche wala y, aakhri block table thi ya nahi)."""
    max_in = max_in or _usable()
    prev_bottom = start_y
    prev_tbl = False
    first = True
    i = 0
    while i < len(blocks):
        b = blocks[i]
        gid = b.get("group")
        is_group = gid is not None and gid < len(groups) and not in_cell
        if is_group:
            j = i
            members = []
            while j < len(blocks) and blocks[j].get("group") == gid:
                members.append(blocks[j])
                j += 1
            top = min(m.get("top", prev_bottom) for m in members)
            is_tbl = True
        else:
            top = b.get("top", prev_bottom)
            is_tbl = b["type"] == "table"
        gap = top - prev_bottom
        brk = page_break and first
        if brk or gap >= 1 or (prev_tbl and is_tbl):
            _spacer(container, state, gap if gap >= 1 else 1.0, brk)
        if is_group:
            bottom = add_docx_group(container, members, groups[gid], page_w, state, hf)
            i = j
        else:
            bottom = top + add_docx_block(container, b, state, max_in, slab, in_cell, hf)
            i += 1
        prev_bottom, prev_tbl, first = bottom, is_tbl, False
    if end_para and prev_tbl:  # header/footer paragraph par hi khatam hona chahiye
        _spacer(container, state, 1.0)
    return prev_bottom, prev_tbl


def fill_docx_hf(container, region):
    """Word ke real header ya footer section me region ke blocks likhta hai."""
    container.is_linked_to_previous = False
    blocks = region["blocks"]
    start = min((b["top"] for b in blocks if "top" in b), default=0.0)
    write_blocks(container, blocks, region.get("groups", []), start, {"fresh": True}, hf=True,
                 page_w=region.get("width"), end_para=True)


def build_docx(data):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    sec = doc.sections[0]

    pages = data["pages"]
    hdr, ftr = data.get("header"), data.get("footer")
    hf_blocks = (hdr["blocks"] if hdr else []) + (ftr["blocks"] if ftr else [])
    body_blocks = [b for p in pages for b in p["blocks"]]
    geo = [b for b in body_blocks + hf_blocks if "x0" in b and "top" in b]
    p0 = pages[0] if pages else {}
    top_pt = 36.0
    if geo and p0.get("width") and p0.get("height"):
        # Word page ko PDF page jaisa banao (size, margins) taake har PDF page = ek Word page
        pw, ph = float(p0["width"]), float(p0["height"])
        sec.page_width, sec.page_height = Pt(pw), Pt(ph)
        left = min(max(18.0, min(b["x0"] for b in geo)), pw / 3)
        right = min(max(18.0, pw - max(b["x1"] for b in geo)), pw / 3)
        tops = [p["blocks"][0]["top"] for p in pages if p["blocks"] and "top" in p["blocks"][0]]
        top_pt = max(18.0, min(tops)) if tops else 36.0
        # neeche ka margin: sirf footer ke liye jagah (page break hum khud lagate hain, is liye body area
        # bara rakhna safe hai; thoda sa drift hone par bhi content agle page par nahi jayega)
        bottom_m = 18.0
        if ftr and ftr["blocks"]:
            fb = [b for b in ftr["blocks"] if "bottom" in b and "top" in b]
            if fb:
                sec.footer_distance = Pt(max(8.0, ph - max(b["bottom"] for b in fb)))
                bottom_m = min(ph / 3, max(18.0, ph - min(b["top"] for b in fb) + 2))
        if hdr and hdr["blocks"]:
            ht = [b["top"] for b in hdr["blocks"] if "top" in b]
            if ht:
                sec.header_distance = Pt(max(8.0, min(ht)))
        sec.left_margin, sec.right_margin = Pt(left), Pt(right)
        sec.top_margin, sec.bottom_margin = Pt(top_pt), Pt(bottom_m)
        _ctx.usable_in = max(1.0, (pw - left - right) / 72.0)
        _ctx.left_pt = left
        _ctx.right_edge_pt = pw - right
        # Word ke default gap (har paragraph ke baad 10pt, 1.15 line) band: PDF jaisa spacing khud lagate hain
        pf = normal.paragraph_format
        pf.space_before, pf.space_after = Pt(0), Pt(0)
        pf.line_spacing = 1.0
        pf.widow_control = False
    else:
        for s in doc.sections:
            s.left_margin = s.right_margin = Inches(0.9)
            s.top_margin = s.bottom_margin = Inches(0.8)
        _ctx.usable_in = WORD_USABLE_IN
        _ctx.left_pt = 0.0
        _ctx.right_edge_pt = None

    # header / footer: Word ke asli header-footer section me (body me nahi)
    for key, attr in (("header", "header"), ("footer", "footer")):
        region = data.get(key)
        if region and region.get("blocks"):
            try:
                fill_docx_hf(getattr(sec, attr), region)
            except Exception:
                pass  # header/footer fail ho to baaki document phir bhi ban jaye

    first = True
    last_tbl = False
    for page in pages:
        if not page["blocks"]:
            continue
        _, last_tbl = write_blocks(doc, page["blocks"], page.get("groups", []), top_pt, {},
                                   page_w=page.get("width"), page_break=not first)
        first = False
    if first:
        doc.add_paragraph("(No content found)")
    elif last_tbl:  # document table par khatam ho to aakhir me chhota paragraph
        _spacer(doc, {}, 1.0)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ------------------------------------------------------------------
# Excel builder
# ------------------------------------------------------------------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="4F46E5")
# ---- Excel settings ----
# "sheet" : (DEFAULT) header/footer har PDF page par sheet ke cells me likha jata hai, bilkul PDF ki tarah.
#           Excel, WPS, Google Sheets, mobile - har jagah ek jaisa dikhta hai. Footer me sahi "Page 1 of 2".
# "real"  : asli Excel Header/Footer (sirf print / Page Layout view me dikhta hai, viewers me farq kar sakta hai)
XLSX_HF_MODE = "sheet"
XLSX_VIEW = "pageLayout"     # "pageLayout" ya "normal" (sirf "real" mode me lagta hai)
XLSX_TABLE_SHEETS = True     # False: sirf "Full Document" sheet (har table ki alag sheet nahi banegi)
SHOW_PAGE_MARKERS = False    # True: sheet me "--- Page 1 ---" wali lines dikhao
INT_RE = re.compile(r"-?(0|[1-9]\d{0,14})")
FLOAT_RE = re.compile(r"-?(0|[1-9]\d{0,14})\.\d+")
COMMA_RE = re.compile(r"-?\d{1,3}(,\d{3})+(\.\d+)?")


def to_value(s):
    t = s.strip()
    if COMMA_RE.fullmatch(t):
        n = t.replace(",", "")
        return float(n) if "." in n else int(n)
    if INT_RE.fullmatch(t):
        return int(t)
    if FLOAT_RE.fullmatch(t):
        return float(t)
    return s


def num_format(raw):
    """Agar PDF me number comma ke sath tha (1,540) to Excel me bhi wahi format."""
    t = (raw or "").strip()
    if COMMA_RE.fullmatch(t):
        dec = len(t.split(".")[1]) if "." in t else 0
        return "#,##0" + ("." + "0" * dec if dec else "")
    return None


def set_cell(ws, r, c, val):
    cell = ws.cell(row=r, column=c)
    cell.value = val
    if isinstance(val, str) and val.startswith("="):
        cell.data_type = "s"  # formula ki tarah execute na ho
    return cell


def add_cell_image(ws, row, col, im):
    """Table ke cell (row, col) ke andar image lagata hai; cell ke size me fit karke."""
    path = STORAGE / Path(im["file"]).name
    if not path.exists():
        return
    letter = get_column_letter(col)
    col_chars = ws.column_dimensions[letter].width or 8.43
    avail_w = max(20.0, col_chars * 7 + 5 - 6)
    w_px = im["width"] * 96.0 / 72.0
    h_px = im["height"] * 96.0 / 72.0
    k = min(1.0, avail_w / w_px, 150.0 / h_px)
    w_px, h_px = max(1, int(w_px * k)), max(1, int(h_px * k))
    img = XLImage(io.BytesIO(path.read_bytes()))
    img.width, img.height = w_px, h_px
    ws.add_image(img, f"{letter}{row}")
    need_pt = h_px * 0.75 + 4
    cur = ws.row_dimensions[row].height or 15
    ws.row_dimensions[row].height = min(409, max(cur, need_pt))


def write_table(ws, start_row, rows, start_col=1, cell_images=None):
    ncols = max(len(r) for r in rows)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            raw = row[ci] if ci < len(row) else ""
            cell = set_cell(ws, start_row + ri, start_col + ci, to_value(raw))
            fmt = num_format(raw)
            if fmt:
                cell.number_format = fmt
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if ri == 0 and len(rows) > 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = HEAD_FILL
    for im in cell_images or []:
        try:
            add_cell_image(ws, start_row + im["r"], start_col + im["c"], im)
        except Exception:
            pass
    return start_row + len(rows)


def table_widths(rows, base=14, cap=45):
    ncols = max(len(r) for r in rows)
    w = [base] * ncols
    for r in rows:
        for i, c in enumerate(r):
            longest = max((len(x) for x in c.split("\n")), default=0)
            w[i] = min(cap, max(w[i], longest + 2))
    return w


def x_to_col(col_px, x):
    """Sheet ke pixel x se (column index, column ke andar offset) nikalta hai."""
    acc = 0.0
    for i, cw in enumerate(col_px):
        if x < acc + cw:
            return i, max(0, int(x - acc))
        acc += cw
    return len(col_px) - 1, max(0, int(col_px[-1] - 1))


def add_xlsx_image(ws, r, b, max_px, col_px=None, x_px=None, k=96.0 / 72.0):
    """Image ko row r par lagata hai (x_px diya ho to PDF wali horizontal position par).
    Returns next free row."""
    path = STORAGE / Path(b["file"]).name
    if not path.exists():
        return r
    w_px = b["width"] * k
    h_px = b["height"] * k
    if w_px > max_px:  # sheet ki width se bahar na jaye
        h_px *= max_px / w_px
        w_px = max_px
    img = XLImage(io.BytesIO(path.read_bytes()))
    img.width, img.height = max(1, int(w_px)), max(1, int(h_px))
    if col_px and x_px is not None:
        x_px = min(max(0.0, x_px), max(0.0, sum(col_px) - img.width))
        ci, off = x_to_col(col_px, x_px)
        marker = AnchorMarker(col=ci, colOff=pixels_to_EMU(off), row=r - 1, rowOff=0)
        img.anchor = OneCellAnchor(
            _from=marker,
            ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
        ws.add_image(img)
    else:
        ws.add_image(img, f"A{r}")
    return r + math.ceil(h_px / 20.0) + 1  # default row height ~20px


def write_xlsx_para(ws, r, b, c1, c2, chars, halign="left"):
    """Paragraph ko columns c1..c2 (1-based) me merge karke likhta hai."""
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    size = b["size"] or 11
    cell = set_cell(ws, r, c1, b["text"])
    cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal=halign)
    cell.font = Font(bold=bool(b["heading"] or b["bold"]),
                     size=max(10, min(size, 24)))
    lines = max(1, math.ceil(len(b["text"]) * (size / 11) / max(1, chars)))
    ws.row_dimensions[r].height = min(409, 15 * lines * max(1, size / 11))


def write_xlsx_group(ws, r0, members, g, col_px, widths, pscale, kscale, max_px, page_w=None):
    """Side-by-side blocks: har column apni sheet-columns me, sab ek hi row se shuru."""
    n_sheet = len(col_px)
    starts, ends, prev_end = [], [], -1
    last = len(g["cols"]) - 1
    text_last = any(b["type"] != "image" and b.get("col") == last for b in members)
    for k, c in enumerate(g["cols"]):
        s1 = c["s1"]
        if k == last and text_last and page_w:
            s1 = max(s1, page_w - g["x0"])  # text wala last column right margin tak failne do
        si = x_to_col(col_px, c["s0"] * pscale)[0]
        ei = x_to_col(col_px, max(c["s0"], s1) * pscale - 1)[0]
        si = min(max(si, prev_end + 1), n_sheet - 1)
        ei = max(ei, si)
        starts.append(si)
        ends.append(ei)
        prev_end = ei
    end_row = r0
    for ci, c in enumerate(g["cols"]):
        cur = r0
        chars = sum(widths[k] for k in range(starts[ci], ends[ci] + 1))
        for b in members:
            if b.get("col") != ci:
                continue
            if b["type"] == "para":
                write_xlsx_para(ws, cur, b, starts[ci] + 1, ends[ci] + 1, chars,
                                slab_align(b, c["s0"], c["s1"]))
                cur += 1
            elif b["type"] == "image":
                cur = add_xlsx_image(ws, cur, b, max_px, col_px, b["x0"] * pscale, kscale)
            else:
                cur = write_table(ws, cur, b["rows"], starts[ci] + 1, b.get("cell_images")) + 1
        end_row = max(end_row, cur)
    return end_row


def _hf_esc(t):
    return t.replace("&", "&&")  # header/footer me & ka matlab code hota hai


def _hf_block_text(b, fmt=True):
    """Ek block ko Excel header/footer ki string me badalta hai (PAGE/NUMPAGES -> &P/&N)."""
    if b["type"] == "table":
        raw = "  ".join(" ".join(c for c in r if c) for r in b["rows"])
        body, bold, size = _hf_esc(raw), False, 9
    else:
        if b.get("segments"):
            body = "".join(_hf_esc(t) if k == "t" else ("&P" if k == "PAGE" else "&N")
                           for k, t in b["segments"])
        elif len(b.get("lines") or []) > 1:   # PDF ki alag alag lines Excel header me bhi alag lines
            body = "\n".join(_hf_esc(x) for x in b["lines"])
        else:
            body = _hf_esc(b["text"])
        bold, size = bool(b["bold"] or b["heading"]), int(max(8, min(b["size"] or 9, 18)))
    if not fmt:
        return body
    code = ('&"-,Bold"' if bold else '&"-,Regular"') + "&%d" % size
    if body[:1].isdigit():
        body = " " + body  # size code ke baad digit se text shuru ho to size ka hissa na ban jaye
    return code + body


def xlsx_hf_slots(region):
    """Region (header/footer) ko Left/Center/Right 3 slots me baantta hai (PDF ki x-position se)."""
    pw = float(region.get("width") or 595.0)
    slots = [{"img": None, "blocks": []} for _ in range(3)]

    def slot_of(b):
        cx = (b["x0"] + b["x1"]) / 2.0 / pw
        return 0 if cx < 0.34 else (1 if cx < 0.66 else 2)

    for b in sorted([b for b in region["blocks"] if b["type"] == "image"], key=lambda b: b["x0"]):
        sl = slot_of(b)
        while sl < 3 and slots[sl]["img"] is not None:
            sl += 1  # ek slot me sirf ek image ho sakti hai
        if sl < 3:
            slots[sl]["img"] = b
    for b in sorted([b for b in region["blocks"] if b["type"] != "image"], key=lambda b: (b["top"], b["x0"])):
        sl = slot_of(b)
        while sl < 2 and slots[sl]["img"] is not None:
            sl += 1  # image wale slot se text agle slot me
        slots[sl]["blocks"].append(b)
    return slots


def xlsx_hf_strings(slots):
    """3 slots se header/footer ki final strings (255 character ki limit ke andar)."""
    for fmt in (True, False):
        texts = []
        for sl in slots:
            parts = [_hf_block_text(b, fmt) for b in sl["blocks"]]
            t = "\n".join(parts)
            if sl["img"] is not None:
                t = "&G" + ("\n" + t if t else "")
            texts.append(t)
        total = sum(len(t) + 2 for t in texts if t)
        if total <= 250 or not fmt:
            break
    while sum(len(t) + 2 for t in texts if t) > 250:  # phir bhi bada ho to sabse lambi string chhoti karo
        i = max(range(3), key=lambda k: len(texts[k]))
        texts[i] = texts[i][:max(0, len(texts[i]) - 10)]
    return texts


def apply_xlsx_hf(ws, data):
    """Excel ke real Page Header/Footer set karta hai. Returns {shape_id: (png_bytes, w_pt, h_pt)} (images)."""
    imgs = {}
    for key, obj, tag in (("header", ws.oddHeader, "H"), ("footer", ws.oddFooter, "F")):
        region = data.get(key)
        if not region or not region.get("blocks"):
            continue
        slots = xlsx_hf_slots(region)
        texts = xlsx_hf_strings(slots)
        for part, t in zip((obj.left, obj.center, obj.right), texts):
            if t:
                part.text = t
        need_in = 0.0
        for si, sl in enumerate(slots):
            if sl["img"] is not None:
                b = sl["img"]
                path = STORAGE / Path(b["file"]).name
                if path.exists():
                    w, h = float(b["width"]), float(b["height"])
                    if h > 100:
                        w, h = w * 100.0 / h, 100.0
                    imgs["LCR"[si] + tag] = (path.read_bytes(), w, h)
                    need_in = max(need_in, h / 72.0)
            text_pt = 0.0   # is slot ke text ki asli height (har block ki apni font size)
            for bb in sl["blocks"]:
                n_ln = len(bb.get("lines") or []) or 1
                sz = 9 if bb["type"] == "table" else min(bb.get("size") or 9, 18)
                text_pt += n_ln * max(8, sz) * 1.3
            need_in = max(need_in, text_pt / 72.0)
        if need_in:  # header/footer ke liye page margin itna rakho ke body se na takraye
            if tag == "H":
                ws.page_margins.top = max(0.75, need_in + 0.45)
                ws.page_margins.header = 0.3
            else:
                ws.page_margins.bottom = max(0.75, need_in + 0.45)
                ws.page_margins.footer = 0.3
    return imgs


def fix_xlsx_hf_newlines(xlsx_bytes, sheet_path="xl/worksheets/sheet1.xml"):
    """openpyxl header/footer ki new line ko '_x000a_' likhta hai; use asli new line bana deta hai."""
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    names = zin.namelist()
    files = {n: zin.read(n) for n in names}
    zin.close()
    sx = files.get(sheet_path, b"").decode("utf-8")
    m = re.search(r"<headerFooter>.*?</headerFooter>", sx, re.S)
    if not m or "_x000a_" not in m.group(0):
        return xlsx_bytes
    files[sheet_path] = (sx[:m.start()] + m.group(0).replace("_x000a_", "&#10;") + sx[m.end():]).encode("utf-8")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    return out.getvalue()


def inject_xlsx_hf_images(xlsx_bytes, imgs, sheet_path="xl/worksheets/sheet1.xml"):
    """openpyxl header/footer me image nahi laga sakta, is liye .xlsx ke andar seedha VML + image jodte hain
    (bilkul wahi structure jo Excel khud banata hai)."""
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    files = {n: zin.read(n) for n in zin.namelist()}
    order = list(zin.namelist())
    zin.close()
    if sheet_path not in files:
        return xlsx_bytes
    n = 1
    while f"xl/drawings/vmlDrawing{n}.vml" in files:
        n += 1
    vml_name = f"xl/drawings/vmlDrawing{n}.vml"
    vml_rels = f"xl/drawings/_rels/vmlDrawing{n}.vml.rels"
    shapes, rels = [], []
    for k, (sid, (data, w, h)) in enumerate(imgs.items(), 1):
        media = f"xl/media/hfimage{n}_{k}.png"
        files[media] = data
        order.append(media)
        rels.append(f'<Relationship Id="rId{k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/image" Target="../media/hfimage{n}_{k}.png"/>')
        shapes.append(
            f'<v:shape id="{sid}" o:spid="_x0000_s{1024 + k}" type="#_x0000_t75" '
            f'style="position:absolute;margin-left:0;margin-top:0;width:{w:.2f}pt;height:{h:.2f}pt;z-index:{k}">'
            f'<v:imagedata o:relid="rId{k}" o:title="hf{k}"/><o:lock v:ext="edit" rotation="t"/></v:shape>')
    vml = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
           'xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit">'
           '<o:idmap v:ext="edit" data="1"/></o:shapelayout>'
           '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" '
           'path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas>'
           '<v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/>'
           '<v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>'
           '<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'
           '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/>'
           '</v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>'
           '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>' + "".join(shapes) + "</xml>")
    files[vml_name] = vml.encode("utf-8")
    files[vml_rels] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       + "".join(rels) + "</Relationships>").encode("utf-8")
    order += [vml_name, vml_rels]

    # sheet ki rels file me vmlDrawing ka relationship
    srel = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
    rel_xml = ('<Relationship Id="rIdHF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
               f'relationships/vmlDrawing" Target="../drawings/vmlDrawing{n}.vml"/>')
    if srel in files:
        t = files[srel].decode("utf-8")
        files[srel] = t.replace("</Relationships>", rel_xml + "</Relationships>").encode("utf-8")
    else:
        files[srel] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       + rel_xml + "</Relationships>").encode("utf-8")
        order.append(srel)

    # sheet xml me <legacyDrawingHF>
    sx = files[sheet_path].decode("utf-8")
    tag = ('<legacyDrawingHF xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'r:id="rIdHF1"/>')
    pos = sx.find("<tableParts")
    if pos < 0:
        pos = sx.find("<extLst")
    if pos < 0:
        pos = sx.rfind("</worksheet>")
    files[sheet_path] = (sx[:pos] + tag + sx[pos:]).encode("utf-8")

    # content types: vml aur png
    ct = files["[Content_Types].xml"].decode("utf-8")
    add = ""
    for ext, typ in (("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
                     ("png", "image/png")):
        if f'Extension="{ext}"' not in ct:
            add += f'<Default Extension="{ext}" ContentType="{typ}"/>'
    files["[Content_Types].xml"] = ct.replace("</Types>", add + "</Types>").encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            if name in files:
                zout.writestr(name, files[name])
    return out.getvalue()


def write_xlsx_blocks(ws, r, blocks, groups, pw, ctx, auto_align=False):
    """Blocks (para / image / table / side-by-side group) ko row r se likhta hai. Returns next free row."""
    widths, col_px, total_px, total_chars, span = ctx
    pscale = total_px / pw                    # PDF point -> sheet pixel (position)
    kscale = min(pscale, 96.0 / 72.0)         # image size factor
    max_px = max(60, total_chars * 7)
    i = 0
    while i < len(blocks):
        b = blocks[i]
        gid = b.get("group")
        if gid is not None and gid < len(groups):
            j = i
            members = []
            while j < len(blocks) and blocks[j].get("group") == gid:
                members.append(blocks[j])
                j += 1
            r = write_xlsx_group(ws, r, members, groups[gid], col_px, widths,
                                 pscale, kscale, max_px, pw)
            i = j
            continue
        if b["type"] == "para":
            ha = slab_align(b, 0.0, pw) if (auto_align and "x0" in b) else "left"
            write_xlsx_para(ws, r, b, 1, span, total_chars, ha)
            r += 1
        elif b["type"] == "image":
            try:
                xp = b["x0"] * pscale if "x0" in b else None
                r = add_xlsx_image(ws, r, b, max_px, col_px, xp, kscale)
            except Exception:
                pass  # ek image kharab ho to baaki sheet na ruke
        else:
            r = write_table(ws, r, b["rows"], 1, b.get("cell_images")) + 1
        i += 1
    return r


def hf_blocks_for_page(region, pno, total):
    """Header/footer ke blocks ki copy; 'Page 1 of 2' jaisi text me is page ka number aur total pages likhta hai."""
    out = []
    for b in region["blocks"]:
        nb = dict(b)
        if nb.get("segments"):
            nb["text"] = "".join(t if k == "t" else str(pno if k == "PAGE" else total)
                                 for k, t in nb["segments"])
            nb.pop("segments", None)
        out.append(nb)
    return out


def setup_xlsx_print(ws, data, breaks, total_px, spans=None):
    """Print/Page setup: PDF jaisa paper size, ek page ki width me fit, har PDF page ke baad page break."""
    pg = data["pages"][0] if data.get("pages") else {}
    pw = float(pg.get("width") or 595.0)
    ph = float(pg.get("height") or 842.0)
    letter = abs(pw - 612) < 8 and abs(ph - 792) < 8
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER if letter else ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if pw > ph else "portrait"
    ws.page_margins.left = ws.page_margins.right = 0.5
    if XLSX_HF_MODE == "sheet":                               # header/footer cells me hain, margin chhota
        ws.page_margins.top = ws.page_margins.bottom = 0.4
        ws.page_margins.header = ws.page_margins.footer = 0.2
    printable_pt = max(100.0, pw - 72.0)                      # 0.5in + 0.5in margin
    scale = printable_pt / (total_px * 0.75) * 100 * 0.98     # width ke hisaab se
    printable_h = ph - (ws.page_margins.top + ws.page_margins.bottom) * 72.0
    for a, b in spans or []:                                  # har PDF page ki rows ek print page me fit hon
        h = sum((ws.row_dimensions[i].height or 15.0) for i in range(a, b + 1))
        if h > 0:
            scale = min(scale, printable_h / h * 100 * 0.94)
    scale = int(min(100, max(10, scale)))
    ws.page_setup.scale = scale                               # sheet ki poori width ek page me
    ws.print_options.horizontalCentered = True
    for rb in breaks:                                         # PDF ka page 2 = print ka page 2
        if rb > 0:
            ws.row_breaks.append(Break(id=rb))
    if XLSX_HF_MODE != "sheet" and XLSX_VIEW in ("pageLayout", "normal", "pageBreakPreview"):
        ws.sheet_view.view = XLSX_VIEW                        # Page Layout view me header/footer screen par dikhte hain


def build_xlsx(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Full Document"

    table_blocks = [(p["number"], b) for p in data["pages"]
                    for b in p["blocks"] if b["type"] == "table"]
    all_tables = [(pno, b["rows"]) for pno, b in table_blocks]
    span = max([6] + [max(len(r) for r in rows) for _, rows in all_tables])
    widths = [18] * span
    for _, rows in all_tables:
        for i, w in enumerate(table_widths(rows)):
            widths[i] = max(widths[i], w) if widths[i] != 18 else w
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    total_chars = sum(widths)
    col_px = [w * 7 + 5 for w in widths]
    total_px = sum(col_px)
    ctx = (widths, col_px, total_px, total_chars, span)
    hf_imgs = {}
    if XLSX_HF_MODE != "sheet":
        try:  # asli Excel header/footer (body me nahi)
            hf_imgs = apply_xlsx_hf(ws, data)
        except Exception:
            hf_imgs = {}

    r = 1
    breaks = []
    hdr, ftr = data.get("header"), data.get("footer")
    in_sheet = XLSX_HF_MODE == "sheet"
    n_pages = len(data["pages"])
    first = True
    spans = []
    for page in data["pages"]:
        if not page["blocks"]:
            continue
        if not first:
            breaks.append(r - 1)   # naya PDF page = naya print page
        span_start = r
        if in_sheet and hdr and hdr.get("blocks"):   # har page ke upar header (PDF ki tarah)
            r = write_xlsx_blocks(ws, r, hf_blocks_for_page(hdr, page["number"], n_pages),
                                  hdr.get("groups", []), float(hdr.get("width") or 595.0), ctx,
                                  auto_align=True) + 1
        if SHOW_PAGE_MARKERS:
            c = set_cell(ws, r, 1, f"--- Page {page['number']} ---")
            c.font = Font(italic=True, color="888888")
            r += 1
        r = write_xlsx_blocks(ws, r, page["blocks"], page.get("groups", []),
                              float(page.get("width") or 595.0), ctx)
        r += 1
        if in_sheet and ftr and ftr.get("blocks"):   # har page ke neeche footer (sahi Page X of Y)
            r = write_xlsx_blocks(ws, r, hf_blocks_for_page(ftr, page["number"], n_pages),
                                  ftr.get("groups", []), float(ftr.get("width") or 595.0), ctx,
                                  auto_align=True) + 1
        spans.append((span_start, r - 1))
        first = False
    setup_xlsx_print(ws, data, breaks, total_px, spans)

    for n, (pno, tb) in enumerate(table_blocks if XLSX_TABLE_SHEETS else [], 1):
        rows = tb["rows"]
        sheet = wb.create_sheet(f"P{pno}_Table{n}"[:31])
        for i, w in enumerate(table_widths(rows), 1):
            sheet.column_dimensions[get_column_letter(i)].width = w
        write_table(sheet, 1, rows, 1, tb.get("cell_images"))
        if len(rows) > 1:
            sheet.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    try:
        buf = io.BytesIO(fix_xlsx_hf_newlines(buf.getvalue()))
    except Exception:
        buf.seek(0)
    if hf_imgs:
        try:
            buf = io.BytesIO(inject_xlsx_hf_images(buf.getvalue(), hf_imgs))
        except Exception:
            buf.seek(0)  # image jodna fail ho to bina header-image ke file phir bhi sahi milegi
    buf.seek(0)
    return buf


# ------------------------------------------------------------------
# Storage / cleanup / preview
# ------------------------------------------------------------------
def cleanup_old(max_age=3600):
    now = time.time()
    for f in STORAGE.glob("*"):  # json + images dono purani files yahin delete hoti hain
        try:
            if now - f.stat().st_mtime > max_age:
                f.unlink()
        except OSError:
            pass


def build_preview(data, limit=6):
    out = []
    for p in data["pages"]:
        for b in p["blocks"]:
            if len(out) >= limit:
                return out
            if b["type"] == "para":
                out.append({"type": "para", "text": b["text"][:240],
                            "heading": b["heading"], "bold": b["bold"]})
            elif b["type"] == "table":
                out.append({"type": "table", "rows": [r[:8] for r in b["rows"][:5]]})
            # image blocks preview me nahi jate (frontend purane format par hi chalta rahe)
    return out


def build_export(fid, fmt):
    """Word/Excel file banata hai. Browser download aur desktop 'Save As' dono yehi use karte hain.
    Returns (BytesIO, file_name, mimetype)."""
    if not FID_RE.match(fid or "") or fmt not in ("word", "excel"):
        raise ValueError("Invalid request.")
    path = STORAGE / f"{fid}.json"
    if not path.exists():
        raise FileNotFoundError("Session expired. Please upload the PDF again.")
    data = json.loads(path.read_text(encoding="utf-8"))
    stem = re.sub(r"[^\w\- ]", "_", Path(data["filename"]).stem) or "converted"
    if fmt == "word":
        return build_docx(data), f"{stem}.docx", MIME_DOCX
    return build_xlsx(data), f"{stem}.xlsx", MIME_XLSX


# ------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------
@app.get("/")
def index():
    return render_template("index.html")


@app.get("/favicon.ico")
def favicon():
    ico = resource_path("convert.ico")
    if ico.exists():
        return send_from_directory(str(ico.parent), ico.name, mimetype="image/x-icon")
    return ("", 204)


@app.post("/api/upload")
def upload():
    cleanup_old()
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Please select a PDF file."), 400
    if not f.filename.lower().endswith(".pdf"):
        return jsonify(error="Only PDF files are allowed."), 400
    head = f.stream.read(5)
    f.stream.seek(0)
    if head != b"%PDF-":
        return jsonify(error="This file is not a valid PDF."), 400

    fid = uuid.uuid4().hex
    pdf_path = STORAGE / f"{fid}.pdf"
    f.save(pdf_path)
    try:
        data = extract_pdf(pdf_path, f.filename, fid)
    except Exception as e:
        remove_images(fid)
        msg = str(e).lower()
        if "password" in msg or "password" in type(e).__name__.lower():
            return jsonify(error="This PDF is password protected."), 422
        return jsonify(error="Could not read this PDF. It may be corrupted."), 422
    finally:
        try:
            pdf_path.unlink()
        except OSError:
            pass

    st = data["stats"]
    if st["paragraphs"] == 0 and st["tables"] == 0:
        remove_images(fid)
        return jsonify(error="No text found. This looks like a scanned/image PDF (OCR required)."), 422

    (STORAGE / f"{fid}.json").write_text(json.dumps(data), encoding="utf-8")
    return jsonify(id=fid, filename=f.filename, stats=st, preview=build_preview(data))


@app.get("/api/download/<fid>/<fmt>")
def download(fid, fmt):
    try:
        buf, name, mime = build_export(fid, fmt)
    except ValueError as e:
        return jsonify(error=str(e)), 400
    except FileNotFoundError as e:
        return jsonify(error=str(e)), 404
    return send_file(buf, as_attachment=True, download_name=name, mimetype=mime)


@app.post("/api/discard/<fid>")
def discard(fid):
    """Reset ke baad extracted temp data (json + images) delete kar deta hai."""
    if FID_RE.match(fid):
        try:
            (STORAGE / f"{fid}.json").unlink()
        except OSError:
            pass
        remove_images(fid)
    return jsonify(ok=True)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="File too large. Maximum size is 50 MB."), 413


if __name__ == "__main__":
    app.run(debug=True, port=5010)