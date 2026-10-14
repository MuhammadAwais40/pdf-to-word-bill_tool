import io
import os
import re
import sys
import json
import time
import uuid
import math
import tempfile
import statistics
from collections import Counter
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import pdfplumber
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# Paths (script mode + PyInstaller exe mode dono ke liye)
# ------------------------------------------------------------------
def resource_path(rel):
    """Bundled files (templates/static/icon). EXE me ye _MEIPASS se aati hain."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


BASE_DIR = Path(__file__).resolve().parent
if getattr(sys, "frozen", False):
    # EXE: temp data user ke AppData folder me (exe ke andar write nahi ho sakta)
    APP_DATA = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "PDFConverter"
else:
    APP_DATA = BASE_DIR
STORAGE = APP_DATA / "storage"
STORAGE.mkdir(parents=True, exist_ok=True)

app = Flask(
    __name__,
    template_folder=str(resource_path("templates")),
    static_folder=str(resource_path("static")),
)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
PUA_RE = re.compile(r"[\uf000-\uf8ff]")
BULLET_CHARS = "\u2022\u25AA\u27A2\u2713\u25E6\u25CF\u25A0"
BULLET_RE = re.compile(
    r"^([%s]\s*|-\s+|\*\s+|\d+[\.\)]\s+|\d+(\.\d+)+\.?\s+|\(?[a-zA-Z]\)\s+|[a-zA-Z]\.\s+)" % BULLET_CHARS
)
MARK_RE = re.compile(r"^([%s]|\(?[a-zA-Z]\)|[a-zA-Z]\.|\d+[\.\)])\s+(.+)$" % BULLET_CHARS, re.S)
BULLET_START_RE = re.compile(r"^([%s])(?=\S)" % BULLET_CHARS)
PAGE_NUM_RE = re.compile(r"^(page\s*)?\d{1,4}(\s*(of|/)\s*\d{1,4})?$|^-\s*\d{1,4}\s*-$", re.I)
FID_RE = re.compile(r"^[0-9a-f]{32}$")
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Word me PDF ke har page ke baad forced page break lagana hai ya nahi.
# False = text/table naturally flow karte hain (tables page ke across sahi merge hoti hain).
PAGE_BREAKS_BETWEEN_PAGES = False
EDGE_ZONE = 0.08      # page ke top/bottom 8% = header/footer area
CELL_PAD = 3.0        # cell ke andar text ka approx left/right padding (points)


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def fix_glyphs(s):
    """PDF ke symbol-font bullets (private-use characters) ko normal bullet bana deta hai."""
    s = re.sub(r"\(cid:(127|149|183|8226)\)", "\u2022", s or "")
    s = s.replace("\uf0fc", "\u2713").replace("\uf0d8", "\u27A2").replace("\uf0a7", "\u25AA")
    return PUA_RE.sub("\u2022", s)


def clean_text(s):
    s = fix_glyphs(CTRL_RE.sub("", s or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return BULLET_START_RE.sub(r"\1 ", s)


def clean_cell(s):
    s = fix_glyphs(CTRL_RE.sub("", s or ""))
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.splitlines()]
    lines = [BULLET_START_RE.sub(r"\1 ", ln) for ln in lines if ln]
    return "\n".join(lines)


def normalize_table(rows, widths=None):
    """Empty rows/columns hata deta hai. Column widths ko bhi saath adjust karta hai."""
    rows = [[clean_cell(c) for c in r] for r in (rows or []) if r]
    rows = [r for r in rows if any(r)]
    if not rows:
        return [], None
    ncols = max(len(r) for r in rows)
    rows = [r + [""] * (ncols - len(r)) for r in rows]
    keep = [i for i in range(ncols) if any(r[i] for r in rows)]
    rows = [[r[i] for i in keep] for r in rows]
    if widths and len(widths) == ncols and keep:
        nw = []
        for pos, i in enumerate(keep):
            end = keep[pos + 1] if pos + 1 < len(keep) else ncols
            nw.append(sum(widths[i:end]))
        nw[0] += sum(widths[:keep[0]])
        widths = nw
    else:
        widths = None
    return rows, widths


def normalize_rows(rows):
    return normalize_table(rows)[0]


def looks_like_table(rows):
    """Borderless (text-aligned) tables ke liye validation, false positives kam karne ke liye."""
    if len(rows) < 3:
        return False
    ncols = len(rows[0])
    if ncols < 2 or ncols > 12:
        return False
    total = len(rows) * ncols
    filled = sum(1 for r in rows for c in r if c)
    multi = sum(1 for r in rows if sum(1 for c in r if c) >= 2)
    return filled / total >= 0.5 and multi / len(rows) >= 0.6


# ---------- cell text (position based) ----------
def page_words(page):
    try:
        return page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False,
                                  extra_attrs=["size"])
    except Exception:
        return []


def cell_lines(words, bbox):
    """Ek cell ke andar ke words ko lines me group karta hai (x0/x1/top/bottom ke saath)."""
    x0, top, x1, bottom = bbox
    inside = [w for w in words
              if x0 - 0.5 <= (w["x0"] + w["x1"]) / 2 <= x1 + 0.5
              and top - 0.5 <= (w["top"] + w["bottom"]) / 2 <= bottom + 0.5]
    inside.sort(key=lambda w: (w["top"], w["x0"]))
    groups = []
    for w in inside:
        if groups and abs(w["top"] - groups[-1][0]["top"]) <= max(2.0, 0.5 * (w["bottom"] - w["top"])):
            groups[-1].append(w)
        else:
            groups.append([w])
    out = []
    for g in groups:
        g.sort(key=lambda w: w["x0"])
        out.append({
            "text": " ".join(w["text"] for w in g),
            "x0": g[0]["x0"], "x1": g[-1]["x1"],
            "top": min(w["top"] for w in g), "bottom": max(w["bottom"] for w in g),
            "first_w": g[0]["x1"] - g[0]["x0"],
            "size": statistics.median([w.get("size", 9) for w in g]),
        })
    return out


def reflow_cell(lines, bbox, col_limit=None):
    """PDF ke line-wrap hata kar proper paragraphs banata hai.
    Naya paragraph tab shuru hota hai jab: bullet/numbering ho, gap bara ho,
    ya agli line ka pehla word pichli line me fit ho sakta tha (matlab wrap nahi, hard break tha)."""
    cell_w = bbox[2] - bbox[0]
    cx = (bbox[0] + bbox[2]) / 2
    inner = max(cell_w - 2 * CELL_PAD, 1)
    limit = bbox[2] - CELL_PAD
    if col_limit:
        limit = min(limit, col_limit + 1)
    centered = bool(lines) and all(abs((l["x0"] + l["x1"]) / 2 - cx) < 3 for l in lines)
    paras, prev = [], None
    for ln in lines:
        text = clean_text(ln["text"])
        if not text:
            continue
        if prev is None:
            paras.append(text)
        else:
            height = max(prev["bottom"] - prev["top"], 1)
            gap = ln["top"] - prev["bottom"]
            if centered:
                fits = (prev["x1"] - prev["x0"]) + 3 + ln["first_w"] <= inner - 1
            else:
                fits = prev["x1"] + 3 + ln["first_w"] <= limit
            # lowercase se shuru hone wali line, jab pichli line sentence-end par na ho = wrap (continuation)
            cont = text[:1].islower() and not re.search(r"[.:;!?)\]]$", paras[-1].strip())
            hard = bool(BULLET_RE.match(text)) or gap > 0.6 * height or (fits and not cont)
            if hard:
                paras.append(text)
            elif paras[-1].endswith("-") and text[:1].islower():
                paras[-1] = paras[-1][:-1] + text
            else:
                paras[-1] += " " + text
        prev = ln
    return "\n".join(paras)


def reflow_rows(grid):
    """grid: rows of (column_key, lines, bbox) ya None.
    Har column ki wrap-limit (sab se lambi line) nikal kar cells reflow karta hai."""
    limits = {}
    for row in grid:
        for c in row:
            if c and c[1]:
                limits[c[0]] = max(limits.get(c[0], 0), max(l["x1"] for l in c[1]))
    return [[reflow_cell(c[1], c[2], limits.get(c[0])) if c else "" for c in row] for row in grid]


SNAP_SETTINGS = {"snap_tolerance": 5, "join_tolerance": 5, "intersection_tolerance": 5}


def _cluster(vals, tol):
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def is_messy(t):
    """Double-border wali tables me pdfplumber bohat patle (sliver) columns bana deta hai."""
    try:
        xs = sorted(set(c[0] for c in t.cells))
        ws = [(xs[i + 1] if i + 1 < len(xs) else t.bbox[2]) - xs[i] for i in range(len(xs))]
        return any(w < 8 for w in ws)
    except Exception:
        return False


def build_grid_table(page, t, get_words):
    """Clean grid: columns = snapped vertical lines, rows = (lagbhag) poori-width horizontal lines.
    Double borders wali tables ko stable rows/columns me badal deta hai."""
    snapped = None
    for st in page.find_tables(SNAP_SETTINGS):
        ox = min(st.bbox[2], t.bbox[2]) - max(st.bbox[0], t.bbox[0])
        oy = min(st.bbox[3], t.bbox[3]) - max(st.bbox[1], t.bbox[1])
        if ox > 0 and oy > 0:
            snapped = st
            break
    if snapped is None:
        return None
    x0, top, x1, bottom = snapped.bbox
    xs = _cluster([c[0] for c in snapped.cells], 3)
    xs.append(x1)

    hs = [e for e in page.edges
          if e.get("orientation") == "h" and top - 3 <= e["top"] <= bottom + 3
          and e["x1"] > x0 and e["x0"] < x1]
    groups = []
    for e in sorted(hs, key=lambda e: e["top"]):
        if groups and e["top"] - groups[-1][-1]["top"] <= 5:
            groups[-1].append(e)
        else:
            groups.append([e])
    seps = []
    for g in groups:
        iv = sorted((max(e["x0"], x0), min(e["x1"], x1)) for e in g)
        cov, cs, ce = 0.0, iv[0][0], iv[0][1]
        for s_, e_ in iv[1:]:
            if s_ <= ce + 2:
                ce = max(ce, e_)
            else:
                cov += ce - cs
                cs, ce = s_, e_
        cov += ce - cs
        if cov >= 0.85 * (x1 - x0):
            seps.append(sum(e["top"] for e in g) / len(g))
    ys = _cluster(seps + [top, bottom], 5)
    if len(xs) < 3 or len(ys) < 2:
        return None

    words = get_words()
    grid, sizes = [], []
    for ri in range(len(ys) - 1):
        r = []
        for ci in range(len(xs) - 1):
            bbox = (xs[ci], ys[ri], xs[ci + 1], ys[ri + 1])
            lines = cell_lines(words, bbox)
            sizes += [l["size"] for l in lines]
            r.append((ci, lines, bbox))
        grid.append(r)
    rows = reflow_rows(grid)
    widths = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
    rows, widths = normalize_table(rows, widths)
    font = round(statistics.median(sizes) * 2) / 2 if sizes else 0
    return {"bbox": snapped.bbox, "rows": rows, "widths": widths, "font": font}


def build_table(page, t, get_words):
    """pdfplumber table -> {bbox, rows, widths, font}  (normal tables)"""
    widths = None
    try:
        xs = sorted(set(c[0] for c in t.cells))
        widths = [(xs[i + 1] if i + 1 < len(xs) else t.bbox[2]) - xs[i] for i in range(len(xs))]
    except Exception:
        widths = None
    sizes = []
    try:
        words = get_words()
        grid = []
        for row in t.rows:
            r = []
            for bbox in row.cells:
                if bbox is None:
                    r.append(None)
                    continue
                lines = cell_lines(words, bbox)
                sizes += [l["size"] for l in lines]
                r.append((round(bbox[0] / 3), lines, bbox))
            grid.append(r)
        rows = reflow_rows(grid)
    except Exception:
        rows, sizes = t.extract(), []
        widths = None
    rows, widths = normalize_table(rows, widths)
    font = round(statistics.median(sizes) * 2) / 2 if sizes else 0
    return {"bbox": t.bbox, "rows": rows, "widths": widths, "font": font}


def find_tables(page):
    """Returns [table dict]. Pehle bordered tables, phir borderless fallback."""
    cache = {}

    def get_words():
        if "w" not in cache:
            cache["w"] = page_words(page)
        return cache["w"]

    found = []
    for t in page.find_tables():
        tb = None
        if is_messy(t):
            try:
                tb = build_grid_table(page, t, get_words)
            except Exception:
                tb = None
        if tb is None:
            tb = build_table(page, t, get_words)
        if tb["rows"] and len(tb["rows"][0]) >= 2:
            found.append(tb)
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
            tb = build_table(page, t, get_words)
            if looks_like_table(tb["rows"]):
                found.append(tb)
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
    return {
        "text": text, "size": size, "bold": bold,
        "x0": float(ln.get("x0", 0)), "x1": float(ln.get("x1", 0)),
        "top": float(ln["top"]), "bottom": float(ln["bottom"]),
    }


def para_align(x0, x1, W, left, right):
    """Paragraph left / center / right aligned hai, position se andaza."""
    cw = max(right - left, W - 144)
    width = x1 - x0
    if width < 0.85 * cw and abs((x0 + x1) / 2 - W / 2) < 0.025 * W:
        return "center"
    if width < 0.75 * cw and abs(x1 - right) < 0.02 * W and x0 - left > 0.2 * cw:
        return "right"
    return "left"


def build_paragraphs(lines, body, W=0, left=0, right=0):
    """Lines ko paragraphs aur headings me convert karta hai."""
    paras, cur = [], None
    for ln in lines:
        ratio = ln["size"] / body if body else 1
        level = 0
        if ratio >= 1.2:
            level = 1 if ratio >= 1.7 else (2 if ratio >= 1.4 else 3)
        cw = max(right - left, W - 144) if W else 0
        centered = bool(W) and (ln["x1"] - ln["x0"]) < 0.85 * cw and abs((ln["x0"] + ln["x1"]) / 2 - W / 2) < 0.025 * W
        split = True
        if cur:
            gap = ln["top"] - cur["bottom"]
            limit = max(ln["size"], cur["size"], 1) * 0.6
            split = bool(
                gap > limit or level or cur["heading"]
                or BULLET_RE.match(ln["text"])
                or abs(ln["size"] - cur["size"]) > 1
                or ln["bold"] != cur["bold"]
                or (centered and cur["centered"])   # alag alag centered title lines
            )
        if cur and not split:
            if cur["text"].endswith("-") and ln["text"][:1].islower():
                cur["text"] = cur["text"][:-1] + ln["text"]
            else:
                cur["text"] += " " + ln["text"]
            cur["bottom"] = ln["bottom"]
            cur["x0"] = min(cur["x0"], ln["x0"])
            cur["x1"] = max(cur["x1"], ln["x1"])
            cur["centered"] = centered
        else:
            if cur:
                paras.append(cur)
            cur = {"text": ln["text"], "size": ln["size"], "bold": ln["bold"],
                   "heading": level, "bottom": ln["bottom"],
                   "x0": ln["x0"], "x1": ln["x1"], "centered": centered}
    if cur:
        paras.append(cur)
    out = []
    for p in paras:
        align = para_align(p["x0"], p["x1"], W, left, right) if W else "left"
        out.append({"type": "para", "text": p["text"], "size": p["size"],
                    "bold": p["bold"], "heading": p["heading"], "align": align})
    return out


# ---------- header / footer / page number noise ----------
def edge_zone(d, H):
    if d["top"] < H * EDGE_ZONE:
        return "top"
    if d["bottom"] > H * (1 - EDGE_ZONE):
        return "bottom"
    return None


def running_keys(raw_pages):
    """Jo text 60%+ pages ke top/bottom me repeat ho (running header/footer)."""
    n = len(raw_pages)
    if n < 3:
        return set()
    cnt = Counter()
    for pg in raw_pages:
        keys = set()
        for d in pg["lines"]:
            z = edge_zone(d, pg["h"])
            if z:
                keys.add((z, re.sub(r"\d+", "#", d["text"].lower())))
        cnt.update(keys)
    return {k for k, v in cnt.items() if v >= max(3, 0.6 * n)}


# ---------- tables across pages ----------
def _norm_row(row):
    return tuple(re.sub(r"\s+", "", c.lower()) for c in row)


def _widths_similar(a, b):
    wa, wb = a.get("widths"), b.get("widths")
    if not wa or not wb or len(wa) != len(wb):
        return False
    return all(abs(x - y) <= max(8.0, 0.12 * max(x, y)) for x, y in zip(wa, wb))


def join_cell(a, b):
    """Page break se kati hui cell ka text wapas jorta hai."""
    if not b:
        return a
    if not a:
        return b
    la, lb = a.split("\n"), b.split("\n")
    ends = re.search(r"[.:;!?)\]\"\u201d']$", la[-1].strip())
    if not ends and not BULLET_RE.match(lb[0]):
        if la[-1].endswith("-") and lb[0][:1].islower():
            la[-1] = la[-1][:-1] + lb[0]
        else:
            la[-1] += " " + lb[0]
        return "\n".join(la + lb[1:])
    return a + "\n" + b


def merge_page_tables(pages):
    """Agar table agle page par continue ho rahi ho to ek hi table bana deta hai:
    repeated header row hata deta hai aur page se kati hui row ko wapas jor deta hai."""
    for i in range(1, len(pages)):
        cur = pages[i]["blocks"]
        if not cur or cur[0]["type"] != "table":
            continue
        j = i - 1
        while j >= 0 and not pages[j]["blocks"]:
            j -= 1
        if j < 0:
            continue
        a = pages[j]["blocks"][-1]
        b = cur[0]
        if a["type"] != "table" or not a["rows"] or not b["rows"]:
            continue
        if abs(pages[j]["w"] - pages[i]["w"]) > 2 or abs(pages[j]["h"] - pages[i]["h"]) > 2:
            continue
        if len(a["rows"][0]) != len(b["rows"][0]):
            continue
        repeated = _norm_row(b["rows"][0]) == _norm_row(a["rows"][0])
        if not repeated and not _widths_similar(a, b):
            continue
        rows_b = [list(r) for r in (b["rows"][1:] if repeated else b["rows"])]
        if rows_b and not rows_b[0][0].strip() and any(c.strip() for c in rows_b[0][1:]) and len(a["rows"]) > 0:
            last = a["rows"][-1]
            for ci, cell in enumerate(rows_b[0]):
                last[ci] = join_cell(last[ci], cell)
            rows_b = rows_b[1:]
        a["rows"].extend(rows_b)
        a["x1"] = max(a.get("x1", 0), b.get("x1", 0))
        cur.pop(0)


# ------------------------------------------------------------------
# PDF -> structured data
# ------------------------------------------------------------------
def extract_pdf(path, original_name):
    raw_pages = []
    with pdfplumber.open(path) as pdf:
        for idx, page in enumerate(pdf.pages, 1):
            W, H = float(page.width), float(page.height)
            try:
                tables = find_tables(page)
            except Exception:
                tables = []
            bboxes = [t["bbox"] for t in tables]
            try:
                src = page.filter(make_filter(bboxes)) if bboxes else page
                raw_lines = src.extract_text_lines(strip=True, return_chars=True)
            except Exception:
                raw_lines = []
            lines = [d for d in (line_info(ln) for ln in raw_lines) if d]
            raw_pages.append({"number": idx, "w": W, "h": H, "lines": lines, "tables": tables})
            try:
                page.flush_cache()
            except Exception:
                pass

    # page numbers aur running header/footer hata do
    running = running_keys(raw_pages)
    sizes = Counter()
    for pg in raw_pages:
        keep = []
        for d in pg["lines"]:
            z = edge_zone(d, pg["h"])
            if z and (PAGE_NUM_RE.match(d["text"])
                      or (z, re.sub(r"\d+", "#", d["text"].lower())) in running):
                continue
            keep.append(d)
            sizes[d["size"]] += len(d["text"])
        pg["lines"] = keep

    body = sizes.most_common(1)[0][0] if sizes else 11
    pages = []
    for pg in raw_pages:
        W, H = pg["w"], pg["h"]
        lines, tables = pg["lines"], pg["tables"]
        left = min([d["x0"] for d in lines] or [0])
        right = max([d["x1"] for d in lines] or [W])
        items = [(d["top"], "line", d) for d in lines]
        items += [(t["bbox"][1], "table", t) for t in tables]
        items.sort(key=lambda x: x[0])

        blocks, buf = [], []
        for _, kind, payload in items:
            if kind == "line":
                buf.append(payload)
            else:
                blocks += build_paragraphs(buf, body, W, left, right)
                buf = []
                blocks.append({"type": "table", "rows": payload["rows"],
                               "widths": payload["widths"], "font": payload["font"],
                               "x0": payload["bbox"][0], "x1": payload["bbox"][2]})
        blocks += build_paragraphs(buf, body, W, left, right)

        xs0 = [d["x0"] for d in lines] + [t["bbox"][0] for t in tables]
        xs1 = [d["x1"] for d in lines] + [t["bbox"][2] for t in tables]
        ys0 = [d["top"] for d in lines] + [t["bbox"][1] for t in tables]
        ys1 = [d["bottom"] for d in lines] + [t["bbox"][3] for t in tables]
        clamp = lambda v: max(36.0, min(72.0, v))
        margins = [clamp(min(xs0)), clamp(min(ys0)), clamp(W - max(xs1)), clamp(H - max(ys1))] \
            if xs0 else [54.0, 54.0, 54.0, 54.0]
        pages.append({"number": pg["number"], "w": W, "h": H, "margins": margins, "blocks": blocks})

    merge_page_tables(pages)

    paras = sum(1 for p in pages for b in p["blocks"] if b["type"] == "para")
    tbls = sum(1 for p in pages for b in p["blocks"] if b["type"] == "table")
    words = sum(len(b["text"].split()) for p in pages for b in p["blocks"] if b["type"] == "para")
    return {
        "filename": original_name,
        "pages": pages,
        "stats": {"pages": len(pages), "paragraphs": paras, "tables": tbls, "words": words},
    }


# ------------------------------------------------------------------
# Word builder
# ------------------------------------------------------------------
ALIGN_MAP = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
             "right": WD_ALIGN_PARAGRAPH.RIGHT}


def fill_paragraph(p, text, size=None, bold=False, space_after=6, align=None):
    """Paragraph me text likhta hai. Bullet / 'a.' / '1)' wali lines par hanging indent lagata hai."""
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    m = MARK_RE.match(text)
    if m:
        text = f"{m.group(1)}\t{m.group(2)}"
        pf.left_indent = Pt(14)
        pf.first_line_indent = Pt(-14)
    run = p.add_run(text)
    run.bold = bool(bold)
    if size:
        run.font.size = Pt(size)
    if align in ALIGN_MAP:
        p.alignment = ALIGN_MAP[align]


def set_cell_fill(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def guess_widths(rows, total):
    ncols = max(len(r) for r in rows)
    weight = [6] * ncols
    for r in rows:
        for i, c in enumerate(r):
            longest = max((len(x) for x in c.split("\n")), default=0)
            weight[i] = max(weight[i], min(longest, 40))
    s = sum(weight)
    return [total * w / s for w in weight]


def add_word_table(doc, b, usable, page_w):
    rows = b["rows"]
    ncols = max(len(r) for r in rows)
    widths = b.get("widths")
    if not widths or len(widths) != ncols:
        widths = guess_widths(rows, usable)
    total = sum(widths) or 1
    scale = min(1.0, usable / total)
    widths = [w * scale for w in widths]

    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    for side, twips in (("top", 20), ("left", 60), ("bottom", 20), ("right", 60)):
        el = OxmlElement("w:" + side)
        el.set(qn("w:w"), str(twips))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:          # Word schema order: ... tblLayout, tblCellMar, tblLook
        look.addprevious(layout)
        look.addprevious(mar)
    else:
        tblPr.append(layout)
        tblPr.append(mar)
    tw = sum(widths)
    if tw < 0.95 * usable and abs((b.get("x0", 0) + b.get("x1", 0)) / 2 - page_w / 2) < 0.03 * page_w:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for ci, w in enumerate(widths):
        table.columns[ci].width = Pt(w)
    size = b.get("font") or 9
    size = max(7, min(size, 12))
    has_head = len(rows) > 1
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = table.cell(ri, ci)
            cell.width = Pt(widths[ci])
            text = row[ci] if ci < len(row) else ""
            head = has_head and ri == 0
            for pi, part in enumerate(text.split("\n") if text else [""]):
                p = cell.paragraphs[0] if pi == 0 else cell.add_paragraph()
                fill_paragraph(p, part, size=size, bold=head, space_after=2)
            if head:
                set_cell_fill(cell, "E7EAF6")
    if has_head:  # header row har page par repeat ho
        trPr = table.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
    doc.add_paragraph()


def setup_section(sec, page):
    w, h = page.get("w") or 595.0, page.get("h") or 842.0
    m = page.get("margins") or [54, 54, 54, 54]
    sec.orientation = WD_ORIENT.LANDSCAPE if w > h else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Pt(w), Pt(h)
    sec.left_margin, sec.top_margin, sec.right_margin, sec.bottom_margin = Pt(m[0]), Pt(m[1]), Pt(m[2]), Pt(m[3])


def build_docx(data):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    sec = doc.sections[0]
    first = True
    cur_size = None
    for page in data["pages"]:
        if not page["blocks"]:
            continue
        size_key = (round(page.get("w", 0)), round(page.get("h", 0)))
        if first:
            setup_section(sec, page)
            cur_size = size_key
        elif size_key != cur_size:
            sec = doc.add_section(WD_SECTION.NEW_PAGE)
            setup_section(sec, page)
            cur_size = size_key
        elif PAGE_BREAKS_BETWEEN_PAGES:
            doc.add_page_break()
        first = False

        usable = (sec.page_width - sec.left_margin - sec.right_margin) / 12700.0   # EMU -> points
        page_w = sec.page_width / 12700.0
        for b in page["blocks"]:
            if b["type"] == "para":
                if b["heading"]:
                    h = doc.add_heading(b["text"], level=min(b["heading"], 3))
                    if b.get("align") in ALIGN_MAP:
                        h.alignment = ALIGN_MAP[b["align"]]
                else:
                    size = b.get("size") or 0
                    size = max(8, min(size, 14)) if size else None
                    fill_paragraph(doc.add_paragraph(), b["text"], size=size,
                                   bold=b["bold"], space_after=6, align=b.get("align"))
            else:
                add_word_table(doc, b, usable, page_w)
    if first:
        doc.add_paragraph("(No content found)")

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


def set_cell(ws, r, c, val):
    cell = ws.cell(row=r, column=c)
    cell.value = val
    if isinstance(val, str) and val.startswith("="):
        cell.data_type = "s"  # formula ki tarah execute na ho
    return cell


def write_table(ws, start_row, rows):
    ncols = max(len(r) for r in rows)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            raw = row[ci] if ci < len(row) else ""
            cell = set_cell(ws, start_row + ri, ci + 1, to_value(raw))
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if ri == 0 and len(rows) > 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = HEAD_FILL
    return start_row + len(rows)


def table_widths(rows, base=14, cap=45):
    ncols = max(len(r) for r in rows)
    w = [base] * ncols
    for r in rows:
        for i, c in enumerate(r):
            longest = max((len(x) for x in c.split("\n")), default=0)
            w[i] = min(cap, max(w[i], longest + 2))
    return w


def build_xlsx(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Full Document"

    all_tables = [(p["number"], b["rows"]) for p in data["pages"]
                  for b in p["blocks"] if b["type"] == "table"]
    span = max([6] + [max(len(r) for r in rows) for _, rows in all_tables])
    widths = [18] * span
    for _, rows in all_tables:
        for i, w in enumerate(table_widths(rows)):
            widths[i] = max(widths[i], w) if widths[i] != 18 else w
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    total_chars = sum(widths)

    r = 1
    for page in data["pages"]:
        if not page["blocks"]:
            continue
        c = set_cell(ws, r, 1, f"--- Page {page['number']} ---")
        c.font = Font(italic=True, color="888888")
        r += 1
        for b in page["blocks"]:
            if b["type"] == "para":
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
                size = b["size"] or 11
                cell = set_cell(ws, r, 1, b["text"])
                al = b.get("align")
                cell.alignment = Alignment(wrap_text=True, vertical="top",
                                           horizontal=al if al in ("center", "right") else None)
                cell.font = Font(bold=bool(b["heading"] or b["bold"]),
                                 size=max(10, min(size, 24)))
                lines = max(1, math.ceil(len(b["text"]) * (size / 11) / total_chars))
                ws.row_dimensions[r].height = min(409, 15 * lines * max(1, size / 11))
                r += 1
            else:
                r = write_table(ws, r, b["rows"]) + 1
        r += 1

    for n, (pno, rows) in enumerate(all_tables, 1):
        sheet = wb.create_sheet(f"P{pno}_Table{n}"[:31])
        for i, w in enumerate(table_widths(rows), 1):
            sheet.column_dimensions[get_column_letter(i)].width = w
        write_table(sheet, 1, rows)
        if len(rows) > 1:
            sheet.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ------------------------------------------------------------------
# Storage / cleanup / preview
# ------------------------------------------------------------------
def cleanup_old(max_age=3600):
    now = time.time()
    for f in STORAGE.glob("*"):
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
            else:
                out.append({"type": "table", "rows": [[(c[:140] + "\u2026") if len(c) > 140 else c for c in r[:8]]
                                                   for r in b["rows"][:5]]})
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
        data = extract_pdf(pdf_path, f.filename)
    except Exception as e:
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
    """Reset ke baad extracted temp data delete kar deta hai."""
    if FID_RE.match(fid):
        try:
            (STORAGE / f"{fid}.json").unlink()
        except OSError:
            pass
    return jsonify(ok=True)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="File too large. Maximum size is 50 MB."), 413


if __name__ == "__main__":
    app.run(debug=True, port=5009)