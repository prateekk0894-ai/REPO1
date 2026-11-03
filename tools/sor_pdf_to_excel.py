#!/usr/bin/env python3
"""Convert MP PWD Schedule-of-Rates PDFs into editable Excel workbooks
that look like the PDF pages.

* One workbook per PDF, one sheet per chapter.
* Every sheet gets a column grid built from the x-positions of the PDF table
  lines, so column widths follow the PDF.  Each PDF cell becomes an Excel cell
  merged over the grid columns (and rows) it covers, with wrapped text.
* Headings / notes / paragraphs are merged across the page width and wrapped;
  centred and bold text stays centred and bold.
* Each PDF page ends with its footer line and a page break, and the sheet is
  set to print one PDF page width per sheet page.
* Rates are stored as numbers so they can be edited and used in formulas.

Usage:  python3 tools/sor_pdf_to_excel.py input.pdf output.xlsx [road|em|grid]
"""
import math
from copy import copy
import os
import re
import sys
from collections import Counter

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

SLOGAN_KD = "yksd fuekZ.k ls yksd dY;k.k"
SLOGAN_UNI = "लोक निर्माण से लोक कल्याण"
NUM_RE = re.compile(r"^-?\d[\d,]*(\.\d+)?$")
RATE_RE = re.compile(r"^\d[\d,]*\.\d\d$")
BAD_SHEET_CHARS = re.compile(r"[\[\]\:\*\?\/\\]")
LIST_START = re.compile(r"^(\(?\d{1,2}(\.\d+)*[.)]|\(?[a-zA-Z]{1,4}\)|\(?[ivx]+\)|•|➢|\*|Note|NOTE|N\.B)")

FONT_NAME = "Times New Roman"
NO_HEADER_RATES = False  # tables without a Rate header: keep numbers as printed text
RUPEE_FMT = '"₹ "0.00'
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_H = PatternFill("solid", start_color="D9E1F2")


def font(size=10, bold=False, color=None):
    kw = {"name": FONT_NAME, "size": size, "bold": bold}
    if color:
        kw["color"] = color
    return Font(**kw)


# ---------------------------------------------------------------- text utils
def fix_text(s):
    if s is None:
        return None
    return s.replace(SLOGAN_KD, SLOGAN_UNI).replace("", "•").replace("", "➢")


KEEP_LINES = True


def keep_lines(parts):
    """PDF lines kept as separate lines inside one cell (same line breaks as the PDF)."""
    out = [p.strip() for p in parts if p and p.strip()]
    return fix_text("\n".join(out)) if out else None


def join_lines(parts):
    out = ""
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if not out:
            out = p
        elif out.endswith("-") and not out.endswith(" -"):
            out += p
        else:
            out += " " + p
    return fix_text(out)


def to_number(s):
    if not isinstance(s, str):
        return s
    t = s.strip()
    if NUM_RE.match(t):
        try:
            v = float(t.replace(",", ""))
            return int(v) if v.is_integer() and "." not in t else v
        except ValueError:
            return s
    return s


def looks_like_rate(t):
    t = t.strip()
    if not RATE_RE.match(t):
        return False
    ip = t.split(".")[0].replace(",", "")
    return t.endswith(".00") or len(ip) >= 3 or "," in t


def group_lines(words, tol=3.0):
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    lines = []
    for w in words:
        if lines and abs(w["top"] - lines[-1]["top"]) <= tol:
            lines[-1]["words"].append(w)
        else:
            lines.append({"top": w["top"], "words": [w]})
    for ln in lines:
        ws = sorted(ln["words"], key=lambda w: w["x0"])
        ln["words"] = ws
        ln["text"] = " ".join(w["text"] for w in ws)
        ln["x0"] = ws[0]["x0"]
        ln["x1"] = max(w["x1"] for w in ws)
        ln["bottom"] = max(w["bottom"] for w in ws)
        ln["size"] = Counter(round(w.get("size", 10)) for w in ws).most_common(1)[0][0]
        ln["bold"] = sum("Bold" in w.get("fontname", "") for w in ws) > len(ws) / 2
    return lines


def split_gaps(words, gap):
    chunks = [[words[0]]]
    for w in words[1:]:
        if w["x0"] - chunks[-1][-1]["x1"] >= gap:
            chunks.append([w])
        else:
            chunks[-1].append(w)
    return chunks


def is_num_token(t):
    return bool(re.fullmatch(r"[\d.,%/-]+", t)) and any(ch.isdigit() for ch in t)


def split_table_line(words):
    """Split a free-text line into cells: at big gaps, and - on lines that look like a
    row of an un-ruled numeric table - also between numbers set a few points apart."""
    nums = sum(is_num_token(w["text"]) for w in words)
    gaps_num = sum(
        1 for a, b in zip(words, words[1:])
        if b["x0"] - a["x1"] >= 6 and (is_num_token(a["text"]) or is_num_token(b["text"]))
    )
    if nums >= 3 and gaps_num >= 3:
        chunks = [[words[0]]]
        for a, b in zip(words, words[1:]):
            gap = b["x0"] - a["x1"]
            if gap >= 25 or (gap >= 6 and (is_num_token(a["text"]) or is_num_token(b["text"]))):
                chunks.append([b])
            else:
                chunks[-1].append(b)
        return chunks
    return split_gaps(words, 25)


def inside(w, bbox):
    cx = (w["x0"] + w["x1"]) / 2
    cy = (w["top"] + w["bottom"]) / 2
    return bbox[0] - 1 <= cx <= bbox[2] + 1 and bbox[1] - 1 <= cy <= bbox[3] + 1


# ---------------------------------------------------------------- chapters
class Chapters:
    def __init__(self, kind, pdf):
        self.kind = kind
        self.current = "Front Matter & General"
        self.titles = {}
        self.road_index = []
        if kind == "road":
            for pg in pdf.pages[:12]:
                txt = pg.extract_text() or ""
                if "Chapter-1" in txt and "Page No" in txt:
                    for m in re.finditer(r"(Chapter-(\d+)|Appendix)\s+(.+?)\s+(\d+)-(\d+)\s*$", txt, re.M):
                        name = m.group(1).replace("Chapter-", "Ch-")
                        self.road_index.append((int(m.group(4)), int(m.group(5)), f"{name} {m.group(3).strip()}"))
                    break

    def page(self, footer_text, top_lines):
        if self.kind == "road":
            m = re.search(r"Page\s+(\d+)\s*$", footer_text)
            if m and self.road_index:
                n = int(m.group(1))
                for a, b, name in self.road_index:
                    if a <= n <= b:
                        self.current = name
                        break
                else:
                    if n > self.road_index[-1][1]:
                        self.current = self.road_index[-1][2]
            return self.current
        if self.kind == "em":
            m = re.search(r"CHAPTER\s*-\s*(\d+)", footer_text, re.I)
            if m:
                n = int(m.group(1))
                if n not in self.titles:
                    title = ""
                    for i, t in enumerate(top_lines):
                        if re.match(rf"CHAPTER\s*[–-]?\s*{n}\b.*$", t.strip(), re.I) and i + 1 < len(top_lines):
                            title = top_lines[i + 1].strip()
                            break
                    self.titles[n] = f"Ch-{n} {title}".strip()
                self.current = self.titles[n]
            return self.current
        m = re.search(r"CHAPTER\s*[–-]?\s*(\d+)\s*[:.]?\s*(.*?)[\s.]*\d*\s*$", footer_text, re.I)
        if m:
            self.current = f"Ch-{int(m.group(1))} {m.group(2).strip(' .:')}".strip()
        return self.current


def sheet_title(name, used):
    t = BAD_SHEET_CHARS.sub("-", name).strip()
    t = re.sub(r"\s+", " ", t)[:31].strip()
    base, i = t, 2
    while t.lower() in used:
        suffix = f" ({i})"
        t = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(t.lower())
    return t


def is_footer(line, page_h, kind):
    if line["top"] < page_h * 0.88:
        return False
    t = line["text"]
    if kind == "road":
        return "SOR for Road" in t or bool(re.search(r"Page\s+\d+\s*$", t)) or SLOGAN_KD in t
    if kind == "em":
        return "ELECTRICAL SOR" in t or bool(re.search(r"Page-\s*\d+", t))
    return bool(re.match(r"\s*CHAPTER", t, re.I)) or bool(re.fullmatch(r"\s*\d+\s*", t))


# ---------------------------------------------------------------- page parse
# Block types (all coordinates in PDF points):
#   {"t": "text", "x0", "x1", "text", "size", "bold", "center"}
#   {"t": "chunks", "chunks": [(x0, text, bold)]}
#   {"t": "table", "cells": [(r0, r1, x0, x1, text)], "nrows", "hdr", "numx"}


def cell_text(words, bb):
    """Text of the PDF words whose centre lies inside the cell bbox, in reading order."""
    inner = [w for w in words if inside(w, bb) and not w.get("_used")]
    if not inner:
        return None
    for w in inner:
        w["_used"] = True
    return keep_lines([ln["text"] for ln in group_lines(inner)])


def table_block(tbl, words):
    rows = tbl.rows
    tops = [r.bbox[1] for r in rows] + [rows[-1].bbox[3]]
    cells = []
    for ri, row in enumerate(rows):
        for ci, bb in enumerate(row.cells):
            if bb is None:
                continue
            # rows spanned: those whose top lies inside the cell
            r1 = ri
            while r1 + 1 < len(rows) and tops[r1 + 1] < bb[3] - 2:
                r1 += 1
            txt = cell_text(words, bb)
            cells.append([ri, r1, bb[0], bb[2], txt, bb])
    # words that fell on a cell border: give them to the nearest cell
    for w in words:
        if w.get("_used"):
            continue
        cx, cy = (w["x0"] + w["x1"]) / 2, (w["top"] + w["bottom"]) / 2

        def dist(c):
            x0, y0, x1, y1 = c[5]
            return max(x0 - cx, 0, cx - x1) + max(y0 - cy, 0, cy - y1)

        if cells:
            c = min(cells, key=dist)
            if dist(c) <= 8:
                c[4] = f"{c[4]} {w['text']}" if c[4] else w["text"]
                w["_used"] = True
    cells = [tuple(c[:5]) for c in cells]
    # header rows: leading rows with Item/Description/Unit/Rate or just numbers 1..n
    hdr = 0
    for ri in range(min(4, len(rows))):
        joined = " ".join(c[4] for c in cells if c[0] == ri and c[4])
        longest = max((len(c[4]) for c in cells if c[0] == ri and c[4]), default=0)
        if (longest < 80 and re.search(r"\b(Item|Description|Descriptions|Unit|Rate|S\.\s*No\.?|Materials)(?=\W|$)", joined, re.I)) or (
            joined and re.fullmatch(r"[\d\s]+", joined)
        ) or (ri > 0 and ri == hdr and re.search(r"\bkm\b", joined)):
            hdr = ri + 1
        else:
            break
    numx = None
    for c in cells:
        if c[0] < max(hdr, 1) and c[4] and re.search(r"\bRate\b", c[4], re.I):
            numx = c[2] - 2
            break
    return {"t": "table", "cells": cells, "nrows": len(rows), "hdr": hdr, "numx": numx, "top": tbl.bbox[1],
            "bbox": tbl.bbox, "row_tops": tops}


ROAD_COLS = ("item", "name", "desc", "unit", "rate")
ROAD_STATE = {}
SUB_MARK = re.compile(r"^(\(?[ivxIVX]{1,5}\)|\(?[a-hA-H]\)|\([A-H]\)|[a-h]\.)$")


def starts_sub_item(words, bounds):
    """Line begins with a sub-item marker: (i), ii), a), (A), b. or a bare 'A' set apart."""
    if not words or words[0]["x0"] > bounds["item"] + 45:
        return False
    t = words[0]["text"]
    if SUB_MARK.match(t):
        return True
    m = re.match(r"^(\(?[ivx]{1,5}\)|\(?[a-h]\))\S", t)  # "(i)for" glued
    if m:
        return True
    if re.fullmatch(r"[1-9]", t) and len(words) > 1 and words[1]["text"][:1].isupper():
        return True
    return bool(re.fullmatch(r"[A-H]", t)) and len(words) > 1 and words[1]["x0"] - words[0]["x1"] > 5


def road_header(lines):
    for ln in lines:
        txt = ln["text"]
        if "Description" in txt and "Unit" in txt and "Rate" in txt and "Item" in txt:
            ws = {w["text"]: w for w in ln["words"]}
            no = ws.get("No.") or ws.get("Item")
            return ln, {
                "item": no["x1"] + 7,
                "unit": ws["Unit"]["x0"] - 12,
                "rate": (ws["Unit"]["x1"] + ws["Rate"]["x0"]) / 2 - 4,
            }
    return None, None


def looks_numeric(t):
    return bool(NUM_RE.match(t.strip()))


def shift_group_rates(recs):
    """Fix rates printed half a line too high in numbered groups.

    Pattern in the PDF:  heading (rate a) / 1 .. (rate b) / 2 .. / ... / N .. (no rate).
    Every rate belongs one row lower, so the heading carries none and N gets the last."""
    i = 0
    while i < len(recs) - 1:
        h = recs[i]
        if h["rate"] and not h["_mk"].isdigit() and recs[i + 1]["_mk"] == "1":
            j = i + 1
            while j + 1 < len(recs) and recs[j + 1]["_mk"] == str(int(recs[j]["_mk"]) + 1):
                j += 1
            grp = recs[i:j + 1]
            if not grp[-1]["rate"] and not grp[-1]["unit"] and all(r["rate"] for r in grp[:-1]):
                for k in range(len(grp) - 1, 0, -1):
                    grp[k]["rate"], grp[k]["unit"] = grp[k - 1]["rate"], grp[k - 1]["unit"]
                h["rate"], h["unit"] = [], []
            i = j + 1
        else:
            i += 1


def road_block(hdr, body, bounds, left, right):
    """Un-ruled road SOR item table -> table block with one record per item."""
    # Unit/rate words are often printed a few points above/below their text
    # line: attach each of them to the nearest text line.
    words = [w for ln in body for w in ln["words"]]
    ur = [w for w in words if (w["x0"] + w["x1"]) / 2 >= bounds["unit"]]
    txt_lines = group_lines([w for w in words if (w["x0"] + w["x1"]) / 2 < bounds["unit"]])
    extra = []
    for ul in group_lines(ur):
        same = [l for l in txt_lines if abs(l["top"] - ul["top"]) <= 2.5]
        below = [l for l in txt_lines if 0 < l["top"] - ul["top"] <= 12]
        if same:
            best = same[0]
        elif below:
            best = min(below, key=lambda l: l["top"])
        else:
            best = min(txt_lines, key=lambda l: abs(l["top"] - ul["top"]), default=None)
        if best is not None and abs(best["top"] - ul["top"]) <= 12 and not any(
            (w["x0"] + w["x1"]) / 2 >= bounds["unit"] for w in best["words"]
        ):
            best["words"] = best["words"] + ul["words"]
        else:
            extra.append(ul)
    body = sorted(txt_lines + extra, key=lambda l: l["top"])
    starts = Counter()
    for ln in body:
        for ch in split_gaps(ln["words"], 14):
            x = ch[0]["x0"]
            if bounds["item"] + 25 < x < bounds["unit"] - 60:
                starts[round(x / 3)] += 1
    col2 = sorted(k * 3 for k, c in starts.items() if c >= 4 and k * 3 < bounds["item"] + 150)[:1]
    name_x = (min(col2) - 4) if col2 else None
    xs = {"item": left, "name": bounds["item"], "desc": name_x or bounds["item"], "unit": bounds["unit"],
          "rate": bounds["rate"], "end": right}

    recs = []
    for ln in body:
        cells = {k: [] for k in ROAD_COLS}
        desc_words = []
        for w in ln["words"]:
            cx = (w["x0"] + w["x1"]) / 2
            if w["x1"] <= bounds["item"] + 2 and cx < bounds["item"]:
                cells["item"].append(w)
            elif cx >= bounds["rate"] and not (desc_words and not cells["unit"] and not looks_numeric(w["text"])
                                               and w["x0"] - desc_words[-1]["x1"] < 6):
                cells["rate"].append(w)
            elif cx >= bounds["unit"] and not (desc_words and not cells["unit"]
                                               and (w["x0"] - desc_words[-1]["x1"] < 6
                                                    or desc_words[-1]["text"].endswith("-"))
                                               and w["x0"] < bounds["unit"] + 4):
                cells["unit"].append(w)
            else:
                desc_words.append(w)
        if desc_words:
            chunks = split_gaps(desc_words, 14)
            first = chunks[0]
            if col2 and len(chunks) >= 2 and first[0]["x0"] < bounds["item"] + 20 and any(
                abs(chunks[1][0]["x0"] - c) <= 4 for c in col2
            ) and first[-1]["x1"] < col2[0]:
                cells["name"] = first
                cells["desc"] = [w for ch in chunks[1:] for w in ch]
            elif len(chunks) == 1 and col2 and first[0]["x0"] < bounds["item"] + 20 and first[-1]["x1"] < min(col2) - 8:
                cells["name"] = desc_words
            else:
                cells["desc"] = desc_words
        line = {k: " ".join(w["text"] for w in v) for k, v in cells.items()}
        zone = cells["name"] or cells["desc"]
        cur = recs[-1] if recs else None
        marked = starts_sub_item(zone, bounds)
        bold = bool(zone) and sum("Bold" in w.get("fontname", "") for w in zone) > len(zone) / 2
        new = (
            cur is None
            or bool(line["item"])
            or marked
            or (bold and (cur["_bold"] is False or bool(cur["rate"])))
            or (not bold and bool(zone) and cur["_bold"] is True)
            or (bool(line["rate"]) and bool(cur["rate"]))
            or (col2 and bool(line["name"]) and bool(cur["name"]) and (bool(line["desc"]) or bool(line["rate"])) and
                bool(cur["rate"] or cur["unit"]))
        )
        if (new and marked and not line["item"] and cur is not None and not cur["_marked"] and not cur["item"]
                and recs[:-1] and ln["top"] - cur["_top"] < 14 and not (line["rate"] and cur["rate"])):
            # marker printed half a line below the text it belongs to
            new = False
            mk = zone[0]["text"]
            first_k = "name" if cur["name"] else "desc"
            if cur[first_k]:
                cur[first_k][0] = f"{mk} {cur[first_k][0]}"
            rest = " ".join(w["text"] for w in zone[1:])
            if cells["name"]:
                line["name"] = rest
            else:
                line["desc"] = rest
            cur["_marked"] = True
        if not new and cur is not None and cur["_bold"] is None and zone:
            cur["_bold"] = bold
        if new:
            cur = {k: [] for k in ROAD_COLS}
            cur["_marked"] = marked or bool(line["item"])
            cur["_bold"] = bold if zone else None
            cur["_top"] = ln["top"]
            cur["_mk"] = zone[0]["text"] if (marked and zone) else ""
            recs.append(cur)
        for k in ROAD_COLS:
            if line[k]:
                cur[k].append(line[k])

    shift_group_rates(recs)
    cells = [(0, 0, xs["item"], xs["name"], "Item No."), (0, 0, xs["name"], xs["unit"], "Description"),
             (0, 0, xs["unit"], xs["rate"], "Unit"), (0, 0, xs["rate"], xs["end"], "Rate")]
    for i, rec in enumerate(recs, start=1):
        txt = {k: (keep_lines(rec[k]) if rec[k] else None) for k in ROAD_COLS}
        cells.append((i, i, xs["item"], xs["name"], txt["item"]))
        if col2 and txt["name"]:
            cells.append((i, i, xs["name"], xs["desc"], txt["name"]))
            cells.append((i, i, xs["desc"], xs["unit"], txt["desc"]))
        else:
            parts = [p for p in (txt["name"], txt["desc"]) if p]
            cells.append((i, i, xs["name"], xs["unit"], " ".join(parts) or None))
        cells.append((i, i, xs["unit"], xs["rate"], txt["unit"]))
        cells.append((i, i, xs["rate"], xs["end"], txt["rate"]))
    return {"t": "table", "cells": cells, "nrows": len(recs) + 1, "hdr": 1, "numx": xs["rate"] - 2,
            "top": hdr["top"], "road": True}


def text_blocks(lines, page_w, right_margin):
    """Turn free text lines into merged paragraphs / chunk rows."""
    out = []
    for ln in lines:
        chunks = split_table_line(ln["words"])
        mid = (ln["x0"] + ln["x1"]) / 2
        centered = abs(mid - page_w / 2) < 25 and (ln["x1"] - ln["x0"]) < page_w * 0.7 and ln["x0"] > 90
        if len(chunks) > 1 and not centered:
            out.append({"t": "chunks", "top": ln["top"], "bottom": ln["bottom"],
                        "chunks": [(c[0]["x0"], fix_text(" ".join(w["text"] for w in c)),
                                    sum("Bold" in w.get("fontname", "") for w in c) > len(c) / 2) for c in chunks]})
            continue
        blk = {"t": "text", "x0": ln["x0"], "x1": ln["x1"], "top": ln["top"], "bottom": ln["bottom"],
               "text": fix_text(ln["text"]), "size": ln["size"], "bold": ln["bold"], "center": centered}
        prev = out[-1] if out else None
        if (
            KEEP_LINES is False
            and prev is not None
            and prev["t"] == "text"
            and not prev["center"]
            and not centered
            and prev["bold"] == blk["bold"]
            and prev["size"] == blk["size"]
            and prev["last_x1"] > right_margin - 45  # previous line ran to the margin -> wrapped line
            and 0 <= blk["top"] - prev["bottom"] < blk["size"] * 0.9
            and blk["x0"] <= prev["x0"] + 4
            and not LIST_START.match(blk["text"])
        ):
            prev["text"] = join_lines([prev["text"], blk["text"]])
            prev["bottom"] = blk["bottom"]
            prev["last_x1"] = blk["x1"]
            prev["x1"] = max(prev["x1"], blk["x1"])
            continue
        blk["last_x1"] = ln["x1"]
        out.append(blk)
    return out


def parse_page(pg, kind):
    tables = [t for t in pg.find_tables() if (t.bbox[3] - t.bbox[1]) > 15 and len(t.rows) >= 1]
    if kind == "road":
        # simple Item/Description/Unit/Rate grids: the line-based reader handles them better
        all_lines = group_lines(pg.extract_words(x_tolerance=1.5, y_tolerance=3))
        hdr_ln, _ = road_header(all_lines)
        if hdr_ln is not None:
            tables = [t for t in tables if t.bbox[3] < hdr_ln["top"]]
    words = pg.extract_words(x_tolerance=1.5, y_tolerance=3, return_chars=True)
    for w in words:
        chs = w.pop("chars")
        w["fontname"] = Counter(c["fontname"] for c in chs).most_common(1)[0][0]
        w["size"] = Counter(round(c["size"], 1) for c in chs).most_common(1)[0][0]
    free = [w for w in words if not any(inside(w, t.bbox) for t in tables)]
    lines = group_lines(free)
    footer = [ln for ln in lines if is_footer(ln, pg.height, kind)]
    lines = [ln for ln in lines if ln not in footer]
    footer_text = " ".join(ln["text"] for ln in footer)
    all_x0 = [w["x0"] for w in words] or [50]
    all_x1 = [w["x1"] for w in words] or [pg.width - 50]
    left = min(min(all_x0), min((t.bbox[0] for t in tables), default=999))
    right = max(max(all_x1), max((t.bbox[2] for t in tables), default=0))

    blocks = [table_block(t, [w for w in words if inside(w, t.bbox)]) for t in tables]
    # words inside a table frame but in no cell (e.g. un-ruled rows below the grid) are free text
    leftover = [w for w in words if not w.get("_used") and any(inside(w, t.bbox) for t in tables)]
    if leftover:
        lines = sorted(lines + [ln for ln in group_lines(leftover) if not is_footer(ln, pg.height, kind)],
                       key=lambda l: l["top"])
    if kind == "road":
        hdr, bounds = road_header(lines)
        if bounds:
            ROAD_STATE["bounds"] = bounds
        elif ROAD_STATE.get("bounds") and ROAD_STATE.get("last") == pg.page_number - 1 and lines and any(
            looks_like_rate(w["text"]) or NUM_RE.match(w["text"])
            for ln in lines for w in ln["words"] if w["x0"] >= ROAD_STATE["bounds"]["rate"] - 10
        ):
            # continuation page printed without the table header
            bounds = ROAD_STATE["bounds"]
            first = min(lines, key=lambda l: l["top"])
            hdr = {"top": first["top"] - 3, "x1": right}
        if bounds:
            ROAD_STATE["last"] = pg.page_number
            body = [ln for ln in lines if ln["top"] > hdr["top"] + 2]
            # stray second header line ("No.")
            body = [ln for ln in body if not (ln["top"] - hdr["top"] < 14 and ln["text"].strip() in ("No.", "(Rs.)"))]
            above = [ln for ln in lines if ln["top"] < hdr["top"] - 2]
            titles = [ln for ln in body if ln["size"] >= 13]
            body = [ln for ln in body if ln["size"] < 13]
            above += titles
            if titles and not body and "words" in hdr:
                above.append(hdr)  # title page: keep the printed table header as a line
            if body:
                blocks.append(road_block(hdr, body, bounds, left, max(right, hdr["x1"] + 8)))
            lines = above
    rmargin = max((ln["x1"] for ln in lines), default=right)
    blocks += text_blocks(lines, pg.width, rmargin)
    # pictures (logos, photos, diagrams): cropped from the page later and placed at the same spot
    page_area = pg.width * pg.height
    for im in pg.images:
        x0, t, x1, bt = max(im["x0"], 0), max(im["top"], 0), min(im["x1"], pg.width), min(im["bottom"], pg.height)
        if x1 - x0 < 15 or bt - t < 15:
            continue
        if (x1 - x0) * (bt - t) > 0.5 * page_area and len(words) > 60:
            continue  # background / watermark behind text
        if t > pg.height * 0.9:
            continue  # footer decoration
        bb = (x0, t, x1, bt)
        tb = next((b for b in blocks if b["t"] == "table" and b.get("bbox")
                   and b["bbox"][1] - 2 <= (t + bt) / 2 <= b["bbox"][3] + 2), None)
        if tb is not None:
            tb.setdefault("imgs", []).append(bb)
        else:
            blocks.append({"t": "img", "top": t, "bbox": bb})
    blocks.sort(key=lambda b: b["top"])
    top_lines = [ln["text"] for ln in sorted(lines, key=lambda l: l["top"])[:8]]
    footer_blk = None
    if footer:
        fl = sorted(footer, key=lambda l: l["top"])
        footer_blk = {"t": "footer", "chunks": [(c[0]["x0"], fix_text(" ".join(w["text"] for w in c)))
                                                 for ln in fl for c in split_gaps(ln["words"], 25)]}
    hdr_words = {"Item", "No.", "Description", "Unit", "Rate", "S.No.", "(Rs.)"}
    body_words = sum(1 for ln in lines for w in ln["words"] if w["text"] not in hdr_words) + sum(
        1 for b in blocks if b["t"] == "table" for c in b["cells"]
        if c[4] and c[4] not in ("Item No.", "Description", "Unit", "Rate"))
    img_area = sum(max(0, im["x1"] - im["x0"]) * max(0, im["bottom"] - im["top"]) for im in pg.images)
    page_area = pg.width * pg.height
    # scanned page (no text layer) or a picture cover page with only a few words
    scan = (body_words == 0 and img_area > 0.1 * page_area) or (body_words < 40 and img_area > 0.7 * page_area)
    return {"blocks": blocks, "footer": footer_blk, "footer_text": footer_text, "top_lines": top_lines,
            "left": left, "right": right, "width": pg.width, "height": pg.height, "scan": scan,
            "landscape": pg.width > pg.height}


# ---------------------------------------------------------------- grid
def cluster(xs, tol=4.0):
    xs = sorted(xs)
    groups = []
    for x in xs:
        if groups and x - groups[-1][0] <= tol:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [sum(g) / len(g) for g in groups]


def cluster_counts(xs, tol=4.0, min_n=3):
    """Centres of x clusters that occur at least min_n times."""
    xs = sorted(xs)
    groups = []
    for x in xs:
        if groups and x - groups[-1][0] <= tol:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [min(g) for g in groups if len(g) >= min_n]


class Grid:
    def __init__(self, pages):
        xs = []
        for p in pages:
            for b in p["blocks"]:
                if b["t"] == "table":
                    for c in b["cells"]:
                        xs += [c[2], c[3]]
        # repeated start positions of gap-separated text (un-ruled tables) become columns too
        starts = []
        for p in pages:
            for b in p["blocks"]:
                if b["t"] == "chunks":
                    starts += [x for x, *_ in b["chunks"]]
        for c in cluster_counts(starts):
            xs.append(c)
        lefts = [p["left"] for p in pages]
        rights = [p["right"] for p in pages]
        self.left = min(lefts) if lefts else 40
        self.right = max(rights) if rights else 550
        xs += [self.left, self.right]
        edges = cluster(xs)
        # drop slivers narrower than 6pt
        clean = [edges[0]]
        for x in edges[1:]:
            if x - clean[-1] >= 6:
                clean.append(x)
            else:
                clean[-1] = max(clean[-1], x) if x == edges[-1] else clean[-1]
        if len(clean) < 2:
            clean = [self.left, self.right]
        self.edges = clean
        self.ncols = len(clean) - 1

    def col_at(self, x):
        """Grid column (1-based) whose span contains x (snaps to nearest edge)."""
        best = 0
        for i, e in enumerate(self.edges[:-1]):
            if e <= x + 4:
                best = i
        return best + 1

    def col_end(self, x):
        """Last grid column ending at/near x."""
        best = 1
        for i, e in enumerate(self.edges[1:], start=1):
            if e <= x + 4:
                best = i
        return max(best, 1)

    def width_pt(self, c0, c1):
        return self.edges[c1] - self.edges[c0 - 1]


# ---------------------------------------------------------------- writer
_FONTS = {}
_FONT_FILES = {
    ("serif", False): "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    ("serif", True): "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
    ("sans", False): "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ("sans", True): "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
}


def _font(size, bold):
    fam = "serif" if "Times" in FONT_NAME else "sans"
    key = (fam, bold, size)
    if key not in _FONTS:
        try:
            from PIL import ImageFont
            _FONTS[key] = ImageFont.truetype(_FONT_FILES[(fam, bold)], max(1, round(size * 96 / 72 * 4)))
        except Exception:
            _FONTS[key] = None
    return _FONTS[key]


def lines_needed(text, width_pt, size, bold=False):
    """Number of wrapped lines Excel shows for text in a cell width_pt wide.

    Measures the words with a metric-compatible font (Liberation = Times New
    Roman / Arial metrics) and wraps them greedily, like Excel does."""
    if not text:
        return 1
    avail = (width_pt / 0.75 - 9) * 4 / 1.06  # px at 4x scale, minus padding, +6% for Excel
    f = _font(size, bold)
    if f is None:
        cpl = max((width_pt - 4) / (size * 0.45), 1)
        return sum(max(1, math.ceil(len(p) / cpl)) for p in str(text).split("\n"))
    space = f.getlength(" ")
    n = 0
    for part in str(text).split("\n"):
        n += 1
        cur = 0.0
        for word in part.split():
            wl = f.getlength(word)
            if cur and cur + space + wl > avail:
                n += 1
                cur = 0.0
            if wl > avail:  # very long word breaks across lines
                extra = int(wl // avail)
                n += extra
                cur = wl - extra * avail
            else:
                cur = cur + space + wl if cur else wl
    return n


MIN_ROW = 13.5


def row_h(lines, size):
    return lines * size * 1.26 + 3


class SheetWriter:
    def __init__(self, ws, title, pages):
        self.ws = ws
        self.grid = Grid(pages)
        self.pages = pages
        self.r = 1
        self.need = {}  # row -> required height (pt)
        self.page_rows = []  # (first row, last row) of each PDF page
        self.landscape = False
        self.footer_rows = True
        self.page_r0 = 1
        self.vneed = []  # (r0, r1, height) for vertically merged cells
        g = self.grid
        for c in range(1, g.ncols + 1):
            pt = g.width_pt(c, c)
            ws.column_dimensions[get_column_letter(c)].width = max(round(pt / 0.75 / 7, 2), 0.5)  # OOXML width includes padding

    def _merge(self, r0, r1, c0, c1):
        if r1 > r0 or c1 > c0:
            self.ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)

    def _height(self, r, h):
        self.need[r] = max(self.need.get(r, MIN_ROW), h)

    def text(self, b):
        g = self.grid
        size = min(max(b["size"], 8), 18)
        if b["center"]:
            c0, c1 = 1, g.ncols
            al = Alignment(horizontal="center", vertical="top", wrap_text=True)
        else:
            c0, c1 = g.col_at(b["x0"]), g.ncols
            indent = max(0, min(15, round((b["x0"] - g.edges[c0 - 1]) / 9)))
            al = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=indent)
        c = self.ws.cell(self.r, c0, b["text"])
        c.font = font(size, b["bold"])
        c.alignment = al
        self._merge(self.r, self.r, c0, c1)
        self._height(self.r, row_h(lines_needed(b["text"], g.width_pt(c0, c1), size, b["bold"]), size))
        self.r += 1

    def chunks(self, chunks, small=False, color=None):
        g = self.grid
        cols = [g.col_at(x) for x, *_ in chunks]
        # keep strictly increasing columns
        for i in range(1, len(cols)):
            if cols[i] <= cols[i - 1]:
                cols[i] = min(cols[i - 1] + 1, g.ncols)
        h = MIN_ROW
        dec_row = sum(bool(re.fullmatch(r"\d+\.\d+", c[1].strip())) for c in chunks) >= 3 and not small
        for i, ch in enumerate(chunks):
            c0 = cols[i]
            if i and c0 <= cols[i - 1]:
                # out of columns: append to previous cell
                prev = self.ws.cell(self.r, cols[i - 1])
                prev.value = f"{prev.value}   {ch[1]}"
                continue
            c1 = (cols[i + 1] - 1) if i + 1 < len(chunks) else g.ncols
            c1 = max(c1, c0)
            bold = ch[2] if len(ch) > 2 else False
            is_dec = bool(re.fullmatch(r"\d+\.\d+", ch[1].strip()))
            cell = self.ws.cell(self.r, c0, to_number(ch[1]) if (looks_like_rate(ch[1]) or (is_dec and dec_row))
                                else ch[1])
            if isinstance(cell.value, (int, float)):
                cell.number_format = RUPEE_FMT
            cell.font = font(8 if small else 10, bold, color)
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="right" if (small and i == len(chunks) - 1 and i) else "left")
            self._merge(self.r, self.r, c0, c1)
            h = max(h, row_h(lines_needed(ch[1], g.width_pt(c0, c1), 10, bold), 10))
        self._height(self.r, h)
        self.r += 1

    def table(self, b):
        g = self.grid
        base = self.r
        hdr = b["hdr"]
        occupied = {}
        for (r0, r1, x0, x1, txt) in b["cells"]:
            c0, c1 = g.col_at(x0), g.col_end(x1)
            if c1 < c0:
                c1 = c0
            R0, R1 = base + r0, base + r1
            if any((rr, cc) in occupied for rr in range(R0, R1 + 1) for cc in range(c0, c1 + 1)):
                # overlapping after snapping: never lose text
                if (R0, c0) in occupied:
                    if txt:
                        anchor = occupied[(R0, c0)]
                        cell = self.ws.cell(*anchor)
                        cell.value = f"{cell.value} {txt}" if cell.value not in (None, "") else (
                            to_number(txt) if looks_like_rate(txt) else txt)
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = RUPEE_FMT
                    continue
                c1, R1 = c0, R0
            for rr in range(R0, R1 + 1):
                for cc in range(c0, c1 + 1):
                    occupied[(rr, cc)] = (R0, c0)
                    self.ws.cell(rr, cc).border = BORDER
            is_hdr = r0 < hdr
            val = txt
            if txt and not is_hdr and ((b["numx"] is not None and x0 >= b["numx"]) or looks_like_rate(txt)):
                val = to_number(txt)
            cell = self.ws.cell(R0, c0, val)
            cell.border = BORDER
            if isinstance(val, (int, float)):
                cell.number_format = RUPEE_FMT
            cell.font = font(10, is_hdr)
            if is_hdr:
                cell.fill = FILL_H
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            elif isinstance(val, (int, float)):
                cell.alignment = Alignment(vertical="top", horizontal="right")
            else:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if is_hdr:
                for rr in range(R0, R1 + 1):
                    for cc in range(c0, c1 + 1):
                        self.ws.cell(rr, cc).fill = FILL_H
            self._merge(R0, R1, c0, c1)
            h = row_h(lines_needed(txt, g.width_pt(c0, c1), 10, is_hdr), 10)
            if R1 == R0:
                self._height(R0, h)
            else:
                self.vneed.append((R0, R1, h))
        for (x0, t, x1, bt), path in zip(b.get("imgs", []), b.get("paths") or []):
            tops = b["row_tops"]
            ri = max([i for i, tp in enumerate(tops[:-1]) if tp <= t + 2] or [0])
            col = g.col_at(x0)
            self._place(path, base + ri, col, x0 - g.edges[col - 1], x1 - x0, bt - t)
            # rows below the picture's top must be at least as tall as the picture
            r_end = ri
            while r_end + 1 < len(tops) - 1 and tops[r_end + 1] < bt - 2:
                r_end += 1
            self.vneed.append((base + ri, base + r_end, (bt - t) + (t - tops[ri]) + 2))
        self.r = base + b["nrows"]

    def _place(self, path, row, col, x_off_pt, w_pt, h_pt):
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        img = XLImage(path)
        emu = 12700
        img.anchor = OneCellAnchor(_from=AnchorMarker(col=col - 1, colOff=int(max(x_off_pt, 0) * emu),
                                                      row=row - 1, rowOff=0),
                                   ext=XDRPositiveSize2D(int(w_pt * emu), int(h_pt * emu)))
        img.width, img.height = int(w_pt / 0.75), int(h_pt / 0.75)
        self.ws.add_image(img)

    def picture(self, b):
        """Free-standing picture: own rows, same width/height and left position as in the PDF."""
        x0, t, x1, bt = b["bbox"]
        g = self.grid
        col = g.col_at(x0)
        self._place(b["path"], self.r, col, x0 - g.edges[col - 1], x1 - x0, bt - t)
        h = bt - t
        while h > 0:
            step = min(h, 30)
            self._height(self.r, max(step, MIN_ROW))
            h -= step
            self.r += 1

    def image(self, path):
        """Place a picture of a scanned PDF page, sized to the sheet width."""
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(path)
        g = self.grid
        width_px = sum((self.ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7
                       for c in range(1, g.ncols + 1))
        ratio = img.height / img.width
        img.width = int(width_px * 0.97)
        img.height = int(img.width * ratio)
        self.ws.add_image(img, f"A{self.r}")
        total_pt = img.height * 0.75
        n = math.ceil(total_pt / 20)
        for _ in range(n):
            self._height(self.r, 20)
            self.r += 1

    def page_end(self, footer):
        if footer and self.footer_rows:
            txt = "     ".join(t for _, t in footer["chunks"])
            c = self.ws.cell(self.r, 1, txt)
            c.font = font(8, False, "595959")
            c.alignment = Alignment(horizontal="center", vertical="top")
            self._merge(self.r, self.r, 1, self.grid.ncols)
            self._height(self.r, 14)
            self.r += 1
        self.ws.row_breaks.append(Break(id=self.r - 1))
        self.page_rows.append((self.page_r0, self.r - 1))
        self.page_r0 = self.r

    def finish(self, title):
        ws = self.ws
        for r0, r1, h in self.vneed:
            have = sum(self.need.get(r, MIN_ROW) for r in range(r0, r1 + 1))
            if h > have:
                self.need[r1] = self.need.get(r1, MIN_ROW) + (h - have)
        for r, h in self.need.items():
            ws.row_dimensions[r].height = min(round(h, 1), 409)
        ws.sheet_view.showGridLines = False
        ws.print_title_rows = None
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_format.defaultRowHeight = MIN_ROW
        ws.sheet_format.customHeight = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.5
        # one print scale per sheet so that every PDF page fits on one printed page
        g = self.grid
        width_pt = sum(((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7) * 0.75
                       for c in range(1, g.ncols + 1))
        pw, ph = (842, 595) if self.landscape else (595, 842)
        ws.page_setup.orientation = "landscape" if self.landscape else "portrait"
        scale = min(1.0, (pw - 0.8 * 72) / width_pt)
        avail = ph - 72 - 18
        for r0, r1 in self.page_rows:
            h = sum(min(self.need.get(r, MIN_ROW), 409) for r in range(r0, r1 + 1))
            if h * scale > avail:
                scale = avail / h
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        ws.page_setup.scale = max(10, int(scale * 100 * 0.95))  # 5% slack for font metric differences
        ws.print_area = f"A1:{get_column_letter(g.ncols)}{max(ws.max_row, 1)}"
        ws.print_options.horizontalCentered = True



def crop(src, page_no, bbox, dpi=200):
    """Crop a region of a PDF page to a PNG file (exact picture as printed)."""
    import subprocess
    import tempfile
    x0, t, x1, bt = bbox
    k = dpi / 72
    out = os.path.join(tempfile.mkdtemp(prefix="sorimg_"), "c")
    subprocess.run(["pdftoppm", "-f", str(page_no), "-l", str(page_no), "-r", str(dpi), "-png", "-singlefile",
                    "-x", str(int(x0 * k)), "-y", str(int(t * k)), "-W", str(max(1, int((x1 - x0) * k))),
                    "-H", str(max(1, int((bt - t) * k))), src, out], check=True)
    return out + ".png"


def _footer_parts(p):
    if not p.get("footer"):
        return None
    texts = [t for _, t in p["footer"]["chunks"] if t and t.strip()]
    if not texts:
        return None
    m = re.search(r"(\d+)\s*$", texts[-1])
    if not m:
        return None
    last = texts[-1][: m.start()] + "&P" + texts[-1][m.end():]
    return (tuple(texts[:-1]), last, int(m.group(1)))


def footer_template(pages):
    """If every page footer is the same text with a page number that rises by one per PDF page,
    return (left texts, right text with &P, first printed number); else None."""
    parts = [_footer_parts(p) for p in pages]
    if not parts or any(x is None for x in parts):
        return None
    if len({(x[0], x[1]) for x in parts}) != 1:
        return None
    offs = {p["no"] - x[2] for p, x in zip(pages, parts)}
    nos = [p["no"] for p in pages]
    if len(offs) != 1 or nos != list(range(nos[0], nos[0] + len(nos))):
        return None
    return parts[0][0], parts[0][1], parts[0][2]


def _hf(part, text, size=8):
    part.text = text
    part.size = size
    part.font = f"{FONT_NAME},Regular"


def set_header_footer(ws, title, pages, tpl):
    """Page numbers in the Excel header and footer, matching the PDF."""
    nos = [p["no"] for p in pages]
    contiguous = nos == list(range(nos[0], nos[0] + len(nos)))
    esc = lambda t: t.replace("&", "&&")
    if tpl is not None:
        left, right, first = tpl
        _hf(ws.oddFooter.left, esc("     ".join(left)))
        _hf(ws.oddFooter.right, esc(right).replace("&&P", "&P"))
        ws.page_setup.firstPageNumber = first
        ws.page_setup.useFirstPageNumber = True
        _hf(ws.oddHeader.right, "Page &P")
    else:
        if contiguous:
            ws.page_setup.firstPageNumber = nos[0]
            ws.page_setup.useFirstPageNumber = True
            _hf(ws.oddHeader.right, "PDF Page &P")
            _hf(ws.oddFooter.right, "PDF Page &P")
        else:
            _hf(ws.oddHeader.right, "Page &P")
            _hf(ws.oddFooter.right, "Page &P")
    _hf(ws.oddHeader.left, esc(title))


# ====================================================================== PDF-exact layout
# Every PDF page is rebuilt on the sheet with the same geometry:
#   * row boundaries = y-positions of text lines, table rows, pictures and ruled lines
#   * column boundaries = x-positions of table lines, text starts, numbers, pictures
#   * row heights / column widths are the PDF distances (points), so printing the sheet
#     (or saving it as PDF) gives the same page.
PAPERS = {9: (595.3, 841.9), 1: (612.0, 792.0)}  # A4, Letter


def font_of(fontname, size, color=None):
    fn = fontname or ""
    low = fn.lower()
    if "times" in low:
        name = "Times New Roman"
    elif "calibri" in low:
        name = "Calibri"
    elif "cambria" in low:
        name = "Cambria"
    elif "verdana" in low:
        name = "Verdana"
    elif "tahoma" in low:
        name = "Tahoma"
    elif "bookman" in low:
        name = "Bookman Old Style"
    elif "courier" in low:
        name = "Courier New"
    else:
        name = "Arial"
    name = "Times New Roman"  # one font everywhere (as asked); size / bold / italic follow the PDF
    bold = "bold" in low or "black" in low or "heavy" in low
    italic = "italic" in low or "oblique" in low
    kw = {"name": name, "size": max(4.0, round(size * 2) / 2), "bold": bold, "italic": italic}
    if color:
        kw["color"] = color
    return Font(**kw)


def rgb_hex(c):
    """pdfplumber colour -> 'RRGGBB' (None for black / unknown)."""
    try:
        if c is None:
            return None
        if isinstance(c, (int, float)):
            c = (c, c, c)
        c = tuple(c)
        if len(c) == 1:
            c = (c[0],) * 3
        if len(c) == 4:  # CMYK
            k = c[3]
            c = tuple((1 - min(1, x + k)) for x in c[:3])
        if len(c) != 3:
            return None
        r, g, b = (max(0, min(255, int(round(x * 255)))) for x in c)
        if r < 40 and g < 40 and b < 40:
            return None
        return f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def word_info(words):
    """Dominant font name / size / colour of a group of words."""
    fn = Counter(w.get("fontname", "") for w in words).most_common(1)[0][0]
    sz = Counter(round(w.get("size", 10), 1) for w in words).most_common(1)[0][0]
    col = Counter(w.get("color") for w in words).most_common(1)[0][0]
    return fn, sz, col


def get_words(pg):
    words = pg.extract_words(x_tolerance=1.5, y_tolerance=3, return_chars=True)
    for w in words:
        chs = w.pop("chars")
        w["fontname"] = Counter(c["fontname"] for c in chs).most_common(1)[0][0]
        w["size"] = Counter(round(c["size"], 1) for c in chs).most_common(1)[0][0]
        w["color"] = rgb_hex(Counter(str(c.get("non_stroking_color")) for c in chs).most_common(1)[0][0]
                             and chs[0].get("non_stroking_color"))
    return words


def parse_abs(pg, kind):
    """Collect everything on a PDF page with its position."""
    words = get_words(pg)
    tables = [t for t in pg.find_tables() if (t.bbox[3] - t.bbox[1]) > 15 and t.rows]
    if kind == "road":
        hl, _ = road_header(group_lines(words))
        if hl is not None:  # the road item list is un-ruled: read it as text lines
            tables = [t for t in tables if t.bbox[3] < hl["top"]]

    if words:  # one text size per scanned page (OCR box heights vary a lot)
        hs_ = sorted(w["size"] for w in words)
        med = hs_[len(hs_) // 2]
        for w in words:
            w["size"] = med
    # a partly detected grid (faint / warped rules) would put text in wrong cells: use it only
    # when it holds most of the page's text
    if tables and words:
        inside_n = sum(1 for w in words if any(inside(w, t.bbox) for t in tables))
        if inside_n < 0.7 * len(words):
            tables = []
    tabs = []
    for t in tables:
        rows = t.rows
        tops = [r.bbox[1] for r in rows] + [rows[-1].bbox[3]]
        tw = [w for w in words if inside(w, t.bbox)]
        cells = []
        for ri, row in enumerate(rows):
            for bb in row.cells:
                if bb is None:
                    continue
                r1 = ri
                while r1 + 1 < len(rows) and tops[r1 + 1] < bb[3] - 2:
                    r1 += 1
                inner = [w for w in tw if inside(w, bb) and not w.get("_used")]
                for w in inner:
                    w["_used"] = True
                cells.append({"bb": bb, "r0": ri, "r1": r1, "words": inner})
        # words sitting on a cell border -> nearest cell
        for w in tw:
            if w.get("_used"):
                continue
            cx, cy = (w["x0"] + w["x1"]) / 2, (w["top"] + w["bottom"]) / 2

            def dist(c):
                x0, y0, x1, y1 = c["bb"]
                return max(x0 - cx, 0, cx - x1) + max(y0 - cy, 0, cy - y1)

            if cells:
                c = min(cells, key=dist)
                if dist(c) <= 8:
                    c["words"].append(w)
                    w["_used"] = True
        # columns under a "Rate" / "Rs" header hold rates
        rate_x = []
        for c in cells:
            if c["r0"] <= 3 and c["words"]:
                txt = " ".join(w["text"] for w in c["words"])
                if re.search(r"\b(Rate|Rates|Rs\.?|Amount)\b", txt, re.I) and len(txt) < 60:
                    rate_x.append((c["bb"][0] - 1, c["bb"][2] + 1))
        # continuation page without the header: same table layout as before -> same rate columns
        prev = ROAD_STATE.get("rate_tab")
        if rate_x:
            ROAD_STATE["rate_tab"] = (t.bbox[0], t.bbox[2], rate_x, pg.page_number)
        elif prev and abs(prev[0] - t.bbox[0]) < 6 and abs(prev[1] - t.bbox[2]) < 6 and \
                pg.page_number - prev[3] <= 3:
            rate_x = prev[2]
            ROAD_STATE["rate_tab"] = (prev[0], prev[1], prev[2], pg.page_number)
        tabs.append({"bbox": t.bbox, "tops": tops, "cells": cells, "rate_x": rate_x})

    free = [w for w in words if not w.get("_used")]
    lines = group_lines(free)
    footer = [ln for ln in lines if is_footer(ln, pg.height, kind)]
    body = [ln for ln in lines if ln not in footer]

    bounds = None
    if kind == "road":
        hl, bounds = road_header(body)
        if bounds:
            ROAD_STATE["bounds"] = bounds
            ROAD_STATE["last"] = pg.page_number
        elif ROAD_STATE.get("last") == pg.page_number - 1:
            bounds = ROAD_STATE["bounds"]
            ROAD_STATE["last"] = pg.page_number

    if bounds:
        # a second text column inside the description zone (Item Name | Description):
        # word starts after a gap, repeated on many lines at the same x
        from collections import Counter as _C
        st = _C()
        for ln in body:
            ws_ = ln["words"]
            for a_, b_ in zip(ws_, ws_[1:]):
                if b_["x0"] - a_["x1"] >= 7 and bounds["item"] + 30 < b_["x0"] < bounds["item"] + 160:
                    st[round(b_["x0"] / 3) * 3] += 1
        best = [x for x, n in st.items() if n >= 4]
        bounds = dict(bounds)
        bounds["col2"] = min(best) if best else None

    def chunks_of(ln):
        ws = ln["words"]
        out = split_table_line(ws)
        if bounds:  # also cut at the road column boundaries
            cuts = [bounds["item"], bounds["unit"], bounds["rate"]] + ([bounds["col2"] - 2] if bounds.get("col2") else [])
            res = []
            for ch in out:
                cur = [ch[0]]
                for a, b in zip(ch, ch[1:]):
                    if any(a["x1"] <= c <= b["x0"] + 2 or ((a["x0"] + a["x1"]) / 2 < c <= (b["x0"] + b["x1"]) / 2)
                           for c in cuts):
                        res.append(cur)
                        cur = [b]
                    else:
                        cur.append(b)
                res.append(cur)
            out = res
        return out

    def mk_lines(lns):
        res = []
        for ln in lns:
            chs = []
            for ch in chunks_of(ln):
                fn, sz, col = word_info(ch)
                chs.append({"x0": ch[0]["x0"], "x1": max(w["x1"] for w in ch),
                            "text": fix_text(" ".join(w["text"] for w in ch)), "font": fn, "size": sz,
                            "color": col})
            res.append({"top": ln["top"], "bottom": ln["bottom"], "chunks": chs})
        return res

    # pictures
    page_area = pg.width * pg.height
    imgs = []
    for im in pg.images:
        x0, t, x1, bt = max(im["x0"], 0), max(im["top"], 0), min(im["x1"], pg.width), min(im["bottom"], pg.height)
        if x1 - x0 < 8 or bt - t < 8:
            continue
        if (x1 - x0) * (bt - t) > 0.6 * page_area and len(words) > 60:
            continue  # background behind text
        imgs.append((x0, t, x1, bt))

    # ruled lines outside tables (header rules, footer bars, underlines of titles)
    rules = []
    segs = [(l["x0"], l["x1"], (l["top"] + l["bottom"]) / 2, max(l.get("linewidth") or 0.5, 0.5),
             l.get("stroking_color")) for l in pg.lines if abs(l["bottom"] - l["top"]) < 1.5]
    segs += [(r["x0"], r["x1"], r["top"], r["bottom"] - r["top"], r.get("non_stroking_color"))
             for r in pg.rects if (r["bottom"] - r["top"]) <= 8 and r.get("fill", True)]
    for x0, x1, y, lw, col in segs:
        if x1 - x0 < 20 or rgb_hex(col) == "FFFFFF":
            continue
        if any(t["bbox"][0] - 2 <= (x0 + x1) / 2 <= t["bbox"][2] + 2 and t["bbox"][1] - 2 <= y <= t["bbox"][3] + 2
               for t in tabs):
            continue
        rules.append((x0, x1, y, lw, rgb_hex(col)))
    # merge rules that continue each other on the same y
    rules.sort(key=lambda r: (round(r[2]), r[0]))
    merged = []
    for r in rules:
        if merged and abs(merged[-1][2] - r[2]) < 1.5 and r[0] <= merged[-1][1] + 3:
            m = merged[-1]
            merged[-1] = (m[0], max(m[1], r[1]), m[2], max(m[3], r[3]), m[4] or r[4])
        else:
            merged.append(r)

    # vertical bars / lines (sides of page frames, column rules of un-ruled tables)
    vrules = []
    vs = [(r["top"], r["bottom"], (r["x0"] + r["x1"]) / 2, r["x1"] - r["x0"], r.get("non_stroking_color"))
          for r in pg.rects if (r["x1"] - r["x0"]) <= 4 and (r["bottom"] - r["top"]) >= 20]
    vs += [(l["top"], l["bottom"], l["x0"], max(l.get("linewidth") or 0.5, 0.5), l.get("stroking_color"))
           for l in pg.lines if abs(l["x1"] - l["x0"]) < 1.5 and (l["bottom"] - l["top"]) >= 20]
    for t, bt, x, lw, col in vs:
        if rgb_hex(col) == "FFFFFF":
            continue
        if any(tb["bbox"][0] - 2 <= x <= tb["bbox"][2] + 2 and tb["bbox"][1] - 2 <= (t + bt) / 2 <= tb["bbox"][3] + 2
               for tb in tabs):
            continue
        if vrules and abs(vrules[-1][2] - x) < 2 and abs(vrules[-1][0] - t) < 3:
            continue
        vrules.append((t, bt, x, lw, rgb_hex(col)))
    frames = []
    for r in pg.rects:
        x0, t, x1, bt = r["x0"], r["top"], r["x1"], r["bottom"]
        if bt - t <= 4 or x1 - x0 < 20 or not r.get("stroke", True):
            continue
        if any(abs(x0 - tb["bbox"][0]) < 4 and abs(t - tb["bbox"][1]) < 4 for tb in tabs):
            continue
        if any(tb["bbox"][0] - 2 <= x0 and x1 <= tb["bbox"][2] + 2 and tb["bbox"][1] - 2 <= t and bt <= tb["bbox"][3] + 2
               for tb in tabs):
            continue
        frames.append((x0, t, x1, bt, max(r.get("linewidth") or 0.5, 0.5), rgb_hex(r.get("stroking_color"))))
    hdr_words = {"Item", "No.", "Description", "Unit", "Rate", "S.No.", "(Rs.)"}
    body_words = sum(1 for ln in body for w in ln["words"] if w["text"] not in hdr_words) + sum(
        len(c["words"]) for t in tabs for c in t["cells"])
    img_area = sum((b[2] - b[0]) * (b[3] - b[1]) for b in imgs)
    scan = (body_words == 0 and img_area > 0.1 * page_area) or (body_words < 40 and img_area > 0.7 * page_area)

    footer_blk = None
    if footer:
        fl = sorted(footer, key=lambda l: l["top"])
        footer_blk = {"chunks": [(c[0]["x0"], fix_text(" ".join(w["text"] for w in c)))
                                 for ln in fl for c in split_gaps(ln["words"], 25)]}
    xs = [w["x0"] for w in words] + [t["bbox"][0] for t in tabs] + [b[0] for b in imgs]
    xe = [w["x1"] for w in words] + [t["bbox"][2] for t in tabs] + [b[2] for b in imgs]
    return {
        "w": pg.width, "h": pg.height, "landscape": pg.width > pg.height,
        "lines": attach_rates(mk_lines(body), bounds), "footer_lines": mk_lines(footer), "tables": tabs,
        "imgs": imgs, "kind": kind,
        "rules": merged, "vrules": vrules, "frames": frames, "road_bounds": bounds, "scan": scan,
        "no_text": body_words == 0, "footer": footer_blk,
        "footer_text": " ".join(ln["text"] for ln in footer),
        "top_lines": [ln["text"] for ln in sorted(body, key=lambda l: l["top"])[:8]],
        "left": min(xs) if xs else 40, "right": max(xe) if xe else pg.width - 40,
    }


class AbsGrid:
    """Column edges shared by all pages of one sheet."""

    def __init__(self, pages):
        xs = []
        for p in pages:
            for t in p["tables"]:
                for c in t["cells"]:
                    xs += [c["bb"][0], c["bb"][2]]
            for ln in p["lines"] + p["footer_lines"]:
                for ch in ln["chunks"]:
                    xs.append(ch["x0"])
                    if looks_like_rate(ch["text"]) or NUM_RE.match(ch["text"].strip()):
                        xs.append(ch["x1"])
            for b in p["imgs"]:
                xs += [b[0], b[2]]
            for r in p["rules"]:
                xs += [r[0], r[1]]
            for f in p.get("frames", []):
                xs += [f[0], f[2]]
            for v in p.get("vrules", []):
                xs.append(v[2])
            xs += [p["left"], p["right"]]
        xs = sorted(xs) or [40, 550]
        groups = []
        for x in xs:
            if groups and x - groups[-1][-1] <= 2.5:
                groups[-1].append(x)
            else:
                groups.append([x])
        edges = [min(g) if len(g) < 3 else sorted(g)[len(g) // 2] for g in groups]
        clean = [edges[0]]
        for x in edges[1:]:
            if x - clean[-1] >= 3:
                clean.append(x)
        if len(clean) < 2:
            clean.append(clean[0] + 500)
        self.edges = clean
        self.ncols = len(clean) - 1
        self.left, self.right = clean[0], clean[-1]

    def col_at(self, x):
        best = 0
        for i, e in enumerate(self.edges[:-1]):
            if e <= x + 1.5:
                best = i
        return best + 1

    def col_end(self, x):
        best = 1
        for i, e in enumerate(self.edges[1:], start=1):
            if e <= x + 1.5:
                best = i
        return max(best, 1)

    def x_of(self, col):
        return self.edges[col - 1]


def text_width_pt(text, ft):
    f = _font_named(ft.name, ft.sz or 10, bool(ft.b))
    if f is None:
        return len(text) * (ft.sz or 10) * 0.5
    return f.getlength(text) / 4 * 0.75


_FONT_FILES2 = {
    "serif": ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf"),
    "sans": ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf"),
}


def _font_named(name, size, bold):
    fam = "serif"
    key = ("n", fam, bold, round(size * 2))
    if key not in _FONTS:
        try:
            from PIL import ImageFont
            fn = _FONT_FILES2[fam][1 if bold else 0]
            _FONTS[key] = ImageFont.truetype("/usr/share/fonts/truetype/liberation/" + fn,
                                             max(1, round(size * 96 / 72 * 4)))
        except Exception:
            _FONTS[key] = None
    return _FONTS[key]


class AbsWriter:
    def __init__(self, ws, title, pages, use_footer_rows):
        self.ws, self.title, self.pages = ws, title, pages
        self.g = AbsGrid(pages)
        self.r = 1
        self.heights = {}
        self.page_rows = []
        self.use_footer_rows = use_footer_rows
        self.extra = []  # (r0, r1, need_pt) for merged cells
        for c in range(1, self.g.ncols + 1):
            pt = self.g.edges[c] - self.g.edges[c - 1]
            ws.column_dimensions[get_column_letter(c)].width = round(pt / 0.75 / 7, 3)
        tops = []
        for p in pages:
            ys = self._content_ys(p)
            if ys:
                tops.append(min(ys))
        self.mt = max(min(tops) - 2, 0) if tops else 30

    def _content_ys(self, p):
        ys = [ln["top"] for ln in p["lines"]]
        if self.use_footer_rows:
            ys += [ln["top"] for ln in p["footer_lines"]]
        for t in p["tables"]:
            ys += t["tops"]
        for b in p["imgs"]:
            ys += [b[1], b[3]]
        ys += [r[2] for r in p["rules"]]
        for f in p.get("frames", []):
            ys += [f[1], f[3]]
        for v in p.get("vrules", []):
            ys += [v[0], v[1]]
        return ys

    # ---------------------------------------------------------------- one page
    def page(self, p, scan_png=None):
        ws, g = self.ws, self.g
        r_start = self.r
        if scan_png:
            bps = [self.mt, p["h"] - self.mt]
            lines, tables, imgs, rules = [], [], [], []
        else:
            lines = p["lines"] + (p["footer_lines"] if self.use_footer_rows else [])
            tables, imgs, rules = p["tables"], p["imgs"], p["rules"]
            ys = self._content_ys(p)
            ends = [ln["bottom"] + 1 for ln in lines] + [b[3] for b in imgs] + [r[2] + 1 for r in rules] + \
                [f[3] + 1 for f in p.get("frames", [])]
            for t in tables:
                ends.append(t["tops"][-1])
            bps = sorted(set([round(y, 1) for y in ys] + [round(max(ends), 1) if ends else self.mt + 10]))
        # text printed a few points above/below its neighbour (markers, rates) shares its row
        hard = set()
        for t in tables:
            hard.update(round(v, 1) for v in t["tops"])
        for r_ in rules:
            hard.add(round(r_[2], 1))
        for f_ in p.get("frames", []):
            hard.update((round(f_[1], 1), round(f_[3], 1)))
        merged = []
        for y in bps:
            if merged and (y - merged[-1] < 1.2 or (y - merged[-1] < 5 and y not in hard and merged[-1] not in hard)):
                continue
            merged.append(y)
        bps = merged
        # spacer from the common top margin to the first item
        if bps and bps[0] - self.mt > 0.8:
            bps = [self.mt] + bps
        # Excel rows are at most 409 pt tall: split big empty gaps
        split = []
        for k, y in enumerate(bps):
            if split:
                while y - split[-1] > 200:
                    split.append(split[-1] + 200)
            split.append(y)
        bps = split
        if len(bps) < 2:
            bps.append(bps[0] + 12)
        row_of_bp = {}
        for k in range(len(bps) - 1):
            self.heights[self.r + k] = bps[k + 1] - bps[k]
            row_of_bp[k] = self.r + k
        last = self.r + len(bps) - 2

        def row(y):
            k = 0
            for i, b in enumerate(bps[:-1]):
                if b <= y + 0.75:
                    k = i
            return self.r + k

        occupied = {}

        # ---- tables
        for t in tables:
            for c in t["cells"]:
                x0, y0, x1, y1 = c["bb"]
                R0 = row(t["tops"][c["r0"]])
                R1 = max(R0, row(t["tops"][c["r1"] + 1]) - 1)
                c0, c1 = g.col_at(x0), max(g.col_at(x0), g.col_end(x1))
                while (R0, c0) in occupied and c0 < c1:  # snapped onto the neighbour: move right
                    c0 += 1
                if (R0, c0) in occupied:
                    if c["words"]:  # nowhere to go: keep the text in the neighbouring cell
                        anc = ws.cell(*occupied[(R0, c0)])
                        extra = keep_lines([ln["text"] for ln in group_lines(c["words"])])
                        anc.value = f"{anc.value}  {extra}" if anc.value not in (None, "") else extra
                    continue
                free = [cc for cc in range(c0, c1 + 1) if all((rr, cc) not in occupied for rr in range(R0, R1 + 1))]
                if len(free) != c1 - c0 + 1:
                    c1 = c0
                    if any((rr, c0) in occupied for rr in range(R0, R1 + 1)):
                        R1 = R0
                for rr in range(R0, R1 + 1):
                    for cc in range(c0, c1 + 1):
                        occupied[(rr, cc)] = (R0, c0)
                        ws.cell(rr, cc).border = BORDER
                if R1 > R0 or c1 > c0:
                    ws.merge_cells(start_row=R0, start_column=c0, end_row=R1, end_column=c1)
                cell = ws.cell(R0, c0)
                cell.border = BORDER
                wds = c["words"]
                if not wds:
                    continue
                lns = group_lines(wds)
                text = keep_lines([ln["text"] for ln in lns])
                fn, sz, col = word_info(wds)
                val = text
                cx = (x0 + x1) / 2
                in_rate_col = any(a <= cx <= b for a, b in t.get("rate_x", []))
                if looks_like_rate(text) and (in_rate_col or not t.get("rate_x") and NO_HEADER_RATES):
                    val = to_number(text)
                ft = font_of(fn, sz, col)
                width = x1 - x0
                # keep every PDF line on one Excel line: shrink the font a little if needed
                avail = width - 4
                widest = max((text_width_pt(l, ft) for l in str(text).split("\n")), default=0)
                if widest > avail > 0:
                    ft = font_of(fn, max(4, sz * avail / widest * 0.98), col)
                cell.value = val
                cell.font = ft
                if isinstance(val, (int, float)):
                    cell.number_format = rupee_fmt(text)
                ha = cell_align(lns, x0, x1, isinstance(val, (int, float)))
                vtop = min(w["top"] for w in wds) - y0
                vbot = y1 - max(w["bottom"] for w in wds)
                va = "center" if abs(vtop - vbot) < 3 and vtop > 4 else ("bottom" if vbot < vtop - 3 else "top")
                cell.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=not isinstance(val, (int, float)),
                                           shrink_to_fit=isinstance(val, (int, float)))
                need = len(lns) * (ft.sz or 10) * 1.17 + 2
                self.extra.append((R0, R1, need, cell))

        # ---- road SOR item list: the description of each item / sub-item goes in ONE cell
        #      (PDF line breaks kept), so a row copied from Item No. to Rate pastes cleanly
        done = set()
        bd = p.get("road_bounds")
        if bd and lines:
            def is_desc(ch):
                cx = (ch["x0"] + ch["x1"]) / 2
                return bd["item"] <= cx < bd["unit"]

            recs, cur = [], None
            for ln in sorted(lines, key=lambda l: l["top"]):
                if ln["top"] > p["h"] * 0.88 or max(c["size"] for c in ln["chunks"]) >= 13:
                    cur = None
                    continue
                txt = " ".join(c["text"] for c in ln["chunks"])
                if "Description" in txt and "Rate" in txt:
                    cur = None
                    continue
                desc = [c for c in ln["chunks"] if is_desc(c)]
                item = [c for c in ln["chunks"] if (c["x0"] + c["x1"]) / 2 < bd["item"]]
                bold = bool(desc) and "bold" in desc[0]["font"].lower()
                first = desc[0]["text"] if desc else ""
                marker = bool(re.match(r"^(\(?[ivxIVX]{1,5}\)|\(?[a-hA-H]\)|[A-H]\s|[1-9]\s|Note)", first)) or \
                    bool(re.fullmatch(r"\(?([ivxlIVXL]{1,6}|[a-hA-H]|\d{1,2})[.)]?", first.strip())) or \
                    bool(re.match(r"^[ivxl]{1,6}\.?\s+[A-Z(@]", first))
                has_rate = any((c["x0"] + c["x1"]) / 2 >= bd["unit"] for c in ln["chunks"])
                if cur is None or item or marker or (desc and cur["bold"] != bold and cur["desc"]) or \
                        (has_rate and cur["rate"]):
                    cur = {"lines": [], "desc": [], "bold": bold, "rate": False}
                    recs.append(cur)
                cur["rate"] = cur["rate"] or has_rate
                cur["lines"].append(ln)
                cur["desc"] += [(ln, c) for c in desc]
            # second text column inside the description zone (Item Name | Description), e.g. Ch-21
            from collections import Counter as _C
            starts = _C(round(c["x0"] / 3) * 3 for ln in lines for c in ln["chunks"]
                        if is_desc(c) and bd["item"] + 30 < c["x0"] < bd["unit"] - 60)
            col2 = bd.get("col2") or min([x for x, n in starts.items() if n >= 4 and x < bd["item"] + 160],
                                         default=None)

            def put(group, stop_x):
                by_line = {}
                for ln, c in group:
                    by_line.setdefault(id(ln), (ln, []))[1].append(c)
                ordered = sorted(by_line.values(), key=lambda v: v[0]["top"])
                parts = [" ".join(c["text"] for c in cs) for _, cs in ordered]
                R0, R1 = row(ordered[0][0]["top"]), row(ordered[-1][0]["top"])
                c0 = g.col_at(min(c["x0"] for _, c in group))
                c1 = max(c0, g.col_at(stop_x) - 1)
                span = [(rr, cc) for rr in range(R0, R1 + 1) for cc in range(c0, c1 + 1)]
                if any(sp in occupied for sp in span):
                    return False
                for sp in span:
                    occupied[sp] = (R0, c0)
                if R1 > R0 or c1 > c0:
                    ws.merge_cells(start_row=R0, start_column=c0, end_row=R1, end_column=c1)
                c = group[0][1]
                cell = ws.cell(R0, c0, keep_lines(parts))
                cell.font = font_of(c["font"], c["size"], c["color"])
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                self.extra.append((R0, R1, len(parts) * c["size"] * 1.17 + 2, cell))
                for _, c in group:
                    done.add(id(c))
                return True

            for rec in recs:
                if not rec["desc"]:
                    continue
                others = [c["x0"] for ln in rec["lines"] for c in ln["chunks"]
                          if not is_desc(c) and c["x0"] > min(cc["x1"] for _, cc in rec["desc"]) - 1]
                stop = min([bd["unit"]] + others)
                if col2 is not None:
                    name = [(ln, c) for ln, c in rec["desc"] if c["x0"] < col2 - 4]
                    desc = [(ln, c) for ln, c in rec["desc"] if c["x0"] >= col2 - 4]
                    if name and desc and max(c["x1"] for _, c in name) < col2 + 2:
                        put(name, col2)
                        put(desc, stop)
                        continue
                put(rec["desc"], stop)

        # ---- free text lines (one PDF line = one Excel row)
        for ln in lines:
            R = row(ln["top"])
            chs = ln["chunks"]
            for k, ch in enumerate(chs):
                if id(ch) in done:
                    continue
                c0 = g.col_at(ch["x0"])
                if (R, c0) in occupied:  # landed on a filled cell: next free one, else append
                    nxt = next((cc for cc in range(c0 + 1, g.ncols + 1) if (R, cc) not in occupied), None)
                    if nxt is None:
                        anc = ws.cell(*occupied[(R, c0)])
                        anc.value = f"{anc.value} {ch['text']}" if anc.value not in (None, "") else ch["text"]
                        continue
                    c0 = nxt
                txt = ch["text"]
                val = to_number(txt) if (looks_like_rate(txt) and (k > 0 or len(chs) == 1 and ch["x0"] > g.left + 0.5 * (g.right - g.left))
                                         or (k > 0 and p.get("kind") == "road" and re.fullmatch(r"\d[\d,]*\.\d+", txt.strip()))) else txt
                ft = font_of(ch["font"], ch["size"], ch["color"])
                if isinstance(val, (int, float)):
                    # numbers do not overflow: give them room up to the next text, right aligned at x1
                    c1 = g.col_end(ch["x1"])
                    left_lim = g.col_at(chs[k - 1]["x1"]) + 1 if k else 1
                    cs = max(left_lim, min(c0, c1))
                    while cs > left_lim and (R, cs - 1) not in occupied and \
                            g.x_of(c1 + 1) - g.x_of(cs) < text_width_pt("₹ " + txt, ft) + 6:
                        cs -= 1
                    span = [(R, cc) for cc in range(cs, c1 + 1)]
                    if any(s in occupied for s in span):
                        cs, c1 = c0, c0
                    for s in [(R, cc) for cc in range(cs, c1 + 1)]:
                        occupied[s] = (R, cs)
                    if c1 > cs:
                        ws.merge_cells(start_row=R, start_column=cs, end_row=R, end_column=c1)
                    cell = ws.cell(R, cs, val)
                    cell.number_format = rupee_fmt(txt)
                    cell.alignment = Alignment(horizontal="right", vertical="top", shrink_to_fit=True)
                else:
                    occupied[(R, c0)] = (R, c0)
                    cell = ws.cell(R, c0, val)
                    cell.alignment = Alignment(horizontal="left", vertical="top")
                cell.font = ft
                self.extra.append((R, R, (ft.sz or 10) * 1.05, cell))

        # ---- ruled lines
        for x0, x1, y, lw, colr in rules:
            R = row(y)
            if abs(bps[R - self.r] - y) > 1.5 and R < last:
                R = R + 1 if abs(bps[min(R - self.r + 1, len(bps) - 1)] - y) < 1.5 else R
            st = "thick" if lw >= 2.2 else ("medium" if lw >= 1.2 else "thin")
            side = Side(style=st, color=colr or "000000")
            for cc in range(g.col_at(x0), g.col_end(x1) + 1):
                cell = ws.cell(R, cc)
                b = cell.border
                cell.border = Border(left=b.left, right=b.right, top=side, bottom=b.bottom)

        # ---- vertical rules: left border of the column starting at x
        if not scan_png:
            for t, bt, x, lw, colr in p.get("vrules", []):
                st = "thick" if lw >= 2.2 else ("medium" if lw >= 1.2 else "thin")
                side = Side(style=st, color=colr or "000000")
                k = min(range(len(g.edges)), key=lambda i: abs(g.edges[i] - x))
                if abs(g.edges[k] - x) > 4:
                    continue
                c, edge = (k + 1, "left") if k < g.ncols else (g.ncols, "right")
                for rr in range(row(t), row(bt)):
                    cell = ws.cell(rr, c)
                    b = cell.border
                    kw = {"left": b.left, "right": b.right, "top": b.top, "bottom": b.bottom}
                    kw[edge] = side
                    cell.border = Border(**kw)

        # ---- frames (boxes drawn round titles / pages)
        if not scan_png:
            for x0, t, x1, bt, lw, colr in p.get("frames", []):
                st = "thick" if lw >= 2.2 else ("medium" if lw >= 1.2 else "thin")
                side = Side(style=st, color=colr or "000000")
                ra, rb = row(t), row(bt)
                ca, cb = g.col_at(x0), max(g.col_at(x0), g.col_end(x1))
                for rr in range(ra, rb):
                    for cc, edge in ((ca, "left"), (cb, "right")):
                        cell = ws.cell(rr, cc)
                        b = cell.border
                        kw = {"left": b.left, "right": b.right, "top": b.top, "bottom": b.bottom}
                        kw[edge] = side
                        cell.border = Border(**kw)
                for cc in range(ca, cb + 1):
                    for rr, edge in ((ra, "top"), (rb, "top")):
                        cell = ws.cell(rr, cc)
                        b = cell.border
                        kw = {"left": b.left, "right": b.right, "top": b.top, "bottom": b.bottom}
                        kw[edge] = side
                        cell.border = Border(**kw)

        # ---- pictures
        if scan_png:
            self._place(scan_png, self.r, 1, 0, p["w"] * 0.0 + (g.right - g.left), (g.right - g.left) * p["h"] / p["w"])
        for (x0, t, x1, bt), path in zip(imgs, p.get("img_paths", [])):
            R = row(t)
            c = g.col_at(x0)
            self._place(path, R, c, x0 - g.x_of(c), x1 - x0, bt - t, t - bps[R - self.r])

        self.r = last + 1
        self.ws.row_breaks.append(Break(id=last))
        self.page_rows.append((r_start, last))
        return r_start, last

    def _place(self, path, row, col, x_off, w_pt, h_pt, y_off=0.0):
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        img = XLImage(path)
        emu = 12700
        img.anchor = OneCellAnchor(
            _from=AnchorMarker(col=col - 1, colOff=int(max(x_off, 0) * emu), row=row - 1,
                               rowOff=int(max(y_off, 0) * emu)),
            ext=XDRPositiveSize2D(int(w_pt * emu), int(h_pt * emu)))
        self.ws.add_image(img)

    def finish(self, landscape):
        ws = self.ws
        # merged / wrapped cells must be tall enough for their lines
        for r0, r1, need, cell in self.extra:
            have = sum(self.heights.get(r, MIN_ROW) for r in range(r0, r1 + 1))
            if need > have + 0.5:
                # Excel's line height is a bit taller than the PDF's line spacing: prefer a slightly
                # smaller font (max 15 %) over taller rows, so every page keeps the PDF's height
                f = (have - 2) / max(need - 2, 1)
                ft = cell.font
                if f >= 0.85 and ft is not None and ft.sz:
                    nf = copy(ft)
                    nf.sz = max(4.0, math.floor(ft.sz * f * 2) / 2)
                    cell.font = nf
                    lines_ = (need - 2) / (ft.sz * 1.17) if ft.sz else 1
                    need = lines_ * nf.sz * 1.17 + 2
                if need > have + 0.5:
                    self.heights[r1] = self.heights.get(r1, MIN_ROW) + (need - have)
        for r, h in self.heights.items():
            ws.row_dimensions[r].height = round(max(h, 0.75), 2)
        ws.sheet_format.defaultRowHeight = MIN_ROW
        ws.sheet_format.customHeight = True
        ws.sheet_view.showGridLines = False
        # paper like the PDF: A4 or Letter, same orientation; same margins
        pw = max(p["w"] for p in self.pages)
        ph = max(p["h"] for p in self.pages)
        short, long_ = min(pw, ph), max(pw, ph)
        paper = 1 if abs(short / long_ - 612 / 792) < abs(short / long_ - 595.3 / 841.9) else 9
        ps, pl = PAPERS[paper]
        f = min(ps / short, pl / long_)  # PDF -> paper scale
        paper_w, paper_h = (pl, ps) if landscape else (ps, pl)
        ws.page_setup.paperSize = paper
        ws.page_setup.orientation = "landscape" if landscape else "portrait"
        g = self.g
        left = max(g.left * f, 14)
        content_w = (g.right - g.left) * f
        top = max(self.mt * f, 18)
        bottom = 36
        scale = f
        # never let a PDF page spill onto two printed pages
        for r0, r1 in self.page_rows:
            h = sum(self.heights.get(r, MIN_ROW) for r in range(r0, r1 + 1))
            if h * scale > paper_h - top - bottom:
                scale = (paper_h - top - bottom) / h
        if content_w / f * scale > paper_w - left - 10:
            scale = (paper_w - left - 10) / (content_w / f)
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        scale *= 0.96  # slack: Excel / LibreOffice round row heights and column widths to pixels
        ws.page_setup.scale = max(10, min(400, int(scale * 100)))
        ws.page_margins.left = round(left / 72, 3)
        ws.page_margins.right = 0.1
        ws.page_margins.top = round(top / 72, 3)
        ws.page_margins.bottom = round(bottom / 72, 3)
        ws.page_margins.header = 0.2
        ws.page_margins.footer = 0.2
        ws.print_options.horizontalCentered = False
        ws.print_area = f"A1:{get_column_letter(g.ncols)}{max(self.r - 1, 1)}"


def normalise(p):
    """Pages printed at an odd size (e.g. 913 x 1182 pt) are scaled to Letter/A4 size,
    so every page of a sheet shares one scale and looks like the PDF."""
    w, h = p["w"], p["h"]
    short, long_ = min(w, h), max(w, h)
    if short <= 625 and long_ <= 860:
        return
    k = min(612 / short, 842 / long_) if abs(short / long_ - 612 / 792) < 0.04 else min(595.3 / short, 841.9 / long_)
    if abs(short / long_ - 612 / 792) < 0.04:
        k = 612 / short
    sc = lambda v: v * k
    p["w"], p["h"], p["left"], p["right"] = w * k, h * k, p["left"] * k, p["right"] * k
    for ln in p["lines"] + p["footer_lines"]:
        ln["top"], ln["bottom"] = sc(ln["top"]), sc(ln["bottom"])
        for ch in ln["chunks"]:
            ch["x0"], ch["x1"], ch["size"] = sc(ch["x0"]), sc(ch["x1"]), ch["size"] * k
    for t in p["tables"]:
        t["bbox"] = tuple(sc(v) for v in t["bbox"])
        t["tops"] = [sc(v) for v in t["tops"]]
        t["rate_x"] = [(sc(a), sc(b)) for a, b in t.get("rate_x", [])]
        for c in t["cells"]:
            c["bb"] = tuple(sc(v) for v in c["bb"])
            for wd in c["words"]:
                for key in ("x0", "x1", "top", "bottom"):
                    wd[key] = sc(wd[key])
                wd["size"] = wd["size"] * k
    p["imgs"] = [tuple(sc(v) for v in b) for b in p["imgs"]]
    p["rules"] = [(sc(a), sc(b), sc(y), lw * k, c) for a, b, y, lw, c in p["rules"]]
    p["frames"] = [(sc(a), sc(t), sc(b), sc(bt), lw * k, c) for a, t, b, bt, lw, c in p.get("frames", [])]
    p["vrules"] = [(sc(t), sc(bt), sc(x), lw * k, c) for t, bt, x, lw, c in p.get("vrules", [])]
    if p.get("road_bounds"):
        p["road_bounds"] = {kk: sc(v) for kk, v in p["road_bounds"].items()}


def attach_rates(lines, bounds):
    """Road SOR: a line holding only Unit / Rate (printed half a line off) joins the
    description line it belongs to, so the rate sits on its item's row."""
    if not bounds:
        return lines

    def is_ur(c):
        return (c["x0"] + c["x1"]) / 2 >= bounds["unit"]

    ur = [ln for ln in lines if ln["chunks"] and all(is_ur(c) for c in ln["chunks"])]
    rest = [ln for ln in lines if ln not in ur]
    for ul in ur:
        def has_rate(ln):
            return any(is_ur(c) and NUM_RE.match(c["text"].replace(",", "").strip()) for c in ln["chunks"])
        def has_unit(ln):
            return any(is_ur(c) and not NUM_RE.match(c["text"].replace(",", "").strip()) for c in ln["chunks"])

        rate_only = all(NUM_RE.match(c["text"].replace(",", "").strip()) for c in ul["chunks"])
        near = [ln for ln in rest if abs(ln["top"] - ul["top"]) <= (20 if rate_only else 12) and not has_rate(ln)]
        # a line that already has its Unit but no rate is the rate's line; otherwise the line just below
        if rate_only and any(has_unit(ln) for ln in near):
            near = [ln for ln in near if has_unit(ln)]
        near = [ln for ln in near if abs(ln["top"] - ul["top"]) <= 12 or has_unit(ln)]
        tgt = min(near, key=lambda l: (abs(l["top"] - ul["top"]), l["top"] < ul["top"])) if near else None
        if tgt is None:
            rest.append(ul)
            continue
        tgt["chunks"] = sorted(tgt["chunks"] + ul["chunks"], key=lambda c: c["x0"])
    rest = sorted(rest, key=lambda l: l["top"])

    # numbered groups printed with every rate half a line too high:
    #   "(i) Solid Slab"  <rate a> / "1 Height .." <rate b> / ... / "4 Height .." (no rate)
    # -> each rate belongs one line lower; the heading gets none
    def marker(ln):
        d = [c for c in ln["chunks"] if bounds["item"] <= (c["x0"] + c["x1"]) / 2 < bounds["unit"]]
        m = re.match(r"^(\d{1,2})\s", d[0]["text"] + " ") if d else None
        return int(m.group(1)) if m else None

    def ur_chunks(ln):
        return [c for c in ln["chunks"] if is_ur(c)]

    i = 0
    while i < len(rest) - 1:
        h = rest[i]
        if ur_chunks(h) and marker(h) is None and marker(rest[i + 1]) == 1:
            j = i + 1
            while j + 1 < len(rest) and marker(rest[j + 1]) == marker(rest[j]) + 1:
                j += 1
            grp = rest[i:j + 1]
            if not ur_chunks(grp[-1]) and all(ur_chunks(g) for g in grp[:-1]):
                moved = [ur_chunks(g) for g in grp]
                for g in grp:
                    g["chunks"] = [c for c in g["chunks"] if not is_ur(c)]
                for k in range(1, len(grp)):
                    grp[k]["chunks"] = sorted(grp[k]["chunks"] + moved[k - 1], key=lambda c: c["x0"])
            i = j + 1
        else:
            i += 1
    return rest


def rupee_fmt(txt):
    """₹ format that shows exactly the decimals printed in the PDF (and thousands commas if used)."""
    t = txt.strip()
    dec = len(t.split(".")[1]) if "." in t else 0
    base = "#,##0" if "," in t else "0"
    return '"₹ "' + base + ("." + "0" * dec if dec else "")


def cell_align(lns, x0, x1, numeric):
    """Horizontal alignment of a PDF cell from where its lines sit."""
    if numeric:
        return "right"
    lefts = [ln["x0"] - x0 for ln in lns]
    rights = [x1 - ln["x1"] for ln in lns]
    if len(lns) >= 2:
        sl = max(lefts) - min(lefts)
        sr = max(rights) - min(rights)
        mids = [(a - b) for a, b in zip(lefts, rights)]
        if min(lefts) > 3 and max(abs(m) for m in mids) <= 6 and not (sl <= 1 and sr <= 1 and min(lefts) < 8):
            return "center"
        if sl <= 3:
            return "left"
        if sr <= 3:
            return "right"
        return "left"
    l, r = lefts[0], rights[0]
    if abs(l - r) <= 4 and l > 3:
        return "center"
    if r < l - 3 and r < 8:
        return "right"
    return "left"


# ---------------------------------------------------------------- OCR of scanned pages
_OCR = {}


def _ocr_engine():
    if "e" not in _OCR:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _OCR["e"] = RapidOCR()
        except Exception:
            print("OCR engine not available (pip install rapidocr_onnxruntime): scanned pages stay pictures")
            _OCR["e"] = None
    return _OCR["e"]


def _dark_segments(dark, minlen):
    """Long dark runs along axis 1, merged over neighbouring rows -> (start, end, centre)."""
    import numpy as np
    out, active = [], []
    for y in range(dark.shape[0]):
        d = np.diff(np.concatenate(([0], dark[y].astype(np.int8), [0])))
        st, en = np.where(d == 1)[0], np.where(d == -1)[0]
        new = []
        for a, b in zip(st, en):
            if b - a < minlen:
                continue
            hit = next((sg for sg in active if sg["y1"] >= y - 1 and min(b, sg["b"]) - max(a, sg["a"]) > 0.5 * (b - a)),
                       None)
            if hit:
                hit["y1"], hit["a"], hit["b"] = y, min(hit["a"], a), max(hit["b"], b)
                new.append(hit)
            else:
                sg = {"a": a, "b": b, "y0": y, "y1": y}
                out.append(sg)
                new.append(sg)
        active = new
    return [(sg["a"], sg["b"], (sg["y0"] + sg["y1"]) / 2) for sg in out if sg["y1"] - sg["y0"] <= 14]


def _chain(segs, gap=45.0, drift=6.0, minlen=50.0):
    """Join collinear pieces of a wavy hand-scanned rule into one straight line."""
    segs = sorted(segs, key=lambda t: t[0])
    chains = []
    for a, b, y in segs:
        best = None
        for ch in chains:
            if 0 <= a - ch["b"] <= gap or (a < ch["b"] and b > ch["b"]):
                if abs(ch["ye"] - y) <= drift and (best is None or abs(ch["ye"] - y) < abs(best["ye"] - y)):
                    best = ch
        if best is not None:
            best["b"] = max(best["b"], b)
            best["ye"] = y
            best["ys"].append(y)
        else:
            chains.append({"a": a, "b": b, "ye": y, "ys": [y]})
    out = []
    for ch in chains:
        if ch["b"] - ch["a"] >= minlen:
            ys = sorted(ch["ys"])
            out.append((ch["a"], ch["b"], ys[len(ys) // 2]))
    return out


def ocr_page(src, page_no, pg, kind, dpi=200):
    """Scanned page -> page dict like parse_abs(): table grid from the ruled lines in the
    picture, text from OCR."""
    import subprocess
    import tempfile
    import numpy as np
    from PIL import Image
    eng = _ocr_engine()
    if eng is None:
        return None
    tmp = tempfile.mkdtemp(prefix="sorocr_")
    png = os.path.join(tmp, "p.png")
    subprocess.run(["pdftoppm", "-f", str(page_no), "-l", str(page_no), "-r", str(dpi), "-png", "-singlefile",
                    src, os.path.join(tmp, "p")], check=True)
    k = 72 / dpi
    W, H = pg.width, pg.height
    # straighten photographed / tilted scans (best angle = sharpest row profile)
    img = Image.open(png).convert("L")
    small = img.resize((img.width // 4, img.height // 4))

    def _score(a):
        r = np.asarray(small.rotate(a, resample=Image.BILINEAR, fillcolor=255)) < 150
        return float((r.sum(1).astype(float) ** 2).sum())

    ang = max(np.arange(-3, 3.01, 0.25), key=_score)
    if abs(ang) >= 0.25:
        ang = max(np.arange(ang - 0.25, ang + 0.26, 0.05), key=_score)
        img = img.rotate(ang, resample=Image.BILINEAR, fillcolor=255)
        img.save(png)
    im = np.asarray(img)
    dark = im < 150
    minlen = int(25 / k)
    hs = [(a * k, b * k, y * k) for a, b, y in _dark_segments(dark, minlen)]
    vs = [(a * k, b * k, x * k) for a, b, x in _dark_segments(dark.T, minlen)]
    hs, vs = _chain(hs), _chain(vs)
    hs = [h for h in hs if h[2] < H * 0.9]
    vs = [v for v in vs if v[1] < H * 0.92]
    Hl = [{"x0": a, "x1": b, "top": y, "bottom": y, "width": b - a, "height": 0, "object_type": "line"}
          for a, b, y in hs]
    Vl = [{"x0": x, "x1": x, "top": a, "bottom": b, "width": 0, "height": b - a, "object_type": "line"}
          for a, b, x in vs]
    tables = []
    if Hl and Vl:
        try:
            tables = pg.find_tables({"vertical_strategy": "explicit", "horizontal_strategy": "explicit",
                                     "explicit_vertical_lines": Vl, "explicit_horizontal_lines": Hl,
                                     "snap_tolerance": 5, "join_tolerance": 8, "intersection_tolerance": 6})
        except Exception:
            tables = []
    res, _ = eng(png)
    words = []
    for box, txt, conf in res or []:
        xs = [pt[0] * k for pt in box]
        ys = [pt[1] * k for pt in box]
        top, bot = min(ys), max(ys)
        if top > H * 0.9 or not txt.strip():
            continue  # page footer is real text, already in the chapter sheet
        txt = txt.strip()
        m = re.fullmatch(r"[-–—~_.,:;'`\s]*(\d[\d,]*\.\d{2})[-–—~_.,:;'`\s]*", txt)
        if m:  # stray dash / dot read next to a number
            txt = m.group(1)
        words.append({"text": fix_text(txt), "x0": min(xs), "x1": max(xs), "top": top, "bottom": bot,
                      "fontname": "Times", "size": max(6.0, min(14.0, (bot - top) * 0.72)), "color": None})
    if words:  # one text size per scanned page (OCR box heights vary a lot)
        hs_ = sorted(w["size"] for w in words)
        med = hs_[len(hs_) // 2]
        for w in words:
            w["size"] = med
    # a partly detected grid (faint / warped rules) would put text in wrong cells: use it only
    # when it holds most of the page's text
    if tables and words:
        inside_n = sum(1 for w in words if any(inside(w, t.bbox) for t in tables))
        if inside_n < 0.7 * len(words):
            tables = []
    tabs = []
    for t in tables:
        rows = t.rows
        tops = [r.bbox[1] for r in rows] + [rows[-1].bbox[3]]
        cells = []
        for ri, row in enumerate(rows):
            for bb in row.cells:
                if bb is None:
                    continue
                r1 = ri
                while r1 + 1 < len(rows) and tops[r1 + 1] < bb[3] - 2:
                    r1 += 1
                inner = [w for w in words if inside(w, bb) and not w.get("_used")]
                for w in inner:
                    w["_used"] = True
                cells.append({"bb": bb, "r0": ri, "r1": r1, "words": inner})
        rate_x = []
        for c in cells:
            if c["r0"] <= 3 and c["words"]:
                txt = " ".join(w["text"] for w in c["words"])
                if re.search(r"\b(Rate|Rates|Rs\.?|Amount)", txt, re.I) and len(txt) < 60:
                    rate_x.append((c["bb"][0] - 1, c["bb"][2] + 1))
        tabs.append({"bbox": t.bbox, "tops": tops, "cells": cells, "rate_x": rate_x})
    free = [w for w in words if not w.get("_used")]
    lines = []
    for ln in group_lines(free, tol=4):
        chs = []
        for ch in split_table_line(ln["words"]):
            chs.append({"x0": ch[0]["x0"], "x1": max(w["x1"] for w in ch),
                        "text": " ".join(w["text"] for w in ch), "font": "Times", "size": ch[0]["size"],
                        "color": None})
        lines.append({"top": ln["top"], "bottom": ln["bottom"], "chunks": chs})
    xs = [w["x0"] for w in words] + [t["bbox"][0] for t in tabs] or [40]
    xe = [w["x1"] for w in words] + [t["bbox"][2] for t in tabs] or [W - 40]
    return {"w": W, "h": H, "landscape": W > H, "lines": lines, "footer_lines": [], "tables": tabs, "imgs": [],
            "img_paths": [], "rules": [], "vrules": [], "frames": [], "road_bounds": None, "scan": False,
            "footer": None, "footer_text": "", "top_lines": [], "left": min(xs), "right": max(xe), "kind": kind}


def convert(src, dst, kind):
    import subprocess
    import tempfile
    with pdfplumber.open(src) as pdf:
        chapters = Chapters(kind, pdf)
        rng = os.environ.get("PAGES")
        sel = range(*map(int, rng.split("-"))) if rng else range(len(pdf.pages))
        sheets, order = {}, []
        ocr_pages = []
        for i in sel:
            pg = pdf.pages[i]
            pg = pg.within_bbox((0, 0, pg.width, pg.height)).dedupe_chars()
            p = parse_abs(pg, kind)
            chap = chapters.page(p["footer_text"], p["top_lines"])
            p["no"] = i + 1
            if p["scan"]:
                tmp = tempfile.mkdtemp(prefix="sorpage_")
                subprocess.run(["pdftoppm", "-f", str(i + 1), "-l", str(i + 1), "-r", "150", "-png",
                                "-singlefile", src, os.path.join(tmp, "p")], check=True)
                p["image"] = os.path.join(tmp, "p.png")
                if p.get("no_text") and os.environ.get("NO_OCR") is None:
                    q = ocr_page(src, i + 1, pg, kind)
                    if q is not None:
                        q["no"] = i + 1
                        q["chap"] = chap
                        ocr_pages.append(q)
            else:
                p["img_paths"] = [crop(src, i + 1, b) for b in p["imgs"]]
            normalise(p)
            if chap not in sheets:
                sheets[chap] = []
                order.append(chap)
            sheets[chap].append(p)
            if (i + 1) % 50 == 0:
                print(f"{src}: parsed {i + 1}/{len(pdf.pages)} pages", flush=True)

    wb = Workbook()
    wb.remove(wb.active)
    # Normal style = Arial 10 so that column-width units are 7 px everywhere
    normal = [st for st in wb._named_styles if st.name == "Normal"][0]
    normal.font = Font(name="Arial", size=10)
    wb._fonts[0] = Font(name="Arial", size=10)
    used, page_map, groups = set(), [], []
    for chap in order:
        runs = []
        for p in sheets[chap]:
            if runs and runs[-1][0] == p["landscape"] and runs[-1][1][-1]["no"] == p["no"] - 1:
                runs[-1][1].append(p)
            else:
                runs.append((p["landscape"], [p]))
        for k, (land, rp) in enumerate(runs, start=1):
            if len(runs) == 1:
                name = chap
            elif land:
                name = f"{chap} (Landscape)"
            else:
                name = chap if k == 1 else f"{chap} ({k})"
            groups.append((name, chap, rp, land))
    for name, chap, pages, land in groups:
        ws = wb.create_sheet(sheet_title(name, used))
        tpl = footer_template(pages)
        w = AbsWriter(ws, chap, pages, use_footer_rows=tpl is None)
        for p in pages:
            r0, r1 = w.page(p, p.get("image"))
            page_map.append({"sheet": ws.title, "page": p["no"], "r0": r0, "r1": r1})
        w.finish(land)
        set_header_footer(ws, chap, pages, tpl)
    # scanned pages: editable copy made by OCR (the picture stays in the chapter sheet)
    if ocr_pages:
        for land in (False, True):
            pages = [q for q in ocr_pages if q["landscape"] == land]
            if not pages:
                continue
            nm = "Scanned pages - OCR text" + (" (Landscape)" if land else "")
            ws = wb.create_sheet(sheet_title(nm, used))
            w = AbsWriter(ws, nm, pages, use_footer_rows=False)
            ws.cell(1, 1)  # keep A1
            for q in pages:
                w.page(q)
            w.finish(land)
            _hf(ws.oddHeader.left, "Editable text read by OCR from scanned PDF pages - check against the picture")
            _hf(ws.oddHeader.right, "PDF page &P")
            _hf(ws.oddFooter.right, "PDF page &P")
            nos = [q["no"] for q in pages]
            if nos == list(range(nos[0], nos[0] + len(nos))):
                ws.page_setup.firstPageNumber = nos[0]
                ws.page_setup.useFirstPageNumber = True
            print(f"OCR sheet {ws.title}: {len(pages)} pages", flush=True)
    wb.save(dst)
    if os.environ.get("MAPFILE"):
        import json
        with open(os.environ["MAPFILE"], "w") as fh:
            json.dump(page_map, fh)
    print(f"saved {dst}: {len(groups)} sheets")



if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "grid")
