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
    return join_lines([ln["text"] for ln in group_lines(inner)])


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
    return {"t": "table", "cells": cells, "nrows": len(rows), "hdr": hdr, "numx": numx, "top": tbl.bbox[1]}


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
        txt = {k: (join_lines(rec[k]) if rec[k] else None) for k in ROAD_COLS}
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
            prev is not None
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
            "left": left, "right": right, "width": pg.width, "scan": scan}


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
    avail = (width_pt / 0.75 - 7) * 4  # px at 4x scale, minus cell padding
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
    return lines * size * 1.17 + 2


class SheetWriter:
    def __init__(self, ws, title, pages):
        self.ws = ws
        self.grid = Grid(pages)
        self.pages = pages
        self.r = 1
        self.need = {}  # row -> required height (pt)
        self.page_rows = []  # (first row, last row) of each PDF page
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
        self.r = base + b["nrows"]

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
        if footer:
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
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_format.defaultRowHeight = MIN_ROW
        ws.sheet_format.customHeight = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.5
        # one print scale per sheet so that every PDF page fits on one printed page
        g = self.grid
        width_pt = sum(((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7) * 0.75
                       for c in range(1, g.ncols + 1))
        scale = min(1.0, (595 - 0.8 * 72) / width_pt)
        avail = 842 - 72 - 18
        for r0, r1 in self.page_rows:
            h = sum(min(self.need.get(r, MIN_ROW), 409) for r in range(r0, r1 + 1))
            if h * scale > avail:
                scale = avail / h
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        ws.page_setup.scale = max(10, int(scale * 100 * 0.95))  # 5% slack for font metric differences
        ws.print_area = f"A1:{get_column_letter(g.ncols)}{max(ws.max_row, 1)}"
        ws.print_options.horizontalCentered = True
        ws.oddFooter.center.text = title


def convert(src, dst, kind):
    global FONT_NAME
    with pdfplumber.open(src) as pdf:
        fonts = Counter()
        for pg in pdf.pages[10:40]:
            for ch in pg.chars[:400]:
                fonts["Times New Roman" if "Times" in ch["fontname"] else "Arial"] += 1
        FONT_NAME = fonts.most_common(1)[0][0] if fonts else "Arial"
        chapters = Chapters(kind, pdf)
        rng = os.environ.get("PAGES")
        sel = range(*map(int, rng.split("-"))) if rng else range(len(pdf.pages))
        sheets = {}  # chapter -> list of parsed pages
        order = []
        for i in sel:
            pg = pdf.pages[i]
            pg = pg.within_bbox((0, 0, pg.width, pg.height)).dedupe_chars()  # drop off-page objects
            p = parse_page(pg, kind)
            chap = chapters.page(p["footer_text"], p["top_lines"])
            p["no"] = i + 1
            if p["scan"]:
                # scanned page (no text layer): keep a picture of it
                import subprocess
                import tempfile
                tmp = tempfile.mkdtemp(prefix="sorpage_")
                subprocess.run(["pdftoppm", "-f", str(i + 1), "-l", str(i + 1), "-r", "110", "-png",
                                "-singlefile", src, os.path.join(tmp, "p")], check=True)
                p["image"] = os.path.join(tmp, "p.png")
                p["blocks"] = []
            if chap not in sheets:
                sheets[chap] = []
                order.append(chap)
            sheets[chap].append(p)
            if (i + 1) % 50 == 0:
                print(f"{src}: parsed {i + 1}/{len(pdf.pages)} pages", flush=True)

    wb = Workbook()
    wb.remove(wb.active)
    # default (Normal) font = the PDF's font, so column-width units match in Excel and LibreOffice
    normal = [st for st in wb._named_styles if st.name == "Normal"][0]
    normal.font = Font(name=FONT_NAME, size=10)
    wb._fonts[0] = Font(name=FONT_NAME, size=10)
    used = set()
    page_map = []
    for chap in order:
        ws = wb.create_sheet(sheet_title(chap, used))
        w = SheetWriter(ws, chap, sheets[chap])
        for p in sheets[chap]:
            row0 = w.r
            if p.get("image"):
                w.image(p["image"])
            for b in p["blocks"]:
                if b["t"] == "text":
                    if b["text"].strip():
                        w.text(b)
                elif b["t"] == "chunks":
                    w.chunks(b["chunks"])
                else:
                    w.table(b)
            if not p["footer"]:
                p["footer"] = {"chunks": [(w.grid.right - 40, f"PDF page {p['no']}")]}
            w.page_end(p["footer"])
            page_map.append({"sheet": ws.title, "page": p["no"], "r0": row0, "r1": w.r - 1})
        w.finish(chap)
    wb.save(dst)
    if os.environ.get("MAPFILE"):
        import json
        with open(os.environ["MAPFILE"], "w") as fh:
            json.dump(page_map, fh)
    print(f"saved {dst}: {len(order)} sheets")


if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "grid")
