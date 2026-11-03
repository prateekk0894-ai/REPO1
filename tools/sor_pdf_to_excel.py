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


def table_block(tbl):
    rows = tbl.rows
    tops = [r.bbox[1] for r in rows] + [rows[-1].bbox[3]]
    data = tbl.extract()
    cells = []
    for ri, row in enumerate(rows):
        for ci, bb in enumerate(row.cells):
            if bb is None:
                continue
            # rows spanned: those whose top lies inside the cell
            r1 = ri
            while r1 + 1 < len(rows) and tops[r1 + 1] < bb[3] - 2:
                r1 += 1
            txt = data[ri][ci] if ci < len(data[ri]) else None
            txt = join_lines(txt.split("\n")) if txt else None
            cells.append((ri, r1, bb[0], bb[2], txt))
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


def road_block(hdr, body, bounds, left, right):
    """Un-ruled road SOR item table -> table block with one record per item."""
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
            elif cx >= bounds["rate"]:
                cells["rate"].append(w)
            elif cx >= bounds["unit"]:
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
        marker = LIST_START.match(line["name"] or line["desc"] or "") is not None and not col2
        cur = recs[-1] if recs else None
        new = (
            cur is None
            or bool(line["item"])
            or (col2 and bool(line["name"]) and bool(cur["name"]) and (bool(line["desc"]) or bool(line["rate"])) and
                bool(cur["rate"] or cur["unit"]))
            or (bool(line["rate"]) and bool(cur["rate"]))
            or (marker and bool(cur["rate"] or cur["desc"]) and not cur["desc"][-1:] == [""])
        )
        if new:
            cur = {k: [] for k in ROAD_COLS}
            recs.append(cur)
        for k in ROAD_COLS:
            if line[k]:
                cur[k].append(line[k])

    cells = [(0, 0, xs["item"], xs["name"], "Item No."), (0, 0, xs["name"], xs["unit"], "Description"),
             (0, 0, xs["unit"], xs["rate"], "Unit"), (0, 0, xs["rate"], xs["end"], "Rate")]
    for i, rec in enumerate(recs, start=1):
        txt = {k: join_lines(v) if v else None for k, v in rec.items()}
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
        chunks = split_gaps(ln["words"], 25)
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
    words = pg.extract_words(x_tolerance=1.5, y_tolerance=3, extra_attrs=["fontname", "size"])
    free = [w for w in words if not any(inside(w, t.bbox) for t in tables)]
    lines = group_lines(free)
    footer = [ln for ln in lines if is_footer(ln, pg.height, kind)]
    lines = [ln for ln in lines if ln not in footer]
    footer_text = " ".join(ln["text"] for ln in footer)
    all_x0 = [w["x0"] for w in words] or [50]
    all_x1 = [w["x1"] for w in words] or [pg.width - 50]
    left = min(min(all_x0), min((t.bbox[0] for t in tables), default=999))
    right = max(max(all_x1), max((t.bbox[2] for t in tables), default=0))

    blocks = [table_block(t) for t in tables]
    if kind == "road":
        hdr, bounds = road_header(lines)
        if bounds:
            body = [ln for ln in lines if ln["top"] > hdr["top"] + 2]
            # stray second header line ("No.")
            body = [ln for ln in body if not (ln["top"] - hdr["top"] < 14 and ln["text"].strip() in ("No.", "(Rs.)"))]
            above = [ln for ln in lines if ln["top"] < hdr["top"] - 2]
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
    return {"blocks": blocks, "footer": footer_blk, "footer_text": footer_text, "top_lines": top_lines,
            "left": left, "right": right, "width": pg.width}


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


class Grid:
    def __init__(self, pages):
        xs = []
        for p in pages:
            for b in p["blocks"]:
                if b["t"] == "table":
                    for c in b["cells"]:
                        xs += [c[2], c[3]]
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
def lines_needed(text, width_pt, size):
    if not text:
        return 1
    cpl = max((width_pt - 4) / (size * 0.47), 1)
    n = 0
    for part in str(text).split("\n"):
        n += max(1, math.ceil(len(part) / cpl))
    return n


class SheetWriter:
    def __init__(self, ws, title, pages):
        self.ws = ws
        self.grid = Grid(pages)
        self.pages = pages
        self.r = 1
        self.need = {}  # row -> required height (pt)
        self.vneed = []  # (r0, r1, height) for vertically merged cells
        g = self.grid
        for c in range(1, g.ncols + 1):
            pt = g.width_pt(c, c)
            ws.column_dimensions[get_column_letter(c)].width = max(round((pt / 0.75 - 5) / 7, 2), 0.8)

    def _merge(self, r0, r1, c0, c1):
        if r1 > r0 or c1 > c0:
            self.ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)

    def _height(self, r, h):
        self.need[r] = max(self.need.get(r, 15), h)

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
        self._height(self.r, lines_needed(b["text"], g.width_pt(c0, c1), size) * size * 1.3 + 2)
        self.r += 1

    def chunks(self, chunks, small=False, color=None):
        g = self.grid
        cols = [g.col_at(x) for x, *_ in chunks]
        # keep strictly increasing columns
        for i in range(1, len(cols)):
            if cols[i] <= cols[i - 1]:
                cols[i] = min(cols[i - 1] + 1, g.ncols)
        h = 15
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
            cell = self.ws.cell(self.r, c0, to_number(ch[1]) if looks_like_rate(ch[1]) else ch[1])
            if isinstance(cell.value, (int, float)):
                cell.number_format = RUPEE_FMT
            cell.font = font(8 if small else 10, bold, color)
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="right" if (small and i == len(chunks) - 1 and i) else "left")
            self._merge(self.r, self.r, c0, c1)
            h = max(h, lines_needed(ch[1], g.width_pt(c0, c1), 10) * 13 + 2)
        self._height(self.r, h)
        self.r += 1

    def table(self, b):
        g = self.grid
        base = self.r
        hdr = b["hdr"]
        occupied = set()
        for (r0, r1, x0, x1, txt) in b["cells"]:
            c0, c1 = g.col_at(x0), g.col_end(x1)
            if c1 < c0:
                c1 = c0
            R0, R1 = base + r0, base + r1
            if any((rr, cc) in occupied for rr in range(R0, R1 + 1) for cc in range(c0, c1 + 1)):
                # overlapping after snapping: write into first free cell only
                if (R0, c0) in occupied:
                    continue
                c1, R1 = c0, R0
            for rr in range(R0, R1 + 1):
                for cc in range(c0, c1 + 1):
                    occupied.add((rr, cc))
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
            h = lines_needed(txt, g.width_pt(c0, c1), 10) * 13 + 3
            if R1 == R0:
                self._height(R0, h)
            else:
                self.vneed.append((R0, R1, h))
        self.r = base + b["nrows"]

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

    def finish(self, title):
        ws = self.ws
        for r0, r1, h in self.vneed:
            have = sum(self.need.get(r, 15) for r in range(r0, r1 + 1))
            if h > have:
                self.need[r1] = self.need.get(r1, 15) + (h - have)
        for r, h in self.need.items():
            ws.row_dimensions[r].height = min(round(h, 1), 409)
        ws.sheet_view.showGridLines = False
        ws.print_title_rows = None
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.5
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
            p = parse_page(pdf.pages[i].dedupe_chars(), kind)
            chap = chapters.page(p["footer_text"], p["top_lines"])
            p["no"] = i + 1
            if chap not in sheets:
                sheets[chap] = []
                order.append(chap)
            sheets[chap].append(p)
            if (i + 1) % 50 == 0:
                print(f"{src}: parsed {i + 1}/{len(pdf.pages)} pages", flush=True)

    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for chap in order:
        ws = wb.create_sheet(sheet_title(chap, used))
        w = SheetWriter(ws, chap, sheets[chap])
        for p in sheets[chap]:
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
        w.finish(chap)
    wb.save(dst)
    print(f"saved {dst}: {len(order)} sheets")


if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "grid")
