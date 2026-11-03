#!/usr/bin/env python3
"""Convert MP PWD Schedule-of-Rates PDFs into editable Excel workbooks.

One workbook per PDF, one sheet per chapter.  Ruled tables are copied
cell-for-cell; text outside tables (headings, notes) is copied line by line
in page order.  Rates are stored as numbers so they can be used in formulas.

Usage:  python3 tools/sor_pdf_to_excel.py input.pdf output.xlsx [mode]
mode: "road" (Road & Bridge SOR, un-ruled item tables) or "grid" (default).
"""
import math
import re
import sys
from collections import Counter

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SLOGAN_KD = "yksd fuekZ.k ls yksd dY;k.k"
SLOGAN_UNI = "लोक निर्माण से लोक कल्याण"
NUM_RE = re.compile(r"^-?\d[\d,]*(\.\d+)?$")
RATE_RE = re.compile(r"^\d[\d,]*\.\d\d$")
BAD_SHEET_CHARS =re.compile(r"[\[\]\:\*\?\/\\]")

FONT = Font(name="Arial", size=10)
FONT_B = Font(name="Arial", size=10, bold=True)
FONT_H = Font(name="Arial", size=11, bold=True)
FONT_PG = Font(name="Arial", size=8, color="808080")
FILL_H = PatternFill("solid", start_color="D9E1F2")
THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def fix_text(s):
    if s is None:
        return None
    s = s.replace(SLOGAN_KD, SLOGAN_UNI).replace("", "•")
    return s


def join_cell(s):
    """Join the visual line-breaks of a PDF cell into flowing text."""
    if s is None:
        return None
    parts = [p.strip() for p in s.split("\n")]
    out = ""
    for p in parts:
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
    if s is None:
        return s
    t = s.strip()
    if NUM_RE.match(t):
        try:
            v = float(t.replace(",", ""))
            return int(v) if v.is_integer() and "." not in t else v
        except ValueError:
            return s
    return s


def group_lines(words, tol=3.0):
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    lines = []
    for w in words:
        if lines and abs(w["top"] - lines[-1]["top"]) <= tol:
            lines[-1]["words"].append(w)
        else:
            lines.append({"top": w["top"], "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
    return lines


def split_gaps(words, gap):
    """Split a line's words into chunks where the horizontal gap >= gap."""
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
    """Decides which chapter (sheet) each page belongs to."""

    def __init__(self, kind, pdf):
        self.kind = kind
        self.current = "Front Matter & General"
        self.titles = {}
        self.road_index = []
        if kind == "road":
            self._road_index(pdf)

    def _road_index(self, pdf):
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
        # building SOR: footer "CHAPTER – 4 : PLAIN CEMENT CONCRETE   34"
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


# ---------------------------------------------------------------- page parse
def is_footer(line, page_h, kind):
    if line["top"] < page_h * 0.88:
        return False
    t = line["text"]
    if kind == "road":
        return "SOR for Road" in t or bool(re.search(r"Page\s+\d+\s*$", t)) or SLOGAN_KD in t
    if kind == "em":
        return "ELECTRICAL SOR" in t or bool(re.search(r"Page-\s*\d+", t))
    return bool(re.match(r"\s*CHAPTER", t, re.I)) or bool(re.fullmatch(r"\s*\d+\s*", t))


ROAD_COLS = ["Item No.", "Item Name", "Description", "Unit", "Rate"]


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


def road_rows(lines, bounds):
    """Split un-ruled item lines into the 5 road columns."""
    # common starting x positions inside the description zone -> 2nd text column
    starts = Counter()
    for ln in lines:
        for ch in split_gaps(ln["words"], 14):
            x = ch[0]["x0"]
            if bounds["item"] + 25 < x < bounds["unit"] - 60:
                starts[round(x / 3)] += 1
    col2 = sorted(k * 3 for k, c in starts.items() if c >= 4)
    rows = []
    for ln in lines:
        cells = {k: [] for k in ROAD_COLS}
        desc_words = []
        for w in ln["words"]:
            cx = (w["x0"] + w["x1"]) / 2
            if w["x1"] <= bounds["item"] + 2 and cx < bounds["item"]:
                cells["Item No."].append(w)
            elif cx >= bounds["rate"]:
                cells["Rate"].append(w)
            elif cx >= bounds["unit"]:
                cells["Unit"].append(w)
            else:
                desc_words.append(w)
        if desc_words:
            chunks = split_gaps(desc_words, 14)
            first = chunks[0]
            if (
                col2
                and len(chunks) >= 2
                and first[0]["x0"] < bounds["item"] + 20
                and any(abs(chunks[1][0]["x0"] - c) <= 4 for c in col2)
            ):
                cells["Item Name"] = first
                cells["Description"] = [w for ch in chunks[1:] for w in ch]
            elif col2 and first[0]["x0"] < bounds["item"] + 20 and first[-1]["x1"] < min(col2) - 8:
                cells["Item Name"] = desc_words
            else:
                cells["Description"] = desc_words
        row = [fix_text(" ".join(w["text"] for w in cells[k])) or None for k in ROAD_COLS]
        row[4] = to_number(row[4]) if row[4] else None
        rows.append(row)
    return rows


def table_rows(tbl):
    data = tbl.extract()
    out, merges = [], []
    for r, row in enumerate(data):
        vals = [join_cell(c) if c is not None else None for c in row]
        # horizontal merges: None to the right of a real cell
        c = 0
        while c < len(row):
            if row[c] is not None:
                e = c
                while e + 1 < len(row) and row[e + 1] is None:
                    e += 1
                if e > c:
                    merges.append((r, c, e))
                c = e + 1
            else:
                c += 1
        out.append(vals)
    return out, merges


def numeric_cols(rows):
    """Columns at/after a 'Rate' header (or after 'Unit') hold numbers."""
    for row in rows[:4]:
        for i, v in enumerate(row):
            if v and re.search(r"\bRate\b", v, re.I):
                return set(range(i, len(row)))
        for i, v in enumerate(row):
            if v and re.fullmatch(r"Unit", v.strip(), re.I):
                return set(range(i + 1, len(row)))
    return set()


def parse_page(pg, kind, state):
    """Return list of blocks: ('text', str) / ('cells', [..]) / ('table', rows, merges, numcols, header_n)."""
    tables = [t for t in pg.find_tables() if len(t.rows) >= 1 and (t.bbox[3] - t.bbox[1]) > 15]
    words = pg.extract_words(x_tolerance=1.5, y_tolerance=3, keep_blank_chars=False)
    free = [w for w in words if not any(inside(w, t.bbox) for t in tables)]
    lines = group_lines(free)
    footer = [ln for ln in lines if is_footer(ln, pg.height, kind)]
    lines = [ln for ln in lines if ln not in footer]
    footer_text = " ".join(ln["text"] for ln in footer)

    items = []  # (top, block)
    for t in tables:
        rows, merges = table_rows(t)
        if numeric_cols(rows):
            state["numcols"] = (len(rows[0]), numeric_cols(rows))
        nc = numeric_cols(rows)
        if not nc and state.get("numcols") and state["numcols"][0] == len(rows[0]):
            nc = state["numcols"][1]
        items.append((t.bbox[1], ("table", rows, merges, nc)))

    if kind == "road":
        hdr, bounds = road_header(lines)
        if bounds:
            state["road_bounds"] = bounds
            body = [ln for ln in lines if ln["top"] > hdr["top"] + 2]
            above = [ln for ln in lines if ln["top"] < hdr["top"] - 2]
            for ln in above:
                items.append((ln["top"], ("text", fix_text(ln["text"]))))
            items.append((hdr["top"], ("header", ROAD_COLS)))
            for ln, row in zip(body, road_rows(body, bounds)):
                items.append((ln["top"], ("cells", row)))
            lines = []
    for ln in lines:
        chunks = split_gaps(ln["words"], 25)
        if len(chunks) == 1:
            items.append((ln["top"], ("text", fix_text(ln["text"]))))
        else:
            items.append((ln["top"], ("cells", [fix_text(" ".join(w["text"] for w in ch)) for ch in chunks])))
    items.sort(key=lambda x: x[0])
    top_lines = [ln["text"] for ln in sorted(lines, key=lambda l: l["top"])[:8]]
    if kind == "road" and not top_lines:
        top_lines = []
    return [b for _, b in items], footer_text, top_lines


# ---------------------------------------------------------------- writer
class SheetWriter:
    def __init__(self, ws, title):
        self.ws = ws
        self.r = 1
        self.widths = Counter()
        ws.sheet_view.zoomScale = 90
        ws.column_dimensions["A"].width = 6
        c = ws.cell(1, 1, title)
        c.font = Font(name="Arial", size=13, bold=True)
        ws.cell(2, 1, "Col A = PDF page no. All cells are editable; rates are numbers.").font = FONT_PG
        self.r = 4
        self.last_page = None

    def _pg(self, page):
        c = self.ws.cell(self.r, 1, page)
        c.font = FONT_PG
        c.alignment = Alignment(vertical="top")
        if page != self.last_page:
            c.font = Font(name="Arial", size=8, color="C00000", bold=True)
            self.last_page = page

    def text(self, page, s):
        self._pg(page)
        c = self.ws.cell(self.r, 2, s)
        bold = bool(re.match(r"\s*(CHAPTER|Note|NOTE|General Notes|GENERAL)", s)) or (s.isupper() and len(s) > 3)
        c.font = FONT_H if bold else FONT
        self.r += 1

    def row(self, page, vals, header=False, border=False, numcols=(), merges=()):
        self._pg(page)
        max_lines = 1
        for i, v in enumerate(vals):
            col = i + 2
            if numcols and i in numcols and isinstance(v, str):
                v = to_number(v)
            c = self.ws.cell(self.r, col, v if v != "" else None)
            c.font = FONT_B if header else FONT
            c.alignment = WRAP_TOP
            if isinstance(v, (int, float)):
                c.number_format = "0.00" if isinstance(v, float) else "0"
            if header:
                c.fill = FILL_H
                c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            if border:
                c.border = BORDER
            if isinstance(v, str):
                self.widths[(col, "len")] = max(self.widths[(col, "len")], min(len(v), 70))
                self.widths[(col, "n")] += 1
        for c0, c1 in merges:
            self.ws.merge_cells(start_row=self.r, start_column=c0 + 2, end_row=self.r, end_column=c1 + 2)
        self.r += 1
        return max_lines

    def finish(self):
        ws = self.ws
        # rates left as text in tables without a detectable "Rate" header
        for row in ws.iter_rows(min_row=4, min_col=5):
            for c in row:
                v = c.value
                if isinstance(v, str) and RATE_RE.match(v.strip()):
                    ip = v.strip().split(".")[0].replace(",", "")
                    if v.strip().endswith(".00") or len(ip) >= 3 or "," in v:
                        c.value = to_number(v)
                        c.number_format = "0.00"
        widths = {}
        for (col, k), v in self.widths.items():
            if k == "len":
                widths[col] = max(8, min(v + 2, 62))
        for col, w in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = w
        # row heights for wrapped cells
        for row in ws.iter_rows(min_row=4):
            lines = 1
            for c in row:
                if isinstance(c.value, str) and c.alignment and c.alignment.wrap_text:
                    w = widths.get(c.column, 10)
                    lines = max(lines, math.ceil(len(c.value) * 1.1 / max(w - 1, 1)))
            if lines > 1:
                ws.row_dimensions[row[0].row].height = min(13.5 * lines, 400)
        ws.freeze_panes = "B4"


def convert(src, dst, kind):
    wb = Workbook()
    wb.remove(wb.active)
    used, writers = set(), {}
    state = {}
    with pdfplumber.open(src) as pdf:
        chapters = Chapters(kind, pdf)
        pending = []  # pages before first chapter footer get the next chapter only if empty context
        import os
        rng = os.environ.get("PAGES")
        sel = range(*map(int, rng.split("-"))) if rng else range(len(pdf.pages))
        for i in sel:
            pg = pdf.pages[i].dedupe_chars()
            blocks, footer_text, top_lines = parse_page(pg, kind, state)
            chap = chapters.page(footer_text, top_lines)
            if chap not in writers:
                ws = wb.create_sheet(sheet_title(chap, used))
                writers[chap] = SheetWriter(ws, chap)
            w = writers[chap]
            for b in blocks:
                if b[0] == "text":
                    if b[1] and b[1].strip():
                        w.text(i + 1, b[1])
                elif b[0] == "header":
                    w.row(i + 1, b[1], header=True, border=True)
                elif b[0] == "cells":
                    w.row(i + 1, b[1], border=(kind == "road"))
                else:
                    _, rows, merges, nc = b
                    hdr_n = 0
                    for k, rw in enumerate(rows[:3]):
                        joined = " ".join(x for x in rw if x)
                        if re.search(r"\b(Item|Description|Unit|Rate|S\.?\s*No)\b", joined, re.I) or re.fullmatch(r"[\d\s]+", joined or "x"):
                            hdr_n = k + 1
                        else:
                            break
                    for k, rw in enumerate(rows):
                        m = [(a, b2) for (r0, a, b2) in merges if r0 == k]
                        w.row(i + 1, rw, header=k < hdr_n, border=True, numcols=nc if k >= hdr_n else (), merges=m)
                    w.r += 1
            if (i + 1) % 50 == 0:
                print(f"{src}: {i + 1}/{len(pdf.pages)} pages", flush=True)
    for w in writers.values():
        w.finish()
    wb.save(dst)
    print(f"saved {dst}: {len(writers)} sheets")


if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "grid")
