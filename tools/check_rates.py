#!/usr/bin/env python3
"""Check that every rate in a converted SOR workbook belongs to exactly one item.

For every ₹ rate cell it finds the item description it belongs to: the nearest text cell
to its left whose (merged) range covers the rate's row.  It reports
  * descriptions that carry more than one rate in the same column (two rates on one item)
  * rates with no description on their row (a rate sitting alone)
Tables with several rate columns (e.g. "flush / modular", "1 km / 2 km ...") give one
rate per column, which is correct and not reported.

Usage: python3 tools/check_rates.py workbook.xlsx [pagemap.json]
"""
import json
import sys
from collections import defaultdict

from openpyxl import load_workbook


def main(path, map_path=None):
    wb = load_workbook(path)
    pages = defaultdict(list)
    if map_path:
        for m in json.load(open(map_path)):
            pages[m["sheet"]].append(m)
    multi, orphan, total = [], [], 0
    for ws in wb.worksheets:
        anchor = {}  # (row, col) -> top-left of its merged range
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    anchor[(r, c)] = (rng.min_row, rng.min_col)

        def page_of(r):
            for m in pages.get(ws.title, []):
                if m["r0"] <= r <= m["r1"]:
                    return m["page"]
            return "?"

        rates_by_desc = defaultdict(list)
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, (int, float)) or "₹" not in (cell.number_format or ""):
                    continue
                total += 1
                r, c = cell.row, cell.column
                owner = None
                # prefer a text cell on exactly this row (cells merged over several rows are shared
                # headings such as "Quantity of sample" and do not identify the item)
                for cc in range(c - 1, 0, -1):
                    a = anchor.get((r, cc), (r, cc))
                    rng_rows = [rg for rg in ws.merged_cells.ranges if (rg.min_row, rg.min_col) == a]
                    single = not rng_rows or rng_rows[0].min_row == rng_rows[0].max_row
                    v = ws.cell(*a).value
                    if single and isinstance(v, str) and len(v.strip()) >= 4 and any(ch.isalpha() for ch in v):
                        owner = a
                        break
                for cc in (range(c - 1, 0, -1) if owner is None else []):
                    a = anchor.get((r, cc), (r, cc))
                    v = ws.cell(*a).value
                    if isinstance(v, str) and len(v.strip()) >= 4 and any(ch.isalpha() for ch in v):
                        owner = a
                        break
                    owner_cell = ws.cell(*a)
                    if v is not None and "₹" not in (owner_cell.number_format or "") and str(v).strip():
                        owner = a  # size / item number in a lookup table (e.g. "100" mm)
                        break
                if owner is None:
                    orphan.append((ws.title, page_of(r), cell.coordinate, cell.value))
                else:
                    rates_by_desc[(owner, c)].append((cell.coordinate, cell.value))
        for ((orow, ocol), col), lst in rates_by_desc.items():
            if len(lst) > 1:
                text = str(ws.cell(orow, ocol).value).replace("\n", " ")[:60]
                multi.append((ws.title, page_of(orow), text, lst))
    print(f"{path}: rates={total}  items with >1 rate={len(multi)}  rates without item={len(orphan)}")
    for t, pg, text, lst in multi:
        print(f"  MULTI  [{t[:24]}] page {pg}: {text!r} -> {[v for _, v in lst]}")
    for t, pg, coord, v in orphan:
        print(f"  ALONE  [{t[:24]}] page {pg}: {coord} = {v}")
    return multi, orphan


if __name__ == "__main__":
    main(*sys.argv[1:3])
