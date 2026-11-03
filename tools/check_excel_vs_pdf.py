#!/usr/bin/env python3
"""Page-by-page check of a converted SOR workbook against its PDF.

For every PDF page it compares, with the Excel rows written for that page
(page map produced by sor_pdf_to_excel.py with MAPFILE=...):
  * rates  - every rate printed on the PDF page must appear in Excel
  * text   - share of the PDF page's letters/digits present in Excel

Usage: python3 tools/check_excel_vs_pdf.py input.pdf output.xlsx pagemap.json
"""
import json
import re
import sys
from collections import Counter

import pdfplumber
from openpyxl import load_workbook

RATE_RE = re.compile(r"^\d[\d,]*\.\d\d$")
KD = {"**yksd", "fuekZ.k", "ls", "yksd", "dY;k.k**", "dY;k.k"}


def is_rate(t):
    t = t.strip()
    if not RATE_RE.match(t):
        return False
    ip = t.split(".")[0].replace(",", "")
    return t.endswith(".00") or len(ip) >= 3 or "," in t


def norm(tok):
    t = tok.strip(".,;:()[]{}\"'“”‘’*")
    if re.fullmatch(r"\d[\d,]*(\.\d+)?", t):
        try:
            return f"{float(t.replace(',', '')):.2f}"
        except ValueError:
            pass
    return t.lower()


def main(pdf_path, xlsx_path, map_path):
    pmap = json.load(open(map_path))
    wb = load_workbook(xlsx_path)
    report = []
    with pdfplumber.open(pdf_path) as pdf:
        for m in pmap:
            ws = wb[m["sheet"]]
            xl_tokens, xl_rates = Counter(), Counter()
            for row in ws.iter_rows(min_row=m["r0"], max_row=m["r1"], values_only=True):
                for v in row:
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        xl_rates[f"{float(v):.2f}"] += 1
                        xl_tokens[f"{float(v):.2f}"] += 1
                    elif isinstance(v, str):
                        for t in v.split():
                            xl_tokens[norm(t)] += 1
                            if is_rate(t):
                                xl_rates[norm(t)] += 1
            pg = pdf.pages[m["page"] - 1].dedupe_chars()
            words = [w["text"] for w in pg.extract_words(x_tolerance=1.5, y_tolerance=3)]
            pdf_tokens = Counter(norm(t) for t in words if t not in KD)
            pdf_rates = Counter(norm(t) for t in words if is_rate(t))
            missing_rates = pdf_rates - xl_rates
            # character-level coverage (letters/digits only) is immune to spacing differences
            pdf_chars = Counter(ch for t in words if t not in KD for ch in t.lower() if ch.isalnum())
            xl_chars = Counter()
            for row in ws.iter_rows(min_row=m["r0"], max_row=m["r1"], values_only=True):
                for v in row:
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        v = f"{float(v):.2f}"
                    if isinstance(v, str):
                        xl_chars.update(ch for ch in v.lower() if ch.isalnum())
            # page footer printed by Excel (footer sections with &P = page number)
            idx = [x["page"] for x in pmap if x["sheet"] == m["sheet"]].index(m["page"])
            first = ws.page_setup.firstPageNumber or 1
            for part in (ws.oddFooter.left, ws.oddFooter.center, ws.oddFooter.right):
                if part.text:
                    t = part.text.replace("&P", str(first + idx)).replace("&&", "&")
                    xl_chars.update(ch for ch in t.lower() if ch.isalnum())
                    for tok in t.split():
                        xl_tokens[norm(tok)] += 1
            found = sum((pdf_chars & xl_chars).values())
            total = sum(pdf_chars.values())
            cov = found / total if total else 1.0
            report.append({"page": m["page"], "sheet": m["sheet"], "words": total, "coverage": round(cov, 3),
                           "rates": sum(pdf_rates.values()), "missing_rates": dict(missing_rates)})
    bad = [r for r in report if r["missing_rates"] or (r["words"] >= 20 and r["coverage"] < 0.98)]
    tot_r = sum(r["rates"] for r in report)
    miss_r = sum(sum(r["missing_rates"].values()) for r in report)
    print(f"{pdf_path}: pages={len(report)} rates_in_pdf={tot_r} rates_missing={miss_r} "
          f"pages_flagged={len(bad)}")
    for r in bad:
        print(f"  page {r['page']:>4} [{r['sheet'][:28]}] coverage={r['coverage']:.0%} words={r['words']} "
              f"missing_rates={r['missing_rates'] or '-'}")
    return report


if __name__ == "__main__":
    main(*sys.argv[1:4])
