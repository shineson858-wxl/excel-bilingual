#!/usr/bin/env python
"""Inventory the translatable text in an .xlsx workbook.

Prints per-sheet counts and writes a TSV of unique source strings with
occurrence counts, so the glossary can be authored against real text.

A cell counts as "translatable text" when either:
  * its value is a non-empty literal string, or
  * its value is a formula whose cached result (data_only view) is a non-empty
    string -- this is where a lot of unseen text hides in linked workbooks.

Usage:
  python scan_text.py --source book.xlsx [--out unique.tsv] [--sheets "A,B"]
"""

import argparse
import re
import sys
import unicodedata
from collections import Counter

from openpyxl import load_workbook


def nfc(s):
    return unicodedata.normalize("NFC", s)


def norm(s):
    return nfc(s).replace("\r\n", "\n").replace("\r", "\n").strip()


def collapse(s):
    return re.sub(r"\s+", " ", norm(s)).strip()


def text_of(literal_value, cached_value):
    """Return the text worth translating for one cell, or None."""
    v = literal_value
    if isinstance(v, str) and v.startswith("="):
        if isinstance(cached_value, str) and cached_value.strip():
            return cached_value, "formula->text"
        return None, None
    if isinstance(v, str) and v.strip():
        return v, "text"
    return None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--out", default="unique_strings.tsv")
    ap.add_argument("--sheets", default="", help="comma separated sheet names to limit to")
    args = ap.parse_args()

    keep = [s.strip() for s in args.sheets.split(",") if s.strip()]
    wf = load_workbook(args.source, data_only=False)
    wv = load_workbook(args.source, data_only=True)

    uniq = Counter()
    per_sheet = {}
    cell_total = 0
    chars_total = 0

    for ws in wf.worksheets:
        if keep and ws.title not in keep:
            continue
        vs = wv[ws.title]
        n_text = n_formula = 0
        for row in ws.iter_rows():
            for c in row:
                txt, kind = text_of(c.value, vs[c.coordinate].value)
                if txt is None:
                    continue
                n_text += 1
                if kind == "formula->text":
                    n_formula += 1
                cell_total += 1
                chars_total += len(txt)
                uniq[collapse(txt)] += 1
        per_sheet[ws.title] = (n_text, n_formula, ws.dimensions)

    print(f"sheets: {len(per_sheet)}  translatable cells: {cell_total}  unique: {len(uniq)}  chars: {chars_total}")
    for name, (n, nf, dims) in per_sheet.items():
        print(f"  {name[:28]:30s} dims={dims:12s} text_cells={n:5d} of which formula-cached={nf:5d}")

    with open(args.out, "w", encoding="utf-8") as f:
        f.write("count\ttext\n")
        for t, n in uniq.most_common():
            f.write(f"{n}\t{t.replace(chr(10), ' ⏎ ')}\n")
    print(f"wrote {args.out}")


if __name__ == "__main__":
    sys.exit(main())
