#!/usr/bin/env python
"""Write "source language + line separator + translation" into the SAME cells of an .xlsx.

Never adds rows or columns. Keeps every cell that is not translated byte-identical
(numbers, real formulas, styles, merged ranges, autoFilter, conditional formatting).
Formula cells whose cached result is text are converted to static bilingual text
(see references/terminology-rules.md for why).

Usage:
  python apply_bilingual.py --source in.xlsx --out out.xlsx \
      --glossary assets/glossary/vi-zh-mep.json [--glossary more/*.json] \
      [--rules assets/rules/vi-spec-rules.json] \
      [--missing missing.txt] [--no-height] [--sheets "A,B"]

Glossary format: {"<source text>": "<translation>", ...}. Keys are matched
case-insensitively, whitespace-collapsed and NFC-normalized, so uppercase source
rows and Vietnamese combining-mark variants still hit.

Rules format: [{"pattern": "<regex>", "replacement": "<re.sub replacement>"}]
applied to the whitespace-collapsed source text when the glossary misses.
"""

import argparse
import glob
import json
import math
import os
import re
import unicodedata
from copy import copy

from openpyxl import load_workbook

SEP = "\n"


def nfc(s):
    return unicodedata.normalize("NFC", s)


def norm(s):
    return nfc(s).replace("\r\n", "\n").replace("\r", "\n").strip()


def collapse(s):
    return re.sub(r"\s+", " ", norm(s)).strip()


def key_forms(s):
    n = norm(s)
    return {n.lower(), collapse(s).lower()}


def load_glossary(paths):
    idx = {}
    for path in paths:
        files = [path]
        if os.path.isdir(path):
            files = sorted(glob.glob(os.path.join(path, "*.json")))
        for fp in files:
            with open(fp, encoding="utf-8") as f:
                data = json.load(f)
            for k, v in data.items():
                if not isinstance(v, str) or not v.strip():
                    continue
                for kf in key_forms(k):
                    if kf in idx and idx[kf] != v:
                        print(f"note: override {kf!r}: {idx[kf]!r} -> {v!r}")
                    idx[kf] = v
    return idx


def load_rules(path):
    if not path:
        return []
    with open(path, encoding="utf-8") as f:
        rules = json.load(f)
    return [(re.compile(r["pattern"], re.IGNORECASE), r["replacement"]) for r in rules]


def translate(raw, idx, rules):
    for kf in (norm(raw).lower(), collapse(raw).lower()):
        if kf in idx:
            return idx[kf]
    flat = collapse(raw)
    for pat, rep in rules:
        new, n = pat.subn(rep, flat)
        if n:
            return new
    return None


def col_width(ws, col_letter, fallback=8.43):
    w = ws.column_dimensions[col_letter].width
    return w if w else fallback


def span_of(ws, cell):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= cell.row <= rng.max_row and rng.min_col <= cell.column <= rng.max_col:
            return rng.max_col - rng.min_col + 1
    return 1


def lines_for(text, width, cjk_units=2.0):
    lines = 0
    for seg in str(text).split("\n"):
        units = sum(cjk_units if ord(ch) > 0x2E7F else 1.0 for ch in seg)
        lines += max(1, math.ceil(units / max(width, 3.0)))
    return lines


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--glossary", action="append", default=[])
    ap.add_argument("--rules", default="")
    ap.add_argument("--missing", default="")
    ap.add_argument("--no-height", action="store_true")
    ap.add_argument("--sheets", default="")
    args = ap.parse_args()

    idx = load_glossary(args.glossary)
    rules = load_rules(args.rules)
    print(f"glossary keys: {len(idx)}  rules: {len(rules)}")

    keep = [s.strip() for s in args.sheets.split(",") if s.strip()]
    wf = load_workbook(args.source, data_only=False)
    wv = load_workbook(args.source, data_only=True)

    changed = 0
    formula_to_text = 0
    missing = {}
    touched_rows = {}

    for ws in wf.worksheets:
        if keep and ws.title not in keep:
            continue
        vs = wv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                raw = None
                is_formula = False
                if isinstance(v, str) and v.startswith("="):
                    cached = vs[c.coordinate].value
                    if isinstance(cached, str) and cached.strip():
                        raw, is_formula = cached, True
                elif isinstance(v, str) and v.strip():
                    raw = v
                if raw is None:
                    continue
                zh = translate(raw, idx, rules)
                if zh is None:
                    missing[collapse(raw)] = missing.get(collapse(raw), 0) + 1
                    continue
                if zh.strip() == "" or norm(zh) == norm(raw):
                    continue
                c.value = raw.rstrip() + SEP + zh
                changed += 1
                if is_formula:
                    formula_to_text += 1
                al = copy(c.alignment)
                al.wrap_text = True
                c.alignment = al
                touched_rows.setdefault((ws.title, c.row), []).append(c.coordinate)

    if not args.no_height:
        for (title, r), coords in touched_rows.items():
            ws = wf[title]
            need, lh = 1, 14.0
            for coord in coords:
                cell = ws[coord]
                letter = re.match(r"[A-Z]+", coord).group(0)
                width = col_width(ws, letter) * span_of(ws, cell)
                need = max(need, lines_for(cell.value or "", width))
                fs = (cell.font.size or 10) if cell.font else 10
                lh = max(12.0, float(fs) * 1.45)
            old = ws.row_dimensions[r].height
            ws.row_dimensions[r].height = max(min(409.0, need * lh + 4.0), old or 0)

    wf.save(args.out)
    print(f"cells updated: {changed}  |  formula cells turned into static text: {formula_to_text}")
    print(f"untranslated unique: {len(missing)} cells: {sum(missing.values())}")
    if args.missing:
        with open(args.missing, "w", encoding="utf-8") as f:
            for t, n in sorted(missing.items()):
                f.write(f"{n}\t{t}\n")
        print(f"wrote {args.missing}")
    print(f"saved: {args.out}")


if __name__ == "__main__":
    main()
