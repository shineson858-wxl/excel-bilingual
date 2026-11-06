#!/usr/bin/env python
"""Verify a same-cell bilingual .xlsx against its source file.

Checks, per sheet:
  * every changed cell is exactly "<original text>\\n<translation>";
  * no numeric cell changed;
  * formulas outside bilingual cells are untouched;
  * merged ranges, autoFilter range and conditional-formatting rule count are unchanged;
  * external-link parts still present in the package.

Exit code 0 = pass, 1 = problems found.

Usage:
  python verify_bilingual.py --source original.xlsx --out translated.xlsx [--sep "\n"] [--allow-relink]
"""

import argparse
import re
import sys
import zipfile

from openpyxl import load_workbook

# Excel renumbers the [n] ordinal of external-link references when it saves a
# workbook, so a COM-written file can look like "=+[8]Sheet!A1 -> =+[3]Sheet!A1".
# --allow-relink compares formula text with those ordinals collapsed.
LINK_ORDINAL = re.compile(r"\[\d+\]")


def norm_links(v):
    return LINK_ORDINAL.sub("[n]", v) if isinstance(v, str) else v


def flat(v):
    """Comparable form: drop the _x000D_ literal Excel substitutes for CR, drop CR, trim ends."""
    return v.replace("_x000D_", "\r").replace("\r", "").strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--sep", default="\n")
    ap.add_argument("--allow-com-artifacts", action="store_true",
                    help="tolerate two known Excel-COM side effects: renumbered external-link ordinals, "
                         "and leading/trailing spaces around the original text being normalized")
    args = ap.parse_args()

    wf = load_workbook(args.source, data_only=False)
    wv = load_workbook(args.source, data_only=True)
    wo = load_workbook(args.out, data_only=False)

    bilingual, bad, relinked, ws_edge = [], [], [], []
    for ws in wf.worksheets:
        vs = wv[ws.title]
        ows = wo[ws.title]
        for row in ws.iter_rows():
            for c in row:
                ov, nv = c.value, ows[c.coordinate].value
                if ov == nv:
                    continue
                if args.allow_com_artifacts and norm_links(ov) == norm_links(nv):
                    relinked.append((ws.title, c.coordinate))
                    continue
                raw = None
                if isinstance(ov, str) and ov.startswith("="):
                    cached = vs[c.coordinate].value
                    raw = cached if isinstance(cached, str) else None
                elif isinstance(ov, str):
                    raw = ov
                if raw is None:
                    bad.append((ws.title, c.coordinate, "non-text change", ov, nv))
                    continue
                head = raw.rstrip()
                if isinstance(nv, str) and nv.startswith(head + args.sep) and len(nv) > len(head) + len(args.sep):
                    bilingual.append((ws.title, c.coordinate))
                    continue
                if args.allow_com_artifacts and isinstance(nv, str) and nv.startswith(head.replace("\r", "_x000D_") + args.sep):
                    ws_edge.append((ws.title, c.coordinate))
                    continue
                if args.allow_com_artifacts and isinstance(nv, str):
                    parts = nv.split(args.sep, 1)
                    if len(parts) == 2 and flat(parts[0]) and flat(parts[0]) == flat(raw) and parts[1].strip():
                        ws_edge.append((ws.title, c.coordinate))
                        continue
                bad.append((ws.title, c.coordinate, "not <source+sep+translation>", raw, nv))

    numbers_changed = []
    for ws in wf.worksheets:
        ows = wo[ws.title]
        for row in ws.iter_rows():
            for c in row:
                ov, nv = c.value, ows[c.coordinate].value
                if ov != nv and not isinstance(ov, str):
                    numbers_changed.append((ws.title, c.coordinate, ov, nv))

    print(f"bilingual cells: {len(bilingual)}")
    if args.allow_com_artifacts:
        print(f"  tolerated: external-link ordinals renumbered: {len(relinked)}")
        print(f"  tolerated: edge spaces around original normalized: {len(ws_edge)}")
    print(f"bad cells: {len(bad)}")
    for x in bad[:15]:
        print("   ", x)
    print(f"numeric cells changed: {len(numbers_changed)}")
    for x in numbers_changed[:10]:
        print("   ", x)

    struct_ok = True
    for ws in wf.worksheets:
        ows = wo[ws.title]
        m1 = sorted(str(r) for r in ws.merged_cells.ranges)
        m2 = sorted(str(r) for r in ows.merged_cells.ranges)
        cf1 = len(ws.conditional_formatting._cf_rules)
        cf2 = len(ows.conditional_formatting._cf_rules)
        ok = m1 == m2 and str(ws.auto_filter.ref) == str(ows.auto_filter.ref) and cf1 == cf2 and ws.dimensions == ows.dimensions
        struct_ok = struct_ok and ok
        print(f"  {ws.title[:26]:28s} merged {len(m1):3d}/{len(m2):3d}  af {str(ws.auto_filter.ref):12s}  cf {cf1}/{cf2}  dims {ws.dimensions}/{ows.dimensions}  {'OK' if ok else 'CHANGED'}")

    z1 = zipfile.ZipFile(args.source).namelist()
    z2 = zipfile.ZipFile(args.out).namelist()
    e1 = len([n for n in z1 if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n)])
    e2 = len([n for n in z2 if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n)])
    print(f"externalLink parts: {e1} -> {e2}")

    fail = bad or numbers_changed or not struct_ok or e2 < e1
    print("RESULT:", "FAIL" if fail else "PASS")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
