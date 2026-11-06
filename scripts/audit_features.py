#!/usr/bin/env python
"""Preflight audit: will an openpyxl round-trip damage this workbook?

openpyxl preserves what it models (values, styles, merged ranges, autoFilter,
conditional formatting, comments, data validation, defined names, images,
charts it can parse) but silently drops several Excel-only parts:
embedded shapes / text boxes, slicers, form controls, pivot tables, VBA,
threaded comments, rich-value cells, x14 extended validations.

Run this BEFORE applying bilingual text. If the verdict is RISKY, write through
the Excel COM path (scripts/apply_bilingual_com.ps1) instead of openpyxl.

Usage:
  python audit_features.py --source book.xlsx
"""

import argparse
import re
import sys
import zipfile

from openpyxl import load_workbook

RISK_PARTS = {
    "xl/drawings/": "drawings (shapes / text boxes / images)",
    "xl/charts/": "charts",
    "xl/pivotTables/": "pivot tables",
    "xl/pivotCache/": "pivot caches",
    "xl/slicers": "slicers",
    "xl/ctrlProps/": "form controls",
    "xl/activeX/": "ActiveX controls",
    "xl/vbaProject.bin": "VBA macros",
    "xl/threadedComments/": "threaded comments",
    "xl/richData/": "rich value cells",
    "xl/media/": "embedded images",
}


def zip_profile(path):
    names = zipfile.ZipFile(path).namelist()
    hits = {}
    for prefix, label in RISK_PARTS.items():
        n = [x for x in names if x.startswith(prefix)]
        if n:
            hits[label] = len(n)
    # text boxes live inside drawing parts as <xdr:sp>; plain pictures are <xdr:pic>
    shapes = 0
    for n in names:
        if re.match(r"xl/drawings/drawing\d+\.xml$", n):
            body = zipfile.ZipFile(path).read(n).decode("utf-8", "ignore")
            shapes += len(re.findall(r"<xdr:sp[ >]", body))
    if shapes:
        hits[f"text shapes inside drawings ({shapes})"] = shapes
    ext_links = len([x for x in names if re.match(r"xl/externalLinks/externalLink\d+\.xml$", x)])
    x14 = 0
    for n in names:
        if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            body = zipfile.ZipFile(path).read(n).decode("utf-8", "ignore")
            x14 += body.count("x14:dataValidation")
    return hits, ext_links, x14, len(names)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    args = ap.parse_args()

    path = args.source
    hits, ext_links, _x14, nparts = zip_profile(path)

    wb = load_workbook(path, data_only=False)
    dv = sum(len(ws.data_validations.dataValidation) for ws in wb.worksheets)
    cf = sum(len(ws.conditional_formatting._cf_rules) for ws in wb.worksheets)
    merged = sum(len(ws.merged_cells.ranges) for ws in wb.worksheets)
    comments = sum(
        1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.comment
    )
    names = len(list(wb.defined_names)) if hasattr(wb.defined_names, "__len__") else 0
    protected = [ws.title for ws in wb.worksheets if ws.protection.sheet]
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]

    print(f"file: {path}")
    print(f"  zip parts: {nparts}  sheets: {len(wb.worksheets)} ({len(hidden)} hidden)")
    print(f"  merged ranges: {merged}  conditional-format rules: {cf}  data validations: {dv}")
    print(f"  comments: {comments}  defined names: {names}  external links: {ext_links}")
    if protected:
        print(f"  protected sheets: {', '.join(protected)}")

    risky = [k for k in hits if "embedded images" not in k]
    workable = [k for k in hits if "embedded images" in k]
    if comments:
        workable.append(f"comments ({comments}) get re-serialised")
    if _x14:
        risky.append(f"x14 extended data validation ({_x14})")

    if hits:
        print("  parts present:")
        for k, v in hits.items():
            print(f"    - {k}: {v}")
    else:
        print("  parts present: none of the fragile ones")

    if risky:
        print("RESULT: RISKY  -> write with the COM engine (scripts/apply_bilingual_com.ps1)")
        print("        openpyxl would drop or alter: " + ", ".join(risky))
        return 1
    if workable:
        print("RESULT: OPENPYXL_OK_WITH_CAUTION  -> these objects may be re-serialised: " + ", ".join(workable))
        return 0
    print("RESULT: OPENPYXL_OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
