#!/usr/bin/env python
"""Mine a glossary out of an already translated bilingual workbook.

Every cell in a same-cell bilingual sheet is "<source>\\n<translation>", where
either side may itself span several lines. This script finds the boundary (the
first line that is Chinese and carries no source-language diacritics), then
writes "<source>" -> "<translation>" pairs so the next similar file can reuse
the exact same wording instead of being re-translated.

Usage:
  python mine_glossary.py --bilingual out.xlsx --out glossary.json [--merge existing.json]
"""

import argparse
import json
import re
import sys

from openpyxl import load_workbook

CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
# Vietnamese / Latin diacritics that a Chinese line never carries
LATIN_DIA = re.compile(r"[ăâđêôơưĂÂĐÊÔƠƯáàảãạắằẳẵặấầẩẫậéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợùủũụứừửữựýỳỷỹỵ]")


def is_chinese_line(line):
    return bool(CJK.search(line)) and not LATIN_DIA.search(line)


def norm_key(s):
    """Same three-step normalization the apply script matches with."""
    import re as _re
    import unicodedata

    s = unicodedata.normalize("NFC", s).replace("\r\n", "\n").replace("\r", "\n")
    return _re.sub(r"\s+", " ", s).strip().lower()


def split_pair(value):
    """Return (source, translation) for a bilingual cell, or None."""
    if not isinstance(value, str) or "\n" not in value:
        return None
    lines = value.split("\n")
    for i in range(1, len(lines)):
        if is_chinese_line(lines[i]):
            src = "\n".join(lines[:i]).strip()
            dst = "\n".join(lines[i:]).strip()
            if src and dst and not is_chinese_line(lines[i - 1]):
                return src, dst
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bilingual", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--merge", default="", help="existing glossary json to merge into")
    ap.add_argument("--rules", default="", help="rules json whose patterns are already handled by rules; skip those pairs")
    args = ap.parse_args()

    skip_patterns = []
    if args.rules:
        with open(args.rules, encoding="utf-8") as f:
            skip_patterns = [re.compile(r["pattern"], re.IGNORECASE) for r in json.load(f)]

    pairs, conflicts = {}, {}
    wb = load_workbook(args.bilingual, data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                pair = split_pair(c.value)
                if not pair:
                    continue
                src, dst = pair
                if src in pairs and pairs[src] != dst:
                    conflicts.setdefault(src, [pairs[src]]).append(dst)
                    continue
                pairs[src] = dst

    pairs = {norm_key(k): v.strip() for k, v in pairs.items()}
    conflicts = {norm_key(k): v for k, v in conflicts.items()}
    if skip_patterns:
        dropped = [k for k in pairs if any(p.match(k) for p in skip_patterns)]
        for k in dropped:
            pairs.pop(k, None)
        print(f"skipped {len(dropped)} pairs already covered by rules")

    if args.merge:
        with open(args.merge, encoding="utf-8") as f:
            base = json.load(f)
        merged = dict(base)
        existing_norm = {norm_key(k): k for k in base}
        added = 0
        for k, v in pairs.items():
            nk = norm_key(k)
            if nk in existing_norm:
                if merged[existing_norm[nk]] != v:
                    conflicts.setdefault(nk, [merged[existing_norm[nk]]]).append(v)
                continue
            merged[nk] = v
            existing_norm[nk] = nk
            added += 1
        pairs = merged
        print(f"merged: {len(base)} existing + {added} new")

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(pairs, f, ensure_ascii=False, indent=1, sort_keys=True)
    print(f"pairs mined: {len(pairs)}  conflicts: {len(conflicts)} -> {args.out}")
    for k, v in list(conflicts.items())[:10]:
        print(f"  CONFLICT {k!r}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
