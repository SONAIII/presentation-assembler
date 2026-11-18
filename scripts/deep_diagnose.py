#!/usr/bin/env python3
# VENDORED — source of truth: pptx-crash-check skill
#   (~/.claude/skills/pptx-crash-check/scripts/deep_diagnose.py).
# Verbatim copy; do not edit here. Fix upstream in the skill, then re-vendor.
# NB: this skill's local reference names differ (backend/app.py upstream →
#     lib/clone.py + scripts/build_deck.py chart-guard here).
"""Discovery-loop helper. When `check_pptx.py` reports clean but PowerPoint
still flags the file, run this against a known-good vs. known-bad pair to
surface candidate failure modes we haven't catalogued yet.

Usage:
    python3 deep_diagnose.py <bad.pptx> <good.pptx>

What it does:
- Diffs the part-name sets.
- For shared parts, diffs SHA-256 hashes.
- Across all XML parts in BAD, looks for cross-part / cross-slide ID-like
  attribute duplicates (anything that ends in `Id` or `Guid`, plus GUID-shaped
  attr values). Anything duplicated in BAD but unique in GOOD is a prime
  candidate.

The output is not a diagnosis — it's a focused candidate list to investigate.
"""
import sys
import hashlib
import posixpath
import re
import zipfile
from collections import Counter, defaultdict

try:
    from lxml import etree
except ImportError:
    sys.exit('lxml is required: pip install lxml')


GUID_RE = re.compile(r'^\{[0-9A-Fa-f-]{36}\}$')


def part_hashes(z):
    return {n: hashlib.sha256(z.read(n)).hexdigest()[:16] for n in z.namelist()}


def collect_id_attrs(z):
    """For every XML part, collect every attribute whose local name ends with
    'Id' or 'id' or whose value looks like a GUID. Group by (attr_tag, value)
    and record which parts each value appears in."""
    by_key = defaultdict(set)
    for name in z.namelist():
        if not name.endswith('.xml') or name == '[Content_Types].xml':
            continue
        try:
            tree = etree.fromstring(z.read(name))
        except etree.XMLSyntaxError:
            continue
        for el in tree.iter():
            for ak, av in el.attrib.items():
                local = ak.split('}')[-1] if '}' in ak else ak
                ns = ak.split('}')[0][1:] if ak.startswith('{') else ''
                # Skip r:id — covered by check #1 already
                if ak == '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':
                    continue
                if local.endswith('Id') or local == 'id' or GUID_RE.match(av or ''):
                    key = (el.tag.split('}')[-1], local, av)
                    by_key[key].add(name)
    return by_key


def main(bad_path, good_path):
    with zipfile.ZipFile(bad_path, 'r') as zb, zipfile.ZipFile(good_path, 'r') as zg:
        bad_h = part_hashes(zb)
        good_h = part_hashes(zg)

        bad_only = set(bad_h) - set(good_h)
        good_only = set(good_h) - set(bad_h)
        diff_content = [n for n in set(bad_h) & set(good_h) if bad_h[n] != good_h[n]]

        print('## Part-name diff')
        print(f'  bad only ({len(bad_only)}): {sorted(bad_only)[:10]}')
        print(f'  good only ({len(good_only)}): {sorted(good_only)[:10]}')
        print(f'  shared but different content ({len(diff_content)})')

        print()
        print('## Cross-part ID-attribute duplicates (BAD)')
        bad_ids = collect_id_attrs(zb)
        good_ids = collect_id_attrs(zg)
        dups_in_bad = {k: parts for k, parts in bad_ids.items() if len(parts) > 1}
        # Filter to interesting: dup in BAD but unique (or absent) in GOOD
        suspicious = []
        for k, parts in dups_in_bad.items():
            good_parts = good_ids.get(k, set())
            if len(good_parts) <= 1:
                suspicious.append((k, parts, good_parts))
        # Sort by tag/local-name for readability
        suspicious.sort(key=lambda x: (x[0][0], x[0][1]))
        print(f'  {len(suspicious)} ID-shaped attribute values are duplicated in BAD but unique in GOOD:')
        for (tag, local, val), parts_bad, parts_good in suspicious[:40]:
            short_val = (val[:30] + '…') if val and len(val) > 30 else val
            print(f'    <{tag} {local}="{short_val}"> in {len(parts_bad)} parts (good: {len(parts_good)})')
        if len(suspicious) > 40:
            print(f'    ... and {len(suspicious) - 40} more')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(f'Usage: {sys.argv[0]} <bad.pptx> <good.pptx>')
    main(sys.argv[1], sys.argv[2])
