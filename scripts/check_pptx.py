#!/usr/bin/env python3
# VENDORED — source of truth: pptx-crash-check skill
#   (~/.claude/skills/pptx-crash-check/scripts/check_pptx.py).
# Verbatim copy; do not edit here. Fix upstream in the skill, then re-vendor.
# NB: this skill's local reference names differ (backend/app.py upstream →
#     lib/clone.py + scripts/build_deck.py chart-guard here).
"""Static checker for the corruption causes catalogued in known_issues.md.

Usage:
    python3 check_pptx.py <path-to-pptx>

Exits 0 if all checks pass. Non-zero with a summary if any fail.
"""
import sys
import posixpath
import re
import zipfile
from collections import Counter

try:
    from lxml import etree
except ImportError:
    sys.exit('lxml is required: pip install lxml')


REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
A16_NS = 'http://schemas.microsoft.com/office/drawing/2014/main'


def _rels_path(part):
    return f'{posixpath.dirname(part)}/_rels/{posixpath.basename(part)}.rels'


def check_zip_integrity(z, names):
    bad = z.testzip()
    if bad:
        return False, f'corrupt zip entry: {bad}'
    return True, f'{len(names)} parts, all CRCs ok'


def check_rel_targets_resolve(z, names):
    """Every internal rels Target must point to an existing part."""
    missing = []
    for rels_name in [n for n in names if n.endswith('.rels')]:
        tree = etree.fromstring(z.read(rels_name))
        base = posixpath.dirname(rels_name).replace('/_rels', '')
        if rels_name == '_rels/.rels':
            base = ''
        for rel in tree.findall(f'{{{REL_NS}}}Relationship'):
            if rel.get('TargetMode') == 'External':
                continue
            target = rel.get('Target', '')
            resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
            if resolved not in names:
                missing.append((rels_name, target, resolved))
    if missing:
        return False, f'{len(missing)} broken targets, e.g. {missing[0]}'
    return True, 'all rels targets resolve'


def check_no_orphan_rids(z, names):
    """Known issue #1: every r:id attribute in any XML part must resolve to
    a Relationship in the sibling rels file."""
    rid_attr = f'{{{R_NS}}}id'
    orphans = []
    for part in [n for n in names if n.endswith('.xml') and n != '[Content_Types].xml']:
        try:
            tree = etree.fromstring(z.read(part))
        except etree.XMLSyntaxError:
            continue
        rp = _rels_path(part)
        valid = set()
        if rp in names:
            rels_xml = z.read(rp).decode('utf-8', errors='replace')
            valid = set(re.findall(r'Id="(rId\d+)"', rels_xml))
        for el in tree.iter():
            v = el.get(rid_attr)
            if v and v.startswith('rId') and v not in valid:
                orphans.append((part, v, el.tag.split('}')[-1]))
    if orphans:
        return False, f'{len(orphans)} orphan r:id refs, e.g. {orphans[0]}'
    return True, 'no orphan r:id refs'


def check_no_empty_extlst(z, names):
    """Known issue #2: empty <p:extLst> (zero children) triggers Mac PP Repair."""
    empties = []
    for part in [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')]:
        tree = etree.fromstring(z.read(part))
        for el in tree.iter(f'{{{P_NS}}}extLst'):
            if len(el) == 0:
                empties.append(part)
                break
    if empties:
        return False, f'{len(empties)} slides with empty <p:extLst>, e.g. {empties[0]}'
    return True, 'no empty <p:extLst>'


def check_unique_p14_creation_ids(z, names):
    """Known issue #3: slide-level <p14:creationId val=N> must be globally unique."""
    vals = []
    for part in [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')]:
        tree = etree.fromstring(z.read(part))
        for el in tree.iter(f'{{{P14_NS}}}creationId'):
            v = el.get('val')
            if v:
                vals.append((part, v))
    counts = Counter(v for _, v in vals)
    dups = {v: c for v, c in counts.items() if c > 1}
    if dups:
        return False, f'{len(dups)} duplicated p14:creationId vals, e.g. {next(iter(dups.items()))}'
    return True, f'{len(vals)} slide creationIds, all unique'


def check_unique_a16_creation_ids(z, names):
    """Known issue #4: shape-level <a16:creationId id="{GUID}"> must be globally unique
    ACROSS slides. Within a slide the same GUID can legitimately appear multiple times
    (mc:AlternateContent siblings)."""
    by_guid = {}  # guid -> set of slide paths
    for part in [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')]:
        tree = etree.fromstring(z.read(part))
        slide_guids = set()
        for el in tree.iter(f'{{{A16_NS}}}creationId'):
            g = el.get('id')
            if g:
                slide_guids.add(g)
        for g in slide_guids:
            by_guid.setdefault(g, set()).add(part)
    cross = {g: ss for g, ss in by_guid.items() if len(ss) > 1}
    if cross:
        g, ss = next(iter(cross.items()))
        return False, f'{len(cross)} a16 GUIDs appear in multiple slides, e.g. {g} in {sorted(ss)}'
    return True, f'{len(by_guid)} unique a16 GUIDs, no cross-slide dups'


def check_unique_p14_mod_ids(z, names):
    """Known issue #5: shape-level <p14:modId val="N"> must not collide across slides.
    Same source slide selected twice → identical modIds on its shapes → Mac PP Repair."""
    by_val = {}
    for part in [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')]:
        tree = etree.fromstring(z.read(part))
        slide_vals = set()
        for el in tree.iter(f'{{{P14_NS}}}modId'):
            v = el.get('val')
            if v:
                slide_vals.add(v)
        for v in slide_vals:
            by_val.setdefault(v, set()).add(part)
    cross = {v: ss for v, ss in by_val.items() if len(ss) > 1}
    if cross:
        v, ss = next(iter(cross.items()))
        return False, f'{len(cross)} p14:modId vals appear in multiple slides, e.g. {v} in {sorted(ss)}'
    return True, f'{len(by_val)} unique p14:modId vals, no cross-slide dups'


CHECKS = [
    ('zip integrity', check_zip_integrity),
    ('rels targets resolve', check_rel_targets_resolve),
    ('no orphan r:id refs (issue #1)', check_no_orphan_rids),
    ('no empty <p:extLst> (issue #2)', check_no_empty_extlst),
    ('unique p14:creationId (issue #3)', check_unique_p14_creation_ids),
    ('unique a16:creationId across slides (issue #4)', check_unique_a16_creation_ids),
    ('unique p14:modId across slides (issue #5)', check_unique_p14_mod_ids),
]


def main(path):
    with zipfile.ZipFile(path, 'r') as z:
        names = set(z.namelist())
        results = []
        for name, fn in CHECKS:
            try:
                ok, detail = fn(z, names)
            except Exception as e:
                ok, detail = False, f'check raised {type(e).__name__}: {e}'
            results.append((name, ok, detail))

    width = max(len(n) for n, _, _ in results)
    failed = 0
    for name, ok, detail in results:
        tag = 'PASS' if ok else 'FAIL'
        print(f'  [{tag}] {name.ljust(width)}  {detail}')
        if not ok:
            failed += 1
    print()
    if failed:
        print(f'{failed}/{len(results)} checks failed.')
        print('If all known checks pass but PowerPoint still shows Repair, see')
        print('SKILL.md → "Discovery loop" to identify a new failure mode.')
        return 1
    print(f'All {len(results)} known checks passed.')
    print('If PowerPoint still shows Repair: this is a new failure mode.')
    print('Run discovery loop and append the cause to known_issues.md.')
    return 0


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(f'Usage: {sys.argv[0]} <path-to-pptx>')
    sys.exit(main(sys.argv[1]))
