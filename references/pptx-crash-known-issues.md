# Known PPTX corruption causes

Living registry of every failure mode we've hit. Update when we find a new one.

> Записи ниже — исторические, с путями проекта-родителя `viscomms-2`
> (`backend/app.py`, `useCart.js`): реестр не переписываем. В этом скилле
> те же ремапы (#3–#5) реализует `lib/clone.py`, а «блокировка дубля
> upstream» (#6) — chart-guard в `scripts/build_deck.py`.

## 1. Orphan `r:id` references in slide XML

**Symptom:** Mac PowerPoint "found a problem" dialog. Other parsers (python-pptx, QuickLook) tolerate it.

**Root cause:** Slide XML contains an inline element like
```xml
<p:extLst>
  <p:ext uri="{6950BFC3-D8DA-4A85-94F7-54DA5524770B}">
    <p188:commentRel xmlns:p188="..." r:id="rId5"/>
  </p:ext>
</p:extLst>
```
that points to a relationship (`rId5` → comments) we stripped from `slide.xml.rels`.

**Where seen:** Yango template — every slide has a `p188:commentRel` backlink to its modern-comments file. Removing the rels entry without removing the inline backlink leaves an orphan ref.

**Detection:**
```python
# For every part with a .rels sibling, every r:id attribute must resolve.
rid_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
for el in tree.iter():
    v = el.get(rid_attr)
    if v and v.startswith('rId') and v not in valid_rids:
        FAIL
```

**Fix (in `backend/app.py:_sanitize_slide_xml`):** track which rIds get stripped from the slide rels, walk the slide XML, remove any `<p:ext>` whose subtree references a stripped rId.

## 2. Empty `<p:extLst>`

**Symptom:** Repair dialog after fix #1 — orphan refs cleaned up but `<p:extLst/>` left behind, especially at slide level (between `</p:cSld>` and `<p:clrMapOvr>`).

**Root cause:** When the only `<p:ext>` in an `<p:extLst>` is removed by fix #1, the parent `<extLst>` is now childless. Schema allows zero ext children, but Mac PowerPoint flags it as content corruption.

**Detection:** any `<p:extLst>` with `len(el) == 0`.

**Fix:** after stripping orphan exts, remove any `<p:extLst>` whose child count is zero.

## 3. Duplicate slide-level `<p14:creationId val="N">`

**Symptom:** Repair dialog when the same source slide is selected more than once (e.g. via the cart UI where a slide appears under two subsections).

**Root cause:** `<p14:creationId>` lives inside `<p:cSld><p:extLst>` and is meant to be globally unique per slide. Copying the slide verbatim twice produces two slides with the same `val`. Mac PowerPoint flags this; Windows is more lenient.

**Detection:**
```python
from collections import Counter
vals = [el.get('val') for el in tree.iter('{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId')]
dups = {v: c for v, c in Counter(vals).items() if c > 1}
```

**Fix:** track `used_creation_ids` across the assembly; when a duplicate is encountered, increment `val` until it's unique.

## 4. Duplicate shape-level `<a16:creationId id="{GUID}">` across slides

**Symptom:** Repair dialog persists even after fix #3, when the same source slide is duplicated.

**Root cause:** Each shape carries an `<a16:creationId>` GUID inside its `<p:cNvPr><a:extLst>`. Copying a slide twice copies its shape GUIDs. Mac PowerPoint requires these to be globally unique across the deck.

**Caveat:** within a single slide, the same shape can appear inside `<mc:AlternateContent><mc:Choice>` and `<mc:Fallback>` — both share the GUID by design (they're alternate representations of the same shape). Remapping must be consistent within a slide so the AlternateContent linkage stays intact.

**Detection:**
```python
# For each slide, collect its set of a16 GUIDs. Then check that no GUID appears
# in more than one slide's set.
A16 = 'http://schemas.microsoft.com/office/drawing/2014/main'
by_guid = {}
for slide_path in slide_paths:
    tree = etree.fromstring(z.read(slide_path))
    guids = {el.get('id') for el in tree.iter(f'{{{A16}}}creationId') if el.get('id')}
    for g in guids:
        by_guid.setdefault(g, set()).add(slide_path)
cross_slide_dups = {g: ss for g, ss in by_guid.items() if len(ss) > 1}
```

**Fix:** track `used_shape_guids` across the assembly; for each slide, build a remap from existing-duplicated GUIDs to fresh `uuid.uuid4()` values, then apply the remap to **every** occurrence in the slide (preserving mc:AlternateContent siblings).

## 5. Duplicate shape-level `<p14:modId val="N">` across slides

**Symptom:** Repair dialog persists even after fixes 1–4, when the same source slide is duplicated.

**Root cause:** Like a16:creationId, `<p14:modId>` lives in each shape's `<p:cNvPr><a:extLst>` and identifies the shape for change-tracking. Duplicate slides share modId values across the deck; Mac PowerPoint flags this.

**Detection:**
```python
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
by_val = {}
for slide in slide_paths:
    tree = etree.fromstring(z.read(slide))
    vals = {el.get('val') for el in tree.iter(f'{{{P14}}}modId') if el.get('val')}
    for v in vals:
        by_val.setdefault(v, set()).add(slide)
cross = {v: ss for v, ss in by_val.items() if len(ss) > 1}
```

**Fix:** track `used_mod_ids` across the assembly; for each slide, build a remap from already-used values to fresh integers (increment until unused), apply consistently to every occurrence in the slide. Same pattern as the a16 remap in fix #4.

## 6. Shared chart / oleObject parts referenced by two slides (suspected, unfixed)

**Symptom:** Repair dialog persists even after fixes 1–5, on a deck that contains the same source slide selected twice.

**Root cause (hypothesis):** When the assembler emits two slide XMLs that came from the same master slide, their `.rels` both point at the **same** `ppt/charts/chartN.xml`, `ppt/charts/colorsN.xml`, `ppt/charts/styleN.xml`, and `ppt/embeddings/oleObjectN.bin`. Mac PowerPoint appears to require each visible chart / OLE instance to have its own private part. Sharing is structurally legal OOXML and other readers tolerate it.

**Detection:**
```python
# For every slide rels, resolve every chart/oleObject/package target. If any
# resolved part is referenced by more than one slide, that's the trigger.
```

**Attempted fix:** rebuilt the assembly so each duplicate-source slide gets its own `chartN_dupSM.xml` (+ colors, style, rels) and its own `oleObjectN_dupSM.bin`, with rels patched accordingly. Static checks still pass. **PowerPoint still showed Repair.** Either the duplication wasn't deep enough (chart-internal IDs like `<c16:uniqueId>` may also need remapping), or the real cause is yet another cross-slide invariant we haven't isolated.

**Resolved by:** stopping the iteration. The product-level fix was to block duplicate slide selection in the cart UI (`frontend/src/hooks/useCart.js` rejects adding a slide already present). The assembly path for duplicate-source slides is now unreachable in production, so the underlying bug is dormant rather than fixed.

**Lesson:** every previous fix in this file was downstream remapping to compensate for a single upstream choice — letting the same slide be selected twice. By the time we were on the 6th identifier class to remap, the right move was to forbid the trigger, not keep patching the symptoms. If a future template (or a new product path) re-enables duplicate selection, expect this class of bug to come back, and prefer the upstream constraint.

## (placeholder for next finding)

Append the next discovery here.
