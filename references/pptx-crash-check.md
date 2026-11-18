# pptx-crash-check (вендоренная копия)

> Копия скилла `pptx-crash-check` внутри presentation-assembler, чтобы папка
> была самодостаточной (и работала в облаке). **Источник правды — скилл
> `~/.claude/skills/pptx-crash-check`** (если он установлен): новые паттерны
> порчи каталогизируются там и копируются сюда (SKILL.md → этот файл,
> known_issues.md → pptx-crash-known-issues.md, scripts/ → ../scripts/).
> Гейт build_deck гоняет `../scripts/check_pptx.py` автоматически при каждой
> сборке; этот документ нужен, когда PowerPoint показывает «Repair» на деке,
> прошедшем все проверки.
>
> **Маппинг на этот скилл.** Упоминания `backend/app.py:_sanitize_slide_xml`,
> `useCart.js` и «cart UI» ниже — контекст проекта-родителя `viscomms-2`, где
> скилл родился; здесь этих файлов нет. В presentation-assembler код сборки —
> `lib/clone.py` (ремап creationId/a16-GUID/modId при клонировании) и
> chart-guard в `scripts/build_deck.py`; «блокировка дубля upstream» из выводов
> ниже уже реализована: слайд с chart/OLE-партом не клонируется дважды.


Mac PowerPoint silently rejects .pptx files that violate certain structural rules that other readers (python-pptx, QuickLook, even Windows PowerPoint) tolerate. Этот документ + ../scripts/check_pptx.py — рабочая копия правил and how to detect them.

**Workflow when invoked:**

1. Ask the user which .pptx file to check (or take the path as `$ARGUMENTS`).
2. Run `../scripts/check_pptx.py <path>` — it reports PASS/FAIL per known check.
3. If all checks PASS but the user still sees Repair: enter the **discovery loop** below.
4. Once a new failure mode is identified and fixed in the assembly code, add the check to this skill and update `pptx-crash-known-issues.md`.

## ⚠️ First-question diagnostic

Before diving into XML diffs, ask: **does the failing deck contain the same source slide selected more than once?** Every cross-slide uniqueness bug we've cataloged (issues #3, #4, #5) — and at least one we couldn't fix at all (chart/embedding sharing) — only triggers in that scenario. The pragmatic fix is to block duplicate selection upstream (e.g. in the cart / picker UI) rather than try to remap every cross-slide identifier downstream. In `viscomms-2` this was the eventual resolution: `useCart.js` rejects adding a slide that's already in the cart. If you find yourself chasing a 4th, 5th, 6th identifier type to remap, stop and propose blocking the trigger instead.

## Known corruption causes

The full registry is in `pptx-crash-known-issues.md`. Quick reference:

| # | Cause | Detection | Fix |
|---|-------|-----------|-----|
| 1 | Orphan `r:id` references in slide XML (e.g. `<p188:commentRel r:id="rIdN"/>` after stripping comments rel) | Every `r:id` attr in a part must resolve to a `<Relationship>` in its sibling `.rels` | Strip the `<p:ext>` (or other element) whose subtree references a dropped rId |
| 2 | Empty `<p:extLst>` after stripping a child ext | An `<extLst>` with zero `<ext>` children, especially at slide-level (between `</p:cSld>` and `<p:clrMapOvr>`) | Remove the element when it becomes childless |
| 3 | Duplicate slide-level `<p14:creationId val="N">` across slides | Same `val` integer appears in more than one slide's cSld extLst | Renumber duplicates to a fresh unsigned int |
| 4 | Duplicate shape-level `<a16:creationId id="{GUID}">` across slides | Same GUID appears in shapes in more than one slide. Within a slide, the same GUID legitimately appears multiple times (mc:AlternateContent siblings share a GUID) | Remap to a fresh `uuid.uuid4()`. Apply the remap consistently within a slide so mc:AlternateContent linkage stays intact |
| 5 | Duplicate shape-level `<p14:modId val="N">` across slides | Same `val` int appears in shapes in more than one slide | Renumber duplicates the same way as #3 (track used vals, +1 until unused). Apply consistently within a slide |
| 6 | Shared chart / oleObject parts referenced by two slides (suspected, downstream-unfixable) | Resolved rels Target appears in >1 slide's rels | **Block duplicate slide selection upstream.** Downstream part-duplication (per-slide chart copy) was tried and didn't help — see issue #6 in `pptx-crash-known-issues.md` |

The above are the ones we've hit through the origin assembly code (`backend/app.py:_sanitize_slide_xml` в viscomms-2); в этом скилле те же ремапы делает `lib/clone.py`. Add new entries as we find them.

## Discovery loop (when known checks pass but PP still fails)

Mac PowerPoint's Repair dialog is generic — it doesn't say what's wrong. The only signal you get is "Repair appears" or "Repair doesn't appear". The loop:

1. **Reproduce locally and byte-identically.** Assemble the failing file from the user's exact slide selection using the local Python (`from app import assemble_pptx`). Hash-compare with what they downloaded to confirm we're debugging the same bytes.

2. **Bisect the input.** The fastest signal — if the failing file uses N slides, try halves. If 1..30 works but a specific other selection fails, you've localized to a few suspect slides or to the duplicate/cross-slide interaction. Tests we've already run:
   - Slides 1..30 from Yango master → opens clean
   - Slides [35, 46, 47, 56, 62, 64, 66] only → opens clean
   - Any selection with the same source slide twice → triggered every cross-slide-uniqueness bug so far

3. **Pull a "good" file alongside the "bad" one** and diff structurally:
   - Part-name set diff (`zipfile.namelist()`)
   - Per-part SHA-256 diff
   - For each XML part, parse with lxml and look for elements present in bad but not good
   - Specifically scan for any **cross-slide uniqueness violations**: tag names ending in `Id`, `creationId`, `modId`, GUID-shaped attribute values, etc.

4. **Form a hypothesis and isolate.** Take the suspect element type and either (a) strip it from the bad file with a minimal patch, then test in PP, or (b) inject it into a known-good file and test in PP.

5. **When fixed:**
   - Add detection to `../scripts/check_pptx.py` (a new function returning `(name, ok, details)`)
   - Append the entry to the table in this file and to `pptx-crash-known-issues.md`
   - Почини сборку в её коде: в этом скилле — `lib/clone.py` (ремап идентификаторов)
     или chart-guard в `scripts/build_deck.py` (в viscomms-2 это был
     `backend/app.py:_sanitize_slide_xml`) — with the fix and the rationale
   - Синхронизируй находку в скилл-источник `~/.claude/skills/pptx-crash-check`,
     если он установлен

## Detection-loop limits

We **cannot** programmatically detect Mac PowerPoint's Repair dialog from the CLI: System Events / osascript can't reach UI without explicit assistive-access grants, and `open -W` returns when PP exits, not when a dialog appears. So the loop has a human in it: `open <path>` → user reports outcome → static checks update.

If we ever automate this (e.g. via XCTest UI tests, AppleScript with granted accessibility), add it under `scripts/`.

## How to run the static checker

```bash
python3 $SKILL_DIR/scripts/check_pptx.py <path-to-pptx>   # вендоренная копия этого скилла
```

Exits 0 if all known checks pass; non-zero (with a summary) if any fail.
Гейт `build_deck.py` гоняет её автоматически при каждой сборке.

## Project context

This skill was born from `viscomms-2`, an assembler that copies slides out of a master `.pptx` using raw zip + lxml manipulation (no python-pptx). See `backend/app.py` for the assembly itself and `_sanitize_slide_xml` for the per-slide fixes that correspond to checks 1-4 above.
