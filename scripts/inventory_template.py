#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inventory_template.py — read-only снимок реальной структуры шаблона (§3Б, FILL-ALL).

Зачем: обновление шаблона сдвигает позиции слайдов → все `slide_index` в спеках
устаревают. Чтобы переанкорить привязку с индекса на СИГНАТУРУ (текст-заглушку),
нужно точно знать, какой текст сейчас на каком слайде и какие строки УНИКАЛЬНЫ
(годятся в дискриминатор сигнатуры), а какие дублируются (Tier B — команда/воронка).

Выдаёт:
  * per-slide: индекс (0- и 1-based), все шейпы с полным текстом (+ таблицы/чарты/фото),
  * `unique_strings`: строки, встречающиеся РОВНО на одном слайде → кандидат в сигнатуру,
  * `dup_shapes`: одинаковые тексты-шейпы внутри слайда (≥2) → нужен occ/repeat,
  * машинный JSON (`--out`) + человекочитаемая сводка в stdout.

  python3 scripts/inventory_template.py                 # Еда, сводка + JSON
  python3 scripts/inventory_template.py --json-only      # только записать JSON
  python3 scripts/inventory_template.py --diff old.json  # drift: что сдвинулось/поменялось
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_TEMPLATE = os.path.join(ROOT, "Templates/eda/eda-template.pptx")
DEFAULT_OUT = os.path.join(ROOT, "layouts/eda/_inventory.json")


def shape_text(sh) -> str:
    """Полный текст шейпа: text-frame, либо ячейки таблицы (по строкам)."""
    if sh.has_text_frame:
        return sh.text_frame.text
    if sh.has_table:
        rows = []
        for row in sh.table.rows:
            rows.append("\t".join(c.text for c in row.cells))
        return "\n".join(rows)
    return ""


def shape_kind(sh) -> str:
    if sh.has_table:
        return "table"
    if sh.has_chart:
        return "chart"
    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if sh.has_text_frame and sh.text_frame.text.strip():
        return "text"
    return "shape"


def walk_shapes(shapes):
    """Плоский обход с заходом в группы (группа тоже отдаётся — для контекста)."""
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(sh.shapes)


def uneditable_shapes(slide) -> list:
    """Шейпы, чей текст НЕЛЬЗЯ надёжно править программно (идея детекции — у
    thinkcell.py из community/presentation): (1) think-cell OLE — p:oleObj с
    «think-cell» в name: картинка рендерится из датасета think-cell, правка
    pptx-текста до экрана не доходит; (2) текстовые шейпы внутри
    mc:AlternateContent — рисуется ветка Choice, python-pptx правит Fallback.
    Это контрольная лампочка после ручной вычистки шаблона: на помеченных
    шейпах слоты fill-map не заводить (dont_touch), пока лампочка не погасла."""
    root = slide._element
    found, seen = [], set()
    for ole in root.iter(qn("p:oleObj")):
        nm = ole.get("name") or ""
        if "think-cell" in nm.lower() and ("ole", nm) not in seen:
            seen.add(("ole", nm))
            found.append({"kind": "thinkcell-ole", "name": nm})
    for ac in root.iter(f"{MC_NS}AlternateContent"):
        for sp in ac.iter(qn("p:sp")):
            if sp.find(f".//{qn('a:t')}") is None:
                continue                       # без текста — нам не мешает
            nv = sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
            nm = nv.get("name") if nv is not None else "?"
            if ("ac", nm) in seen:             # Choice/Fallback — один шейп
                continue
            seen.add(("ac", nm))
            found.append({"kind": "alternate-content", "name": nm})
    return found


def collect_slide(idx0, slide) -> dict:
    shapes = []
    for sh in walk_shapes(slide.shapes):
        txt = shape_text(sh)
        shapes.append({
            "name": sh.name,
            "kind": shape_kind(sh),
            "text": txt,
        })
    all_text = "\n".join(s["text"] for s in shapes if s["text"])
    # дубли внутри слайда (одинаковый непустой текст ≥2) → Tier B (occ/repeat)
    counts = {}
    for s in shapes:
        t = s["text"].strip()
        if t:
            counts[t] = counts.get(t, 0) + 1
    dup = [{"text": t, "count": c} for t, c in counts.items() if c >= 2]
    return {
        "index0": idx0,
        "index1": idx0 + 1,
        "shapes": shapes,
        "all_text": all_text,
        "dup_shapes": sorted(dup, key=lambda d: -d["count"]),
        "uneditable": uneditable_shapes(slide),
    }


def slide_lines(sl) -> set:
    """Набор непустых строк слайда (для расчёта уникальности между слайдами)."""
    lines = set()
    for s in sl["shapes"]:
        for ln in s["text"].splitlines():
            ln = ln.strip()
            if ln:
                lines.add(ln)
    return lines


def _portable_path(path: str) -> str:
    """Путь в снимке — переносимый (снимок коммитится как бейзлайн дрифт-чека):
    относительный от корня скилла, для кэша/чужих мест — только имя файла."""
    ap = os.path.abspath(path)
    return (os.path.relpath(ap, ROOT) if ap.startswith(ROOT + os.sep)
            else os.path.basename(ap))


def build_inventory(template: str) -> dict:
    prs = Presentation(template)
    slides = [collect_slide(i, s) for i, s in enumerate(prs.slides)]

    # частота строки ПО СЛАЙДАМ (на скольких разных слайдах встречается)
    freq = {}
    per_slide_lines = [slide_lines(sl) for sl in slides]
    for lines in per_slide_lines:
        for ln in lines:
            freq[ln] = freq.get(ln, 0) + 1
    for sl, lines in zip(slides, per_slide_lines):
        # (len, str) — детерминированный порядок между процессами (set-итерация
        # плавает; key=len не различала строки равной длины)
        uniq = sorted([ln for ln in lines if freq[ln] == 1],
                      key=lambda s: (len(s), s))
        sl["unique_strings"] = uniq

    return {
        "template": _portable_path(template),
        "mtime": time.ctime(os.path.getmtime(template)),
        "slide_count": len(slides),
        "slides": slides,
    }


def print_summary(inv: dict):
    print(f"template: {inv['template']}")
    print(f"mtime:    {inv['mtime']}")
    print(f"slides:   {inv['slide_count']}\n")
    for sl in inv["slides"]:
        head = ""
        for s in sl["shapes"]:
            t = s["text"].strip().replace("\n", " ")
            if t:
                head = t[:70]
                break
        tags = []
        kinds = {s["kind"] for s in sl["shapes"]}
        for k in ("table", "chart", "picture"):
            if k in kinds:
                tags.append(k.upper())
        tagstr = (" [" + ",".join(tags) + "]") if tags else ""
        print(f"#{sl['index1']:>2}{tagstr} | {head}")
        uq = sl["unique_strings"]
        if uq:
            # самый короткий уникальный текст — лучший дискриминатор
            print(f"       uniq✓: «{uq[0]}»" + (f"  (+{len(uq)-1} more)" if len(uq) > 1 else ""))
        else:
            print("       uniq✗: НЕТ уникальных строк — сигнатуре нужен all+none")
        for d in sl["dup_shapes"]:
            print(f"       dup×{d['count']}: «{d['text'][:40].replace(chr(10),' ')}»")
        un = sl.get("uneditable") or []
        if un:
            kinds = {}
            for u in un:
                kinds[u["kind"]] = kinds.get(u["kind"], 0) + 1
            det = ", ".join(f"{k}×{c}" for k, c in sorted(kinds.items()))
            print(f"       ⚠ uneditable: {len(un)} шейп(ов) ({det}) — текст "
                  f"правится только в think-cell; слоты на них не заводить")


def do_diff(inv: dict, old_path: str):
    if not os.path.exists(old_path):
        print(f"\n[diff] базовый снимок {old_path} не найден — сравнивать не с чем.")
        return
    with open(old_path, encoding="utf-8") as f:
        old = json.load(f)
    print(f"\n[diff] {old.get('slide_count')} → {inv['slide_count']} слайдов")
    # сопоставляем слайды по первой уникальной строке (стабильный ключ)
    def key_of(sl):
        uq = sl.get("unique_strings")
        # min по (len, str): sorted(key=len) не различает строки равной длины, а
        # set-итерация нестабильна между процессами → дифф дребезжал бы на тай-брейках
        return min(uq, key=lambda s: (len(s), s)) if uq else sl["all_text"][:40]
    old_pos = {key_of(sl): sl["index1"] for sl in old["slides"]}
    new_pos = {key_of(sl): sl["index1"] for sl in inv["slides"]}
    moved = [(k, old_pos[k], new_pos[k]) for k in new_pos if k in old_pos and old_pos[k] != new_pos[k]]
    added = [k for k in new_pos if k not in old_pos]
    removed = [k for k in old_pos if k not in new_pos]
    for k, o, n in sorted(moved, key=lambda x: x[2]):
        print(f"  MOVED  {o:>2} → {n:>2}  «{k[:50]}»")
    for k in added:
        print(f"  ADDED  → {new_pos[k]:>2}  «{k[:50]}»")
    for k in removed:
        print(f"  REMOVED (was {old_pos[k]:>2})  «{k[:50]}»")
    if not (moved or added or removed):
        print("  без изменений сигнатур.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--json-only", action="store_true", help="не печатать сводку")
    ap.add_argument("--diff", metavar="OLD.json", help="сравнить с прошлым снимком (drift)")
    a = ap.parse_args()

    if a.template == DEFAULT_TEMPLATE and not os.path.exists(a.template):
        import fetch_template   # дистрибуция без .pptx: кэш → Я.Диск по пину
        a.template = fetch_template.resolve_template("eda")

    inv = build_inventory(a.template)
    if not a.json_only:
        print_summary(inv)
    if a.diff:
        do_diff(inv, a.diff)

    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(inv, f, ensure_ascii=False, indent=2)
    print(f"\n[inventory] JSON → {a.out}  ({inv['slide_count']} слайдов)")


if __name__ == "__main__":
    main()
