#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fit_check.py — ПРИМЕРКА контента deckspec ДО сборки (0.33.0; user-запрос
2026-08-07 «дек собирался почти час»). Каждый круг реворда через сборку стоит
минуты (пересборка → рендер → пере-просмотр рендеров агентом); примерка
отвечает за секунды БЕЗ .pptx: агент итерирует тексты на дешёвом чекере и
собирает дек ОДИН раз.

Меряет контент каждого слота против капасити fill-map ТОЙ ЖЕ математикой, что
fill/lint (fit.text_height_in / text_width_in):
  * высота текста vs высота коробки (по шрифту/кеглю слота, включая
    двухкегельные компоуз-карточки head_font/head_size_pt);
  * одиночный токен шире коробки (перенос по слогам «bette-/r fleet»);
  * value-слоты: ужмётся ли значение до min_pt (value_shrink);
  * значений больше, чем повторяемых шейпов у дизайнера (режь/сплить).

НЕ покрывает (мерится другими механиками — честно перечисляется в отчёте):
custom (шрифтовая проба + обмер html2pptx + превью), таблицы (trim-логика
сборки), клон-генераторы timeline/statgrid (геометрия строится сборкой).

Usage:
    python3 scripts/fit_check.py deckspec.json
Exit: 0 — всё влезает (зелёная примерка → собирай); 1 — есть переполнения.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "lib"))
sys.path.insert(0, os.path.join(SKILL_DIR, "scripts"))

import fit                                       # noqa: E402
import fill as F                                 # noqa: E402

TOL = 1.05          # допуск высоты — как fit_tolerance карт
DEF_LS = 1.0        # как default_line_spacing у slot_overflows: фактический
                    # интерлиньяж знает только донор; примерка — пре-фильтр,
                    # финальную строгость держит гейт сборки


def _lines_of(value) -> list[str]:
    if value is None or value == "":
        return []
    if isinstance(value, (list, tuple)):
        return [str(x) for x in value]
    return [str(value)]


def _fmt_pct(over: float) -> str:
    if over >= 1:
        return f"в ~{over + 1:.1f} раза"
    return f"на ~{max(5, int(round(over * 100 / 5)) * 5)}%"


def check_text(slot: dict, vals: list, tag: str, ls: float, issues: list):
    """Текстовый слот (single/repeat/compose): vals — список значений;
    значение = строка или список строк (буллеты/абзацы)."""
    font, size = slot.get("font"), slot.get("size_pt")
    bw, bh = slot.get("box_w_in"), slot.get("box_h_in")
    if not (font and size and bw):
        return
    for i, v in enumerate(vals):
        lines = _lines_of(v)
        if not lines:
            continue
        where = f"{tag}[{i}]" if len(vals) > 1 else tag
        total_h = 0.0
        for k, line in enumerate(lines):
            lf, lsz = font, float(size)
            if k == 0 and slot.get("head_size_pt"):
                lf = slot.get("head_font") or font
                lsz = float(slot["head_size_pt"])
            try:
                for tok in line.split():
                    tw = fit.text_width_in(tok, lf, lsz)
                    if tw > bw + 0.02:
                        issues.append(
                            f"✗ {where}: токен {tok!r} {tw:.2f}\" шире коробки "
                            f"{bw:.2f}\" — переносится по слогам; сократи слово "
                            f"или замени")
                total_h += fit.text_height_in(line, lf, lsz, bw, line_spacing=ls)
            except fit.FontNotFound:
                issues.append(f"⚠ {where}: шрифт {lf!r} не установлен — примерка "
                              f"слота пропущена (scripts/fetch_template.py --fonts)")
                return
        if bh and total_h > bh * TOL:
            over = total_h / bh - 1
            issues.append(
                f"✗ {where}: текст ≈{total_h:.2f}\" не влезает в коробку "
                f"{bh:.2f}\" — сократи {_fmt_pct(over)} "
                f"(сейчас {sum(len(x) for x in lines)} знаков)")


def check_value(slot: dict, vals: list, tag: str, issues: list):
    """value_shrink: влезет ли значение хотя бы на min_pt."""
    font, size = slot.get("font"), slot.get("size_pt")
    bw = slot.get("box_w_in")
    minpt = float(slot.get("min_pt", 24))
    if not (font and size and bw):
        return
    for i, v in enumerate(vals):
        s = str(v).strip()
        if not s:
            continue
        where = f"{tag}[{i}]" if len(vals) > 1 else tag
        try:
            w = fit.text_width_in(s, font, float(size))
        except fit.FontNotFound:
            issues.append(f"⚠ {where}: шрифт {font!r} не установлен — примерка "
                          f"слота пропущена")
            return
        if w <= bw:
            continue
        need_pt = float(size) * bw / w
        if need_pt >= minpt - 0.1:
            continue                      # ужмётся штатным value_shrink — ок
        max_chars = max(1, int(len(s) * bw / (w * minpt / float(size))))
        issues.append(
            f"✗ {where}: значение {s!r} не влезает даже на {minpt:g}pt — "
            f"сократи до ~{max_chars} знаков (напр. формат 22.4K вместо 22,400)")


def check_plan(plan: dict, issues: list, notes: list):
    fam = plan["family"]
    mode = plan.get("mode")
    if mode == "CUSTOM":
        notes.append(f"{fam}: custom — примеркой не мерится (его стерегут "
                     f"шрифтовая проба, поля A15 и превью)")
        return
    if mode != "FILL" or not plan.get("variant"):
        notes.append(f"{fam}: {plan.get('generator') or mode} — геометрию "
                     f"строит сборка, примерка не мерит")
        return
    variant = plan["variant"]
    content = plan["content"]
    slotvals = getattr(content, "slots", None) or {}
    ls = DEF_LS
    for name, slot in variant.get("slots", {}).items():
        tag = f"{fam}/{variant.get('name', '?')}: {name}"
        if slot.get("kind") == "table":
            notes.append(f"{tag} — таблица: ужимается trim-логикой сборки")
            continue
        if slot.get("compose"):
            vals = F._compose_slot_values(slotvals, content, slot["compose"])
        else:
            v = slotvals.get(name, getattr(content, name, None))
            vals = list(v) if isinstance(v, (list, tuple)) and slot.get("repeat") \
                else ([v] if v not in (None, "") else [])
        if not vals:
            continue
        mn = slot.get("match_names")
        if slot.get("repeat") and mn and len(vals) > len(mn):
            issues.append(f"✗ {tag}: значений {len(vals)} > слотов дизайнера "
                          f"{len(mn)} — режь контент или сплить слайд")
        if slot.get("fit_policy") == "value_shrink":
            check_value(slot, vals, tag, issues)
        else:
            check_text(slot, vals, tag, ls, issues)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deckspec")
    a = ap.parse_args()

    import build_deck as B                        # тяжёлый импорт — после argparse
    from pptx import Presentation

    with open(a.deckspec, encoding="utf-8") as f:
        ds = json.load(f)
    spec_dir = os.path.dirname(os.path.abspath(a.deckspec))
    tdir = B.TDIR_ALIASES.get(ds.get("template", "еда").lower(),
                              ds.get("template", "еда").lower())
    if tdir not in B.LAYOUT_DIRS:
        print(f"неизвестный шаблон {tdir!r}")
        return 1
    prs = Presentation(B.template_path(tdir))

    total_issues = 0
    for i, entry in enumerate(ds["slides"], 1):
        entry = B.typo_walk(dict(entry))          # та же типографика, что у сборки
        issues, notes = [], []
        if entry.get("family") == "custom":
            notes.append("custom — примеркой не мерится (шрифтовая проба, "
                         "поля A15, превью)")
        else:
            try:
                plan = B.resolve_entry_plan(prs, entry, tdir)
                check_plan(plan, issues, notes)
            except B.DeckError as e:
                issues.append(f"✗ резолв: {e}")
        if issues or notes:
            print(f"слайд {i} [{entry.get('family')}]")
            for msg in issues:
                print(f"  {msg}")
            for msg in notes:
                print(f"  · {msg}")
        total_issues += len([x for x in issues if x.startswith("✗")])

    if total_issues:
        print(f"\nпримерка: {total_issues} переполнений — правь тексты по "
              f"подсказкам и повтори; сборку запускай после зелёной примерки")
        return 1
    print("\nпримерка зелёная: контент влезает — собирай "
          "(scripts/build_deck.py), реворд-циклов не ожидается")
    return 0


if __name__ == "__main__":
    sys.exit(main())
