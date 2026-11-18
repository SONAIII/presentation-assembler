#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""anchor.py — привязка слайда шаблона по СИГНАТУРЕ (тексту-заглушке), не по индексу.

Обновление шаблона сдвигает позиции → `slide_index` устаревает (§3Б, §5). resolve_slide
находит нужный дизайнерский слайд по стабильным признакам:
  * all       — строки-заглушки, которые ДОЛЖНЫ быть на слайде (exact line);
  * none      — строки, которых быть НЕ должно (разводит слайд-надмножество);
  * has_shapes / no_shapes — имена шейпов, что есть / нет (разводит близнецов);
  * has_picture — есть ли на слайде картинка (photo-вариант vs plain).

Матчинг ПОСТРОЧНЫЙ (exact line), а НЕ подстрокой — иначе «…должность» поймает и
«…должность или дату выступления». Ошибки ГРОМКИЕ: при апдейте шаблона сигнатура должна
упасть заметно, а не молча заполнить чужой слайд. Драйвит FILL-ветку роутера (route.decide*).
"""
from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE


def _iter_shapes(shapes):
    """Плоский обход с заходом в группы."""
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(sh.shapes)


def slide_sig(slide):
    """(множество строк, множество имён шейпов, есть_ли_фото) — для матчинга сигнатуры."""
    lines, names, has_pic = set(), set(), False
    for sh in _iter_shapes(slide.shapes):
        names.add(sh.name)
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            has_pic = True
        txt = ""
        if sh.has_text_frame:
            txt = sh.text_frame.text
        elif sh.has_table:
            txt = "\n".join(c.text for row in sh.table.rows for c in row.cells)
        for ln in txt.splitlines():
            ln = ln.strip()
            if ln:
                lines.add(ln)
    return lines, names, has_pic


def _matches(anchor: dict, lines: set, names: set, has_pic: bool) -> bool:
    for s in anchor.get("all", []) or anchor.get("signature", []):
        if s not in lines:
            return False
    for s in anchor.get("none", []):
        if s in lines:
            return False
    for n in anchor.get("has_shapes", []):
        if n not in names:
            return False
    for n in anchor.get("no_shapes", []):
        if n in names:
            return False
    if "has_picture" in anchor and bool(anchor["has_picture"]) != has_pic:
        return False
    return True


def resolve_slide(prs, anchor: dict, *, hint=None) -> int:
    """Индекс (0-based) слайда, чья сигнатура совпала. 1 матч → он; >1 → _hint как
    тайбрейкер (иначе ошибка неоднозначности); 0 → ошибка дрейфа."""
    hint = anchor.get("_hint", hint)
    matches = []
    for i, slide in enumerate(prs.slides):
        if _matches(anchor, *slide_sig(slide)):
            matches.append(i)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        if hint in matches:
            return hint
        raise ValueError(
            f"сигнатура {anchor} неоднозначна: слайды {[m + 1 for m in matches]} "
            f"(1-based); уточни all/none/has_shapes/has_picture или задай _hint")
    raise ValueError(
        f"сигнатура {anchor} не найдена ни на одном слайде (дрейф шаблона? _hint={hint}); "
        f"перезапусти scripts/inventory_template.py и обнови anchor")


def resolve_entry(entry: dict, prs) -> int:
    """Общий резолвер для записи спеки/варианта: anchor → по сигнатуре; иначе legacy
    slide_index (0-based, PoC-фолбэк, помечен _index_is_poc_debt)."""
    if entry.get("anchor"):
        # _hint (вариант-уровень поля fill-map) — тайбрейкер для доноров
        # с ИДЕНТИЧНОЙ сигнатурой (тема/акцент-варианты gor-services). anchor-встроенный
        # _hint (если есть) перекрывает внутри resolve_slide. Еда: сигнатуры уникальны →
        # matches==1 → hint не читается → байт-в-байт.
        return resolve_slide(prs, entry["anchor"],
                             hint=entry.get("_hint", entry.get("slide_index")))
    if "slide_index" in entry:
        return entry["slide_index"]
    raise ValueError(f"запись без anchor и без slide_index: {entry}")
