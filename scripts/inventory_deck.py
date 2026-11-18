#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inventory_deck.py — инвентарь ЧУЖОГО дека для режима «перевёрстка».

Не путать с inventory_template.py (снимок НАШЕГО шаблона для якорей): здесь
вход — произвольная пользовательская презентация («причеши мой дек»), выход —
сырьё для сопоставления слайдов семействам шаблона (references/restyle-mode.md):

  per-slide: заголовок (placeholder/крупнейший кегль), ВСЕ тексты в порядке
  чтения (полные, с абзацами — из них собирается deckspec ДОСЛОВНО), таблицы
  (размерность), чарты, счёт картинок, нередактируемое (think-cell OLE /
  AlternateContent), заметки спикера (переносятся в deckspec "notes").

Картинки этот скрипт только СЧИТАЕТ — извлечение и опись делает
media_catalog.py (запускать парой).

Использование:
  python3 scripts/inventory_deck.py deck.pptx            # сводка
  python3 scripts/inventory_deck.py deck.pptx --json     # JSON в stdout
  python3 scripts/inventory_deck.py deck.pptx -o inv.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import inventory_template as ITV                  # noqa: E402  (shape_text, uneditable)

EMU = 914400


def _is_title(sh) -> bool:
    """Титул = плейсхолдер idx 0 (инвариант pptx-раскладок; enum-значения
    PP_PLACEHOLDER_TYPE у python-pptx свои, по ним не сравниваем)."""
    try:
        return sh.is_placeholder and sh.placeholder_format.idx == 0
    except Exception:
        return False


def _max_run_pt(sh):
    pts = [r.font.size.pt for p in sh.text_frame.paragraphs
           for r in p.runs if r.font.size]
    return max(pts) if pts else None


def _text_entry(sh) -> dict:
    paras = [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()]
    txt = "\n".join(paras)
    return {"text": txt, "paras": paras, "pt": _max_run_pt(sh),
            "words": len(txt.split())}


def _iter_reading_order(shapes):
    """Топ-левел по (top, left) — в чужих деках порядок в XML произволен;
    дети групп — в порядке группы (их координаты в child-space)."""
    def key(sh):
        try:
            return (sh.top or 0, sh.left or 0)
        except Exception:
            return (0, 0)
    for sh in sorted(shapes, key=key):
        if sh.shape_type == 6:                    # GROUP
            yield from sh.shapes
        else:
            yield sh


def collect_slide(idx1: int, slide) -> dict:
    title, texts, table, n_charts, n_pics = None, [], None, 0, 0
    for sh in _iter_reading_order(slide.shapes):
        st = int(sh.shape_type) if sh.shape_type is not None else -1
        if st == 13:                              # PICTURE
            n_pics += 1
            continue
        if getattr(sh, "has_chart", False):
            n_charts += 1
            continue
        if getattr(sh, "has_table", False):
            t = sh.table
            table = {"rows": len(t.rows), "cols": len(t.columns),
                     "cells": [[c.text.strip() for c in row.cells]
                               for row in t.rows]}
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            entry = _text_entry(sh)
            if title is None and _is_title(sh):
                title = entry["text"]
                continue
            texts.append(entry)
    if title is None and texts:                   # нет TITLE-плейсхолдера —
        big = max(texts, key=lambda e: e["pt"] or 0)   # берём крупнейший кегль
        if big["pt"]:
            title = big["text"]
            texts.remove(big)
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    bullets = sum(len(e["paras"]) for e in texts)
    hint = ("chart" if n_charts else "table" if table else
            "pictures" if n_pics >= 2 and bullets <= 2 else
            "statement" if title and bullets == 0 else
            "bullets" if bullets >= 3 else "mixed")
    return {"index1": idx1, "title": title, "texts": texts,
            "bullets_total": bullets, "table": table, "n_charts": n_charts,
            "n_pictures": n_pics, "uneditable": ITV.uneditable_shapes(slide),
            "notes": notes, "hint": hint}


def build_inventory(path: str) -> dict:
    prs = Presentation(path)
    slides = [collect_slide(i, sl) for i, sl in enumerate(prs.slides, 1)]
    return {"source": os.path.abspath(path), "slide_count": len(slides),
            "slides": slides}


def print_summary(inv: dict):
    print(f"ИНВЕНТАРЬ ДЕКА: {inv['source']}  ({inv['slide_count']} слайдов)")
    print("-" * 72)
    for sl in inv["slides"]:
        bits = [f"hint={sl['hint']}"]
        if sl["bullets_total"]:
            bits.append(f"абзацев {sl['bullets_total']}")
        if sl["table"]:
            bits.append(f"таблица {sl['table']['rows']}×{sl['table']['cols']}")
        if sl["n_charts"]:
            bits.append(f"чартов {sl['n_charts']}")
        if sl["n_pictures"]:
            bits.append(f"картинок {sl['n_pictures']}")
        if sl["notes"]:
            bits.append("есть заметки")
        print(f"#{sl['index1']:>2} «{(sl['title'] or '—')[:44]}»  "
              f"[{', '.join(bits)}]")
        un = sl.get("uneditable") or []
        if un:
            det = ", ".join(sorted({u["kind"] for u in un}))
            print(f"      ⚠ нередактируемое: {det} — как есть НЕ переносится "
                  f"(чарт пересоздаётся из данных, иначе слайд адаптируется)")
    print("\nдальше по references/restyle-mode.md: сопоставление семействам "
          "(карта §1 content-assembly-rules) → storyline → deckspec ДОСЛОВНО; "
          "картинки исходника — media_catalog.py")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", help="чужой .pptx")
    ap.add_argument("--json", action="store_true", help="JSON в stdout")
    ap.add_argument("-o", "--out", default=None, help="записать JSON в файл")
    a = ap.parse_args()
    inv = build_inventory(a.deck)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(inv, f, ensure_ascii=False, indent=1)
        print(f"JSON → {a.out}")
    if a.json:
        print(json.dumps(inv, ensure_ascii=False, indent=1))
    elif not a.out:
        print_summary(inv)


if __name__ == "__main__":
    main()
