#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""content_parity.py — сверка сохранности текста: HTML-исходник ↔ собранный .pptx.

Режим «контент не менять» (перенос готового HTML, restyle-mode.md): каждый
текстовый фрагмент исходника обязан читаться в результате. Раньше такую сверку
приходилось делать руками (фидбек переноса 2026-08-10, п.8) — скрипт сравнивает
текстовые ноды HTML (stdlib-парсер, Chrome/node НЕ нужны) с текстами
шейпов/таблиц/чартов/заметок .pptx и печатает пропуски отдельным списком.

Usage:
    python3 scripts/content_parity.py deck.pptx --html src.html [--html b.html] [--json]

Единица сверки — текстовая НОДА HTML (не блок): инлайн-разметка (<b>, .mark)
не рвёт сопоставление, нода не пересекает <br>/границы блока. Фрагменты без
букв и цифр (стрелки-декор «→») не сверяются.

Из сверки честно исключены документированные не-текстовые пути конвейера:
- `.pagenum` — зеркало номера страницы (в .pptx номер даёт лэйаут);
- потомки `[data-chart]` — заглушка браузерного превью (в .pptx — нативный
  чарт; его категории/серии/значения сверяются из chart-парта);
- `<svg>…<text>` — печатаются отдельной секцией «в растре»: содержимое видно
  на слайде картинкой, но текстом НЕ редактируется (это не пропуск, но и не
  текст — числовые ряды в SVG всё равно валит гейт A17 при сборке).

Пропуски найдены → exit 1 (после «контент не менять» дек с пропусками не
показывается — сначала верни текст или согласуй потерю списком). Всё на
месте → exit 0.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from html.parser import HTMLParser

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "lib"))

_WS = re.compile(r"\s+")


def _norm(s: str) -> str:
    return _WS.sub(" ", s.replace(" ", " ")).strip()


def _meaningful(s: str) -> bool:
    return any(ch.isalpha() or ch.isdigit() for ch in s)


class _HtmlText(HTMLParser):
    """Текстовые ноды HTML с контекстом: индекс .slide-секции, скип-зоны
    (script/style/pagenum/data-chart), svg-тексты — отдельно."""

    _VOID = {"br", "img", "meta", "link", "hr", "input", "source", "wbr"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.fragments: list[tuple[int | None, str]] = []   # (слайд, текст)
        self.raster: list[tuple[int | None, str]] = []      # svg-тексты
        self._stack: list[str] = []       # маркеры контекста ('skip'/'svg'/'')
        self._slide = -1

    def handle_starttag(self, tag, attrs):
        if tag in self._VOID:
            return
        a = dict(attrs)
        cls = a.get("class") or ""
        mark = ""
        if tag in ("script", "style") or "data-chart" in a \
                or "pagenum" in cls.split():
            mark = "skip"
        elif tag == "svg":
            mark = "svg"
        if "slide" in cls.split():
            self._slide += 1
        self._stack.append(mark)

    def handle_endtag(self, tag):
        if tag in self._VOID:
            return
        if self._stack:
            self._stack.pop()

    def handle_data(self, data):
        t = _norm(data)
        if not t or not _meaningful(t):
            return
        si = self._slide if self._slide >= 0 else None
        if "skip" in self._stack:
            return
        if "svg" in self._stack:
            self.raster.append((si, t))
        else:
            self.fragments.append((si, t))


def html_fragments(path: str):
    p = _HtmlText()
    with open(path, encoding="utf-8") as f:
        p.feed(f.read())
    return p.fragments, p.raster


def _num(v) -> str:
    f = float(v)
    return str(int(f)) if f == int(f) else str(f)


def pptx_pool(path: str) -> list[str]:
    """Строки-пулы текста дека: каждый абзац + конкатенация абзацев каждого
    text_frame (fill-ветка враппит текст многострочно — фрагмент может
    пересекать абзацы), ячейки таблиц, категории/серии/значения чартов,
    заметки спикера. Всё нормализовано."""
    from pptx import Presentation

    pool: list[str] = []

    def eat_tf(tf):
        paras = [_norm(p.text) for p in tf.paragraphs]
        pool.extend(x for x in paras if x)
        joined = _norm(" ".join(paras))
        if joined:
            pool.append(joined)

    def eat_shape(sh):
        if getattr(sh, "shapes", None) is not None and sh.shape_type is not None \
                and sh.__class__.__name__ == "GroupShape":
            for s in sh.shapes:
                eat_shape(s)
            return
        if getattr(sh, "has_text_frame", False):
            eat_tf(sh.text_frame)
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                for cell in row.cells:
                    eat_tf(cell.text_frame)
        if getattr(sh, "has_chart", False):
            ch = sh.chart
            for plot in ch.plots:
                pool.extend(_norm(str(c)) for c in plot.categories if c)
                for ser in plot.series:
                    if ser.name:
                        pool.append(_norm(str(ser.name)))
                    pool.extend(_num(v) for v in ser.values if v is not None)

    prs = Presentation(path)
    for slide in prs.slides:
        for sh in slide.shapes:
            eat_shape(sh)
        if slide.has_notes_slide:
            t = _norm(slide.notes_slide.notes_text_frame.text)
            if t:
                pool.append(t)
    return pool


def missing_fragments(fragments, pool):
    """Фрагменты, не найденные ни в одной строке пула. Учитывает типографику
    движка: фрагмент ищется и без конечной точки (правило «абзац не
    заканчивается точкой» fill-ветки)."""
    out = []
    for si, frag in fragments:
        cands = {frag}
        if frag.endswith(".") and not frag.endswith(".."):
            cands.add(frag[:-1].rstrip())
        if not any(c in line for c in cands for line in pool):
            out.append((si, frag))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--html", action="append", required=True,
                    help="HTML-исходник (можно несколько раз)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    frags, raster = [], []
    for hp in a.html:
        f, r = html_fragments(hp)
        base = os.path.basename(hp)
        frags.extend((base, si, t) for si, t in f)
        raster.extend((base, si, t) for si, t in r)

    pool = pptx_pool(a.pptx)
    miss = missing_fragments([(f"{b}[{si}]", t) for b, si, t in frags], pool)

    if a.json:
        print(json.dumps({
            "fragments": len(frags), "missing": [
                {"where": w, "text": t} for w, t in miss],
            "raster_svg": [{"where": f"{b}[{si}]", "text": t}
                           for b, si, t in raster],
        }, ensure_ascii=False, indent=1))
        return 1 if miss else 0

    print(f"PARITY: {a.pptx}")
    print(f"  фрагментов исходника: {len(frags)}; найдено: "
          f"{len(frags) - len(miss)}; пропущено: {len(miss)}")
    if raster:
        print(f"  в растре (SVG, видно картинкой, текстом не редактируется): "
              f"{len(raster)}")
        for b, si, t in raster[:10]:
            print(f"    ~ {b}[{si}] «{t[:60]}»")
    if miss:
        print("  ПРОПУСКИ (в .pptx не читаются — верни текст или согласуй "
              "потерю списком):")
        for w, t in miss:
            print(f"    ✗ {w} «{t[:80]}»")
        return 1
    print("  ✅ весь текст исходника читается в деке.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
