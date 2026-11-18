#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""library_from_pptx.py — pptx-библиотека ассетов → полка именованных файлов.

Дизайнеры ведут библиотеки (логотипы, флаги, иконки) слайдами: картинка +
подпись под ней. Для агента рабочий формат — файлы с говорящими именами
(выбор логотипа идёт по имени сервиса, а не разглядыванием сотни картинок).
Конвертер (user-решение 2026-07-07):

  * имя = ПОДПИСЬ ПОД КАРТИНКОЙ — пространственный матчинг (подпись центром
    под картинкой, зазор до 1.2"); проверено на «Логотипы сервисов Яндекса»:
    92/99 имён автоматически;
  * SVG-пары PowerPoint: <asvg:svgBlip> (мастер) + PNG-фолбэк — забираем ОБА
    (<имя>.png в слоты, <имя>.svg мастером);
  * дубли имён → суффиксы _2/_3; безымянные → _unnamed/ на ручной разбор;
  * manifest.json (имя, слайд, px, аспект, svg) — рядом с файлами.

Результат готов к выкладке на полку Я.Диска в assets/<категория>/
(скачивание — scripts/fetch_asset.py).

  sh scripts/run.sh scripts/library_from_pptx.py "Логотипы сервисов Яндекса.pptx" -o /tmp/logos
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import media_catalog as MC                        # noqa: E402  (_px, флаги)

EMU = 914400
ASVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"
R_EMBED = qn("r:embed")
STOP_LABELS = {"soon"}                            # заглушки «скоро будет»
MAX_LABEL_LEN = 40      # «Центральноафриканская Республика» = 32 символа


def sanitize(name: str) -> str:
    name = re.sub(r"[\x0b\r\n\t]+", " ", name)
    name = re.sub(r'[\\/:*?"<>|]+', "-", name)
    return re.sub(r"\s+", " ", name).strip()


def match_labels(pics, labels):
    """pic → подпись. Дизайнеры ведут библиотеки двумя раскладками:
    СНИЗУ (логотипы: центр подписи под картинкой, |Δx| < 0.8·w, зазор
    -0.1..1.2") и СПРАВА (флаги, 2026-07-07: подпись начинается у правого
    края картинки, зазор до 0.8", вертикальный центр в y-полосе картинки).
    Ближайшая по суммарному расстоянию побеждает."""
    out = {}
    for i, (sh, pl, pt, pw, ph) in enumerate(pics):
        pcx, pcy, pbot, pr = pl + pw / 2, pt + ph / 2, pt + ph, pl + pw
        best, bestd = None, 9e9
        for t, ll, lt, lw, lh in labels:
            dx = abs((ll + lw / 2) - pcx)
            if dx < pw * 0.8 and -0.1 < lt - pbot < 1.2:       # СНИЗУ (логотипы)
                d = dx + (lt - pbot)
                if d < bestd:
                    best, bestd = t, d
            gap, lcy = ll - pr, lt + lh / 2
            if -0.05 < gap < 0.8 and pt - 0.1 < lcy < pbot + 0.1:  # СПРАВА (флаги)
                d = gap + abs(lcy - pcy)
                if d < bestd:
                    best, bestd = t, d
        out[i] = best
    return out


def _blip_parts(sh, slide):
    """(png_part, svg_part|None) картинки: PNG-фолбэк из a:blip, SVG-мастер
    из расширения asvg:svgBlip (PowerPoint хранит вектор парой)."""
    blip = sh._element.find(f".//{qn('a:blip')}")
    if blip is None or not blip.get(R_EMBED):
        return None, None
    png = slide.part.rels[blip.get(R_EMBED)].target_part
    svg = None
    svg_el = blip.find(f".//{ASVG_BLIP}")
    if svg_el is not None and svg_el.get(R_EMBED) in slide.part.rels:
        svg = slide.part.rels[svg_el.get(R_EMBED)].target_part
    return png, svg


def convert(path: str, outdir: str) -> dict:
    prs = Presentation(path)
    os.makedirs(outdir, exist_ok=True)
    un_dir = os.path.join(outdir, "_unnamed")
    taken, items, grouped = {}, [], 0

    for si, slide in enumerate(prs.slides, 1):
        pics, labels = [], []
        for sh in slide.shapes:
            st = int(sh.shape_type) if sh.shape_type is not None else -1
            if st == 13:
                pics.append((sh, sh.left / EMU, sh.top / EMU,
                             sh.width / EMU, sh.height / EMU))
            elif st == 6:
                grouped += sum(1 for s in sh.shapes if int(s.shape_type or -1) == 13)
            elif sh.has_text_frame:
                t = sanitize(sh.text_frame.text)
                if t and len(t) <= MAX_LABEL_LEN and t.lower() not in STOP_LABELS:
                    labels.append((t, sh.left / EMU, sh.top / EMU,
                                   sh.width / EMU, (sh.height or 0) / EMU))
        names = match_labels(pics, labels)
        for i, (sh, *_geo) in enumerate(pics):
            png, svg = _blip_parts(sh, slide)
            if png is None:
                continue
            name = names[i]
            if name:
                n = taken.get(name, 0) + 1
                taken[name] = n
                base = name if n == 1 else f"{name}_{n}"
                dst_dir = outdir
            else:
                base = f"s{si:02d}_{os.path.basename(str(png.partname))}"
                base = os.path.splitext(base)[0]
                dst_dir = un_dir
                os.makedirs(un_dir, exist_ok=True)
            ext = os.path.splitext(str(png.partname))[1].lower() or ".png"
            with open(os.path.join(dst_dir, base + ext), "wb") as f:
                f.write(png.blob)
            if svg is not None:
                with open(os.path.join(dst_dir, base + ".svg"), "wb") as f:
                    f.write(svg.blob)
            px = MC._px(png.blob)
            items.append({"file": base + ext, "named": bool(name), "slide": si,
                          "px": px,
                          "aspect": round(px[0] / px[1], 3) if px and px[1] else None,
                          "kb": round(len(png.blob) / 1024),
                          "svg": svg is not None})
    manifest = {"source": os.path.abspath(path), "outdir": os.path.abspath(outdir),
                "images": items}
    with open(os.path.join(outdir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    named = [i for i in items if i["named"]]
    print(f"извлечено {len(items)}: с именами {len(named)}, "
          f"безымянных {len(items) - len(named)} (_unnamed/ — разобрать руками), "
          f"svg-пар {sum(1 for i in items if i['svg'])}")
    dupes = {k: v for k, v in taken.items() if v > 1}
    if dupes:
        print(f"дубли имён (варианты одного лого, суффиксы _2/_3): {dupes}")
    if grouped:
        print(f"⚠ {grouped} картинок внутри ГРУПП пропущено — разгруппируй "
              f"в мастер-файле и перегони")
    print(f"→ {outdir} (+ manifest.json); выкладку на полку делает "
          f"команда скилла")
    return manifest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", help="pptx-библиотека (картинки с подписями)")
    ap.add_argument("-o", "--out", required=True, help="папка результата")
    a = ap.parse_args()
    convert(a.source, a.out)


if __name__ == "__main__":
    main()
