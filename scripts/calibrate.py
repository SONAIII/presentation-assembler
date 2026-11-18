#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""calibrate.py — вывести калибровочный K[font] из РЕАЛЬНОГО рендера PowerPoint.

Зачем (plan §12): Pillow меряет ширину по advance-глифам, но PowerPoint рисует
текст иначе (трекинг/хинтинг/округление) — на дисплейных кеглях расхождение
заметное. Если зазор подбирать руками, мы возвращаемся к ручной итерации, от
которой и уходим. Поэтому K выводим машинно: один раз печём пробный дек, рендерим
его как реальный PowerPoint, меряем фактическую ширину глифов в пикселях.

Метод: на каждом слайде ровно ОДИН чёрный текст-бокс (wrap=off) на белом фоне.
В PNG берём bounding-box тёмных пикселей -> ширина в дюймах. K = real_in / raw_in
(raw_in — fit.raw_width_in). Инсеты бокса не влияют: они сдвигают текст, но не его
ширину, а мы меряем сами глифы.

Результат: lib/calibration.json = { "_default": K, "fonts": { name: K } }.

Запуск:
    python3 scripts/calibrate.py            # печь дек, рендерить, измерить, записать
    python3 scripts/calibrate.py --dry      # только напечатать замеры, не писать json
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from lib import fit  # noqa: E402

DPI = 200
SLIDE_W_IN = 26.665
SLIDE_H_IN = 15.0
LEFT_IN = 1.0
BOX_W_IN = 24.0
# Рендерим через /private/tmp: PowerPoint-песочница читает её без Full Disk Access,
# а ~/Documents (TCC-защищён) даёт -9074, если FDA слетел (см. render-powerpoint-fda).
WORKDIR = "/private/tmp/_calib_pptx"
CALIB_JSON = os.path.join(ROOT, "lib", "calibration.json")

# Пробы: (font_name, [строки], [кегли]). Строки и кегли подобраны под реальные
# слоты Еды (числа/проценты/×/запятая на дисплейных кеглях; кириллица в тексте),
# по несколько на шрифт — чтобы K был усреднён и проверена его кегле-независимость.
PROBES = {
    "YS Geo Heavy": (
        ["40%", "193", "32%", "4,8", "2×", "1280", "88,5%", "ЗАКАЗ"],
        [138, 96, 72],
    ),
    "YS Geo Bold": (
        ["повторили заказ за неделю", "быстрее оформление", "1280", "Сравнение"],
        [44, 34, 28],
    ),
    "YS Text Regular": (
        ["меньше шагов до заказа", "Почти все повторяют недавние заказы", "0123456789"],
        [44, 32, 24],
    ),
    "YS Compressed Heavy": (
        ["ГОРОДСКИЕ СЕРВИСЫ", "40%", "1280", "ИТОГИ КВАРТАЛА"],
        [120, 80, 48],
    ),
    "YS SemiMono Regular": (
        ["2026 Q4", "+128%", "metric_value"],
        [40, 28, 20],
    ),
}


def build_probe_deck(path: str) -> list[dict]:
    """Печём дек: слайд на пробу. Возвращаем список {slide, text, font, size_pt}."""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_AUTO_SIZE

    prs = Presentation()
    prs.slide_width = Emu(int(SLIDE_W_IN * fit.EMU_PER_INCH))
    prs.slide_height = Emu(int(SLIDE_H_IN * fit.EMU_PER_INCH))
    blank = prs.slide_layouts[6]
    manifest: list[dict] = []
    idx = 0
    for font_name, (strings, sizes) in PROBES.items():
        if not fit.resolve_font(font_name):
            sys.stderr.write(f"calibrate: пропуск {font_name!r} — нет файла шрифта\n")
            continue
        for s in strings:
            for sz in sizes:
                slide = prs.slides.add_slide(blank)
                # blank-layout дефолтного шаблона уже белый — отдельный фон не льём
                # (ручная XML-инъекция фона рискует repair-диалогом PowerPoint)
                top_in = (SLIDE_H_IN - sz / 72.0) / 2.0
                tb = slide.shapes.add_textbox(
                    Inches(LEFT_IN), Inches(top_in), Inches(BOX_W_IN), Inches(sz / 72.0 * 1.6)
                )
                tf = tb.text_frame
                tf.word_wrap = False
                tf.auto_size = MSO_AUTO_SIZE.NONE
                for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                    setattr(tf, m, 0)
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = s
                run.font.name = font_name
                run.font.size = Pt(sz)
                run.font.color.rgb = RGBColor(0, 0, 0)
                manifest.append(
                    {"slide": idx, "text": s, "font": font_name, "size_pt": sz}
                )
                idx += 1
    prs.save(path)
    return manifest


def render(pptx_path: str, outdir: str) -> list[str]:
    """Рендер через общий render.py (PowerPoint→PDF→PNG)."""
    r = subprocess.run(
        [sys.executable, os.path.join(HERE, "render.py"),
         pptx_path, "--dpi", str(DPI), "-o", outdir],
        capture_output=True, text=True,
    )
    sys.stderr.write(r.stderr)
    pngs = [l for l in r.stdout.splitlines() if l.endswith(".png")]
    return sorted(pngs)


def dark_bbox_width_in(png: str) -> float | None:
    """Ширина bounding-box тёмных пикселей в дюймах (или None, если пусто)."""
    from PIL import Image

    im = Image.open(png).convert("L")
    # бинаризуем: тёмное = глиф. bbox инвертированного даёт рамку тёмного.
    from PIL import ImageOps

    bw = im.point(lambda v: 0 if v > 160 else 255)  # тёмное -> 255
    box = bw.getbbox()
    if not box:
        return None
    left, _, right, _ = box
    return (right - left) / DPI


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true", help="не записывать calibration.json")
    a = ap.parse_args()

    os.makedirs(WORKDIR, exist_ok=True)
    deck = os.path.join(WORKDIR, "calib_probes.pptx")
    preview = os.path.join(WORKDIR, "preview")

    manifest = build_probe_deck(deck)
    print(f"probe deck: {deck} — {len(manifest)} проб")
    pngs = render(deck, preview)
    if len(pngs) < len(manifest):
        sys.stderr.write(
            f"calibrate: отрендерено {len(pngs)} из {len(manifest)} слайдов — рендер не полный\n"
        )
        if not pngs:
            return 1

    # измерения
    per_font: dict[str, list[float]] = {}
    rows = []
    for m in manifest:
        if m["slide"] >= len(pngs):
            continue
        real = dark_bbox_width_in(pngs[m["slide"]])
        raw = fit.raw_width_in(m["text"], m["font"], m["size_pt"])
        if not real or raw <= 0:
            continue
        k = real / raw
        per_font.setdefault(m["font"], []).append(k)
        rows.append((m["font"], m["text"], m["size_pt"], raw, real, k))

    print(f"\n{'font':20} {'text':22} {'pt':>4} {'raw_in':>7} {'real_in':>7} {'K':>6}")
    for f, t, s, raw, real, k in rows:
        print(f"{f:20} {t[:22]:22} {s:4.0f} {raw:7.3f} {real:7.3f} {k:6.3f}")

    fonts_k: dict[str, float] = {}
    print(f"\n{'font':20} {'n':>3} {'K_mean':>7} {'K_min':>7} {'K_max':>7} {'spread':>7}")
    all_k: list[float] = []
    for f, ks in per_font.items():
        mean = sum(ks) / len(ks)
        fonts_k[f] = round(mean, 4)
        all_k += ks
        print(f"{f:20} {len(ks):3} {mean:7.3f} {min(ks):7.3f} {max(ks):7.3f} {max(ks)-min(ks):7.3f}")
    default_k = round(sum(all_k) / len(all_k), 4) if all_k else 1.0

    out = {"_default": default_k, "fonts": fonts_k,
           "_meta": {"dpi": DPI, "n_probes": len(rows), "method": "dark-bbox vs raw_width_in"}}
    print(f"\n_default K = {default_k}")
    if a.dry:
        print("--dry: calibration.json НЕ записан")
    else:
        with open(CALIB_JSON, "w", encoding="utf-8") as fh:
            json.dump(out, fh, indent=2, ensure_ascii=False)
        print(f"written: {CALIB_JSON}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
