#!/usr/bin/env python3
"""Скрин приложения → мокап телефона в бренд-стиле (фидбек E медиакита,
2026-07-16): ТОНКАЯ ink-рамка edge-to-edge, скругление, прозрачный фон.
Толстая тёмная обводка выбивается из бренда — deck.css построен на мягких
скруглённых карточках без тяжёлых бордеров; тень по умолчанию ВЫКЛЮЧЕНА
(в шаблоне Еды теней нет), --shadow — для внешних употреблений.

    python3 scripts/phone_mockup.py screen.png [ещё.png …] [-o outdir]
        [--bezel N] [--radius N] [--shadow]

Выход: <имя>_phone.png (RGBA) рядом с исходником или в outdir. Дефолты
масштабируются от ширины скрина (bezel ≈ 1.2%, radius ≈ 10%).

Куда ставить: PNG-мокап — в фото-слоты (images/three_pics и др.) через
"photos" deckspec. Для КАСТОМ-слайдов телефон лучше собирать нативно —
рецепт «витрина скринов» в custom-recipes-eda.md: рамка там остаётся
редактируемым roundRect, а не запечённой картинкой.
"""
from __future__ import annotations

import argparse
import os

from PIL import Image, ImageDraw, ImageFilter

_INK_FALLBACK = (35, 15, 10, 255)   # ink Еды (#230F0A) — фолбэк, если нет brand-tokens


def _ink(tdir: str) -> tuple[int, int, int, int]:
    """Цвет рамки — brand-tokens colors.text шаблона (#230F0A у Еды)."""
    import json
    bt = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      "layouts", tdir, "brand-tokens.json")
    try:
        with open(bt, encoding="utf-8") as f:
            h = json.load(f)["colors"]["text"]["hex"]
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 255)
    except (OSError, ValueError, KeyError):
        return _INK_FALLBACK


def mockup(src: str, out: str, bezel: int | None, radius: int | None,
           shadow: bool, ink: tuple = _INK_FALLBACK) -> str:
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    b = bezel if bezel else max(8, round(w * 0.012))
    r = radius if radius else round(w * 0.10)

    # скруглённый скрин
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], r, fill=255)
    im.putalpha(mask)

    # рамка-устройство: roundRect ink с радиусом r+b, скрин внутри
    fw, fh = w + 2 * b, h + 2 * b
    frame = Image.new("RGBA", (fw, fh), (0, 0, 0, 0))
    ImageDraw.Draw(frame).rounded_rectangle([0, 0, fw - 1, fh - 1], r + b,
                                            fill=ink)
    frame.alpha_composite(im, (b, b))

    if shadow:
        m = round(fw * 0.06)
        canvas = Image.new("RGBA", (fw + 2 * m, fh + 2 * m), (0, 0, 0, 0))
        sh = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        ImageDraw.Draw(sh).rounded_rectangle(
            [m, m + m // 2, m + fw - 1, m + fh - 1 + m // 2], r + b,
            fill=ink[:3] + (70,))
        canvas.alpha_composite(sh.filter(ImageFilter.GaussianBlur(m // 2)))
        canvas.alpha_composite(frame, (m, m))
        frame = canvas

    frame.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("screens", nargs="+", help="PNG/JPG скрины приложения")
    ap.add_argument("-o", "--outdir", default=None)
    ap.add_argument("--bezel", type=int, default=None,
                    help="толщина рамки, px (дефолт ≈1.2%% ширины)")
    ap.add_argument("--radius", type=int, default=None,
                    help="радиус скругления скрина, px (дефолт ≈10%% ширины)")
    ap.add_argument("--shadow", action="store_true",
                    help="мягкая тень (по умолчанию выключена — бренд без теней)")
    ap.add_argument("--template", default="eda",
                    help="tdir шаблона — цвет рамки из его brand-tokens (дефолт eda)")
    a = ap.parse_args()
    ink = _ink(a.template)
    for src in a.screens:
        stem = os.path.splitext(os.path.basename(src))[0]
        outdir = a.outdir or os.path.dirname(os.path.abspath(src))
        os.makedirs(outdir, exist_ok=True)
        out = os.path.join(outdir, f"{stem}_phone.png")
        print(mockup(src, out, a.bezel, a.radius, a.shadow, ink))


if __name__ == "__main__":
    main()
