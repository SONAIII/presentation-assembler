#!/usr/bin/env python3
"""Скриншот одного слайда custom-HTML — такт-1 превью пользователю
(SKILL.md шаг 3.7): макет кастом-слайда показывается ДО сборки дека,
правка на этом такте стоит один HTML-файл, а не пересборку и переприёмку
всего дека.

Использование:
    python3 scripts/html_preview.py slide.html [-s N] [-o out.png]

N — индекс слайда в файле (0-based, тот же, что html_slide в deckspec).
Рендерит канву 1920×1080 тем же headless Chrome, что и обмерщик конвейера, —
превью совпадает с тем, что увидит конвертер. Скриншот НЕ заменяет рендер
собранного .pptx (STOP шага 4): финальную картинку даёт только сборка.
"""
import argparse
import os
import subprocess
import sys
import tempfile

# Единый резолвер пути к Chrome (lib/chrome.py) — одна точка правды на весь
# custom-контур: обмер (scripts/html_measure.mjs), конверсия (lib/html2pptx.py)
# и это превью берут браузер одинаково ($CHROME_PATH → $PLAYWRIGHT_BROWSERS_PATH →
# /opt/pw-browsers → штатные mac/win/linux). Раньше здесь жил свой статический
# список — дрейфовал от резолвера и был слеп к env/playwright/облаку.
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "lib"))
import chrome as _chrome                          # noqa: E402
from html_security import inject_csp, validate_html_file  # noqa: E402


def chrome_path() -> str:
    p = _chrome.find_chrome()
    if p:
        return p
    raise SystemExit("Chrome/Chromium не найден — нужен для family:custom "
                     "(обмер, конверсия и это превью); см. preflight.py")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("html", help="HTML-файл с кастом-слайдом (канва 1920×1080)")
    ap.add_argument("-s", "--slide", type=int, default=0,
                    help="индекс слайда в файле (0-based, = html_slide)")
    ap.add_argument("-o", "--out", default=None,
                    help="куда положить PNG (дефолт: <html>_s<N>.png рядом)")
    a = ap.parse_args()

    src = os.path.abspath(a.html)
    if not os.path.exists(src):
        raise SystemExit(f"нет файла {src}")
    out = os.path.abspath(a.out or os.path.splitext(src)[0] + f"_s{a.slide}.png")

    try:
        html = validate_html_file(src)
    except ValueError as exc:
        raise SystemExit(f"небезопасный custom HTML: {exc}") from exc

    # Показать только слайд N — тот же приём, что у обмерщика: слайды —
    # секции-сиблинги внутри #deck (контракт custom-html.md). !important
    # перекрывает и вьюер .deck--single.
    iso = ("<style>#deck .slide{display:none!important}"
           f"#deck .slide:nth-of-type({a.slide + 1})"
           "{display:flex!important}</style>")
    html = html.replace("</body>", iso + "</body>") if "</body>" in html \
        else html + iso
    html = inject_csp(html)

    # Копия во временный файл (как у обмерщика) — поэтому пути к картинкам
    # в custom-HTML обязаны быть абсолютными file:// (custom-html.md).
    tmp = tempfile.NamedTemporaryFile(
        "w", suffix=".html", delete=False, encoding="utf-8")
    tmp.write(html)
    tmp.close()
    # Флаги — эмпирика живых проб 2026-07-23/24 (0.21.0): СВОЙ --user-data-dir
    # в headless=new НЕ давать — с ним виснут и скриншот, и dump-dom с
    # virtual-time (Chrome сам делает эфемерный профиль, он стабилен);
    # mock-keychain/basic гасят невидимый Keychain-диалог macOS на свежем
    # профиле; disable-background-networking и соседи глушат GoogleUpdater-возню
    # между запусками. Зеркало — html_measure.mjs и html2pptx.py.
    try:
        subprocess.run(
            [chrome_path(), "--headless=new", "--disable-gpu",
             "--no-first-run", "--no-default-browser-check",
             "--use-mock-keychain", "--password-store=basic",
             *_chrome.security_flags(),
             "--hide-scrollbars", "--force-device-scale-factor=1",
             "--window-size=1920,1080", "--screenshot=" + out,
             "file://" + tmp.name],
            check=True, capture_output=True, timeout=60)
    finally:
        os.unlink(tmp.name)
    print(out)


if __name__ == "__main__":
    main()
