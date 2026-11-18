#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""chrome.py — единый резолвер пути к Chrome/Chromium (family:custom-контур).

Одна точка правды на всех потребителей: lib/html2pptx.py (обмер + растеризация
SVG), scripts/preflight.py (диагностика), scripts/html_measure.mjs (получает путь
через переменную окружения CHROME_PATH, которую питон-вызыватель проставляет сам,
а в standalone-запуске читает env и тот же список сам — JS-зеркало в конце файла).

Порядок резолва (фидбек «перевёрстка Байк» п.10, Linux/облако 2026-07-16: в облаке
Chromium живёт не в /usr/bin, а в /opt/pw-browsers — статический список его не
находил, и весь custom-контур падал; расширен 0.21.0):
  1. $CHROME_PATH — явное указание;
  2. $PLAYWRIGHT_BROWSERS_PATH — скан chromium-*/chrome-linux/chrome (и mac/win);
  3. глоб /opt/pw-browsers/** — playwright-браузеры облака;
  4. дефолтные кэши playwright (~/Library/Caches/ms-playwright, ~/.cache/…,
     %LOCALAPPDATA%) — `npx playwright install chromium` кладёт сюда;
  5. статический список (mac/win/linux): только Chrome/Chromium. Форки в
     авто-подбор не входят: живая проба Яндекс.Браузера 2026-07-23 — headless
     виснет даже с изолированным профилем (см. _STATIC_CANDIDATES); явный
     $CHROME_PATH работает с любым бинарём на страх и риск указавшего.
"""
from __future__ import annotations

import glob
import os

# playwright кладёт бинарь под этими именами внутри chromium-<rev>/
_PW_BINARIES = [
    "chrome-linux/chrome",
    "chrome-mac/Chromium.app/Contents/MacOS/Chromium",
    "chrome-win/chrome.exe",
]

_STATIC_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/usr/bin/google-chrome", "/usr/bin/chromium-browser", "/usr/bin/chromium",
    # win-пути Chrome (0.21.0)
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    # Форки (Яндекс.Браузер, Edge) в авто-подбор НЕ входят: живая проба ЯБ
    # 2026-07-23 — его headless ВИСНЕТ даже с изолированным --user-data-dir
    # и --use-mock-keychain (сломан в самой сборке ЯБ); авто-подбор сломанного
    # измерителя хуже честной ошибки «браузера нет». Сознательный путь —
    # CHROME_PATH=<путь>: таймауты конвейера (60–120 с) превратят висяк
    # в честную ошибку.
]

# дефолтные кэши playwright (npx playwright install chromium кладёт сюда)
_PW_DEFAULT_ROOTS = [
    os.path.expanduser("~/Library/Caches/ms-playwright"),   # macOS
    os.path.expanduser("~/.cache/ms-playwright"),           # Linux
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright"),  # Windows
]


def _scan_pw_root(root: str) -> str | None:
    """Найти бинарь chromium внутри playwright-корня (root/chromium-*/…)."""
    if not root or not os.path.isdir(root):
        return None
    for base in sorted(glob.glob(os.path.join(root, "chromium-*")), reverse=True):
        for rel in _PW_BINARIES:
            p = os.path.join(base, rel)
            if os.path.exists(p):
                return p
    # некоторые сборки кладут бинарь прямо в корень (фидбек: /opt/pw-browsers/chromium)
    for rel in _PW_BINARIES + ["chromium", "chrome"]:
        p = os.path.join(root, rel)
        if os.path.exists(p):
            return p
    return None


def find_chrome() -> str | None:
    """Путь к Chrome/Chromium по порядку резолва, либо None."""
    explicit = os.environ.get("CHROME_PATH")
    if explicit and os.path.exists(explicit):
        return explicit
    pw = _scan_pw_root(os.environ.get("PLAYWRIGHT_BROWSERS_PATH", ""))
    if pw:
        return pw
    pw = _scan_pw_root("/opt/pw-browsers")
    if pw:
        return pw
    for root in _PW_DEFAULT_ROOTS:
        pw = _scan_pw_root(root)
        if pw:
            return pw
    return next((p for p in _STATIC_CANDIDATES if os.path.exists(p)), None)


def security_flags() -> list[str]:
    """Chrome flags shared by preview, measurement and SVG rasterization.

    The sandbox stays enabled by default.  Root-only containers may opt into
    Chromium's unsafe mode explicitly; normal desktop runs never need it.
    CSP validation is the primary page-level network gate, while resolver rules
    provide defence in depth against accidental external requests.
    """
    flags = [
        "--disable-background-networking",
        "--disable-component-update",
        "--disable-sync",
        "--metrics-recording-only",
        "--host-resolver-rules=MAP * ~NOTFOUND",
    ]
    if os.environ.get("PRESENTATION_ASSEMBLER_ALLOW_UNSANDBOXED_CHROME") == "1":
        flags.append("--no-sandbox")
    return flags


if __name__ == "__main__":       # быстрый диагностический вывод пути
    print(find_chrome() or "не найден")
