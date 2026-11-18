#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""template_profile.py — пер-шаблонный профиль кастом-контура (family:custom).

ОДИН дом для того, что было захардкожено под Еду: шрифты (имена/файлы/резолвер
CSS-веса), инлайн-жирный fill-пути, лэйауты хрома по фону, калибровка заголовка,
канва, санкция фото-панели A15. Данные — layouts/<tdir>/custom-profile.json;
значения eda = прежним константам html2pptx/fill 1:1 (юнит в tests замораживает
равенство). НЕ brand-tokens.json — у того другой потребитель (generate_tiles).

Ambient tdir: сборка одного дека идёт под ОДНИМ шаблоном. build_deck.main()
ставит set_active(tdir) один раз; глубокий fill-хот-путь (bold-карта в
_set_line_md) читает active() без прокида tdir через десяток сигнатур.
html2pptx-входы (build_custom_slide / pick_layout / measure_html) получают tdir
ЯВНО — контракт Г.4.1. Дефолт — 'eda': тесты и standalone работают как прежде.
"""
from __future__ import annotations

import json
import os

_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CACHE: dict[str, dict] = {}
_ACTIVE = "eda"
DEFAULT_TDIR = "eda"


def profile_path(tdir: str) -> str:
    return os.path.join(_REPO, "layouts", tdir, "custom-profile.json")


def load(tdir: str = DEFAULT_TDIR) -> dict:
    """Профиль по tdir с кэшем. Отсутствие файла — понятная ошибка (пак шаблона
    без custom-profile.json неполон; профиль готовит команда скилла)."""
    if tdir not in _CACHE:
        p = profile_path(tdir)
        if not os.path.exists(p):
            raise FileNotFoundError(
                f"нет профиля кастома {p} — пак шаблона {tdir!r} неполон "
                f"(custom-profile.json готовит команда скилла)")
        with open(p, encoding="utf-8") as f:
            _CACHE[tdir] = json.load(f)
    return _CACHE[tdir]


def set_active(tdir: str) -> None:
    global _ACTIVE
    _ACTIVE = tdir


def active() -> str:
    return _ACTIVE


# --------------------------------------------------------------- шрифты
def font_file_name(tdir: str, style: str) -> str:
    """Имя ttf-файла для стиля PP; фолбэк — файл fallback-стиля (как прежний
    `_FONT_FILES.get(name) or _FONT_FILES["YS Text"]`)."""
    fonts = load(tdir)["fonts"]
    files = fonts["files"]
    return files.get(style) or files[fonts["fallback"]]


def resolve_font(tdir: str, family: str, weight: float) -> tuple[str, bool]:
    """CSS family+weight → (имя стиля PP, synthetic_bold). Первое совпавшее
    правило resolve; эквивалент прежней html2pptx.font_name."""
    for rule in load(tdir)["fonts"]["resolve"]:
        fc = rule.get("family_contains")
        if fc is not None and fc not in family:
            continue
        wmin = rule.get("weight_min")
        if wmin is not None and weight < wmin:
            continue
        wmax = rule.get("weight_max")
        if wmax is not None and weight > wmax:
            continue
        return rule["style"], bool(rule.get("synthetic_bold", False))
    # хвост-фолбэк (у Еды недостижим — последнее правило безусловно)
    return load(tdir)["fonts"]["fallback"], False


def bold_family(tdir: str) -> dict:
    """Инлайн **жирный** fill-пути: имя стиля → его bold-файл-семейство."""
    return load(tdir)["fonts"]["bold_family"]


# --------------------------------------------------------------- лэйауты хрома
def _in(v, rng) -> bool:
    return True if not rng else rng[0] <= v <= rng[1]


def _bg_rule(lay: dict, bg):
    if not bg:
        return None
    r, g, b = bg[:3]
    for rule in lay["by_bg"]:
        if _in(r, rule.get("r")) and _in(g, rule.get("g")) and _in(b, rule.get("b")):
            return rule
    return None


def layout_for_bg(tdir: str, bg) -> str:
    """Фон HTML-слайда → имя лэйаута хрома (default, если ни одно правило by_bg
    не совпало). Эквивалент прежних html2pptx.LAYOUT_BY_BG / LAYOUT_DEFAULT."""
    lay = load(tdir)["layouts"]
    rule = _bg_rule(lay, bg)
    return rule["layout"] if rule else lay["default"]


def bg_rule_matched(tdir: str, bg) -> bool:
    """Фон попал в правило by_bg (хром лэйаута рассчитан на него)."""
    return _bg_rule(load(tdir)["layouts"], bg) is not None


def default_bg(tdir: str):
    """Фон, на который рассчитан default-лэйаут хрома (RGB), или None — ключа
    в профиле нет: движок не судит «вне палитры» и фон слайда не красит."""
    return load(tdir)["layouts"].get("default_bg")


# --------------------------------------------------------------- прочие числа
def title_dy_px(tdir: str) -> float:
    return load(tdir)["calibration"]["title_dy_px"]


def canvas_px(tdir: str) -> tuple[int, int]:
    w, h = load(tdir)["canvas"]["px"]
    return w, h


def a15_photo_panel(tdir: str) -> tuple[float, float]:
    """(top_max, bottom_min_from_h) — санкция фото-панели канона в гварде A15."""
    pp = load(tdir)["a15"]["photo_panel"]
    return pp["top_max"], pp["bottom_min_from_h"]


def margins(tdir: str):
    """Контентные поля канвы custom-слайда (px: left/right/top/bottom) или None —
    ключа в профиле нет, движок сторожит только канву и якоря низа (как раньше)."""
    return load(tdir).get("margins")
