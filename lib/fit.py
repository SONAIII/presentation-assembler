#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fit.py — единый калиброванный измеритель текста для pptx-engine.

ОДИН измеритель и для рендера, и для статического линта (plan §6/§12): если
блок-генератор и линт меряют по-разному, их вердикты «влезает / не влезает»
расходятся. Поэтому здесь — чистая текстовая математика без зависимости от
python-pptx, чтобы и рендерер, и lint, и калибровка импортировали одно и то же.

Ключевые отличия от char-based оценки (как est_lines в gokit):
  * ширина по ГЛИФАМ через Pillow (читает файл шрифта), а не по числу символов —
    иначе «40%» и «193» считаются равными, хотя глиф «%» широкий (plan §12);
  * шрифт резолвится по NAME TABLE (family+style), а не по имени файла —
    «YS SemiMono Regular» лежит в YS_SemiMono_Pixels_June2025.ttf, подстрокой
    по имени файла не находится;
  * калибровочный коэффициент K[font] под реальный рендер PowerPoint
    (scripts/calibrate.py пишет calibration.json) — Pillow по advance-ширинам
    систематически расходится с тем, что PowerPoint рисует.

Перенос — ПО СЛОВАМ, как PowerPoint. Неразрывный пробел (U+00A0) не делит слова.
"""
from __future__ import annotations

import glob
import json
import os
import re
import unicodedata

PT_PER_INCH = 72.0
EMU_PER_INCH = 914_400
NBSP = " "

# Измеряем на крупном EM и линейно масштабируем — advance-ширины скейлятся
# линейно у векторных шрифтов, а ImageFont.truetype(file, px) дорого пересоздавать
# на каждый кегль при shrink-поиске.
_REF_PX = 1000

_FONT_DIRS = [d for d in [
    os.path.expanduser("~/Library/Fonts"),
    "/Library/Fonts",
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    # Windows: пользовательская установка (ПКМ → «Установить») и системная
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    if os.environ.get("LOCALAPPDATA") else "",
    r"C:\Windows\Fonts",
    # Linux: сюда ставит fetch_template --fonts…
    os.path.expanduser("~/.local/share/fonts"),
    # …а руками шрифты кладут и сюда (фидбек H, Linux/облако 2026-07-16);
    # системные каталоги вложенные — скан рекурсивный
    os.path.join(os.environ.get("XDG_DATA_HOME", ""), "fonts")
    if os.environ.get("XDG_DATA_HOME") else "",
    os.path.expanduser("~/.fonts"),
    "/usr/share/fonts",
    "/usr/local/share/fonts",
] if d]

_CALIB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "calibration.json")


def _norm(s: str) -> str:
    """Сравнение имён шрифтов: только буквы/цифры, в нижнем регистре."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# --------------------------------------------------------------- индекс шрифтов

_font_index_cache: dict[str, str] | None = None


def font_index() -> dict[str, str]:
    """norm("Family Style") и norm("Family") -> путь к файлу шрифта.

    Строится один раз по name-table всех TTF/OTF в системных папках.
    «Family Style» имеет приоритет над «Family» (полное имя точнее)."""
    global _font_index_cache
    if _font_index_cache is not None:
        return _font_index_cache
    from PIL import ImageFont

    idx: dict[str, str] = {}
    files: list[str] = []
    for d in _FONT_DIRS:
        # рекурсивно: линуксовые системные каталоги держат шрифты в подпапках
        # (truetype/…); ** матчит и корень — mac/win пути не меняются
        files += glob.glob(os.path.join(d, "**", "*.ttf"), recursive=True)
        files += glob.glob(os.path.join(d, "**", "*.otf"), recursive=True)
    for f in files:
        try:
            family, style = ImageFont.truetype(f, 40).getname()
        except Exception:
            continue
        full = f"{family} {style}".strip()
        # full-name заявляем «жёстко» (перезаписываем), family — только если пусто,
        # чтобы «YS Geo Bold» не затирался первым попавшимся «YS Geo».
        idx[_norm(full)] = f
        idx.setdefault(_norm(family), f)
    _font_index_cache = idx
    return idx


def resolve_font(name: str | None) -> str | None:
    """Имя шрифта из OOXML (run.font.name) -> файл. None, если не нашли."""
    if not name:
        return None
    name = re.sub(r"^[+]m[jn]-lt$", "", name)  # +mj-lt/+mn-lt — тематические ссылки
    if not name:
        return None
    idx = font_index()
    key = _norm(name)
    if key in idx:
        return idx[key]
    # запасной путь: подстрока по имени файла (как старый _font_file)
    for cand in idx.values():
        if key and key in _norm(os.path.basename(cand)):
            return cand
    return None


# ------------------------------------------------------------------ калибровка

_calib_cache: dict | None = None


def load_calibration() -> dict:
    """{ norm(font): K, "_default": K }. K умножает сырую Pillow-ширину,
    приближая её к тому, что реально рисует PowerPoint (см. scripts/calibrate.py)."""
    global _calib_cache
    if _calib_cache is not None:
        return _calib_cache
    data = {"_default": 1.0}
    if os.path.isfile(_CALIB_PATH):
        try:
            with open(_CALIB_PATH, encoding="utf-8") as f:
                raw = json.load(f)
            data["_default"] = float(raw.get("_default", 1.0))
            for k, v in (raw.get("fonts") or {}).items():
                data[_norm(k)] = float(v)
        except Exception:
            pass
    _calib_cache = data
    return data


def calibration_factor(font_name: str | None) -> float:
    c = load_calibration()
    return c.get(_norm(font_name or ""), c["_default"])


# ------------------------------------------------------------ измерение ширины

_imagefont_cache: dict[str, object] = {}


def _ref_font(font_file: str):
    ft = _imagefont_cache.get(font_file)
    if ft is None:
        from PIL import ImageFont

        ft = ImageFont.truetype(font_file, _REF_PX)
        _imagefont_cache[font_file] = ft
    return ft


class FontNotFound(Exception):
    pass


def raw_width_in(text: str, font_name: str | None, size_pt: float) -> float:
    """Ширина строки в дюймах по advance-глифам (БЕЗ калибровки).
    Бросает FontNotFound, если шрифт не нашёлся — молча врать нельзя,
    иначе лишний перенос/overflow проходит мимо линта."""
    ff = resolve_font(font_name)
    if not ff:
        raise FontNotFound(font_name or "<none>")
    if not text:
        return 0.0
    w_ref = _ref_font(ff).getlength(text)  # px при EM=_REF_PX
    return w_ref / _REF_PX * size_pt / PT_PER_INCH


def text_width_in(
    text: str, font_name: str | None, size_pt: float, *, calibrated: bool = True
) -> float:
    """Ширина строки в дюймах. По умолчанию калибрована под PowerPoint."""
    w = raw_width_in(text, font_name, size_pt)
    return w * calibration_factor(font_name) if calibrated else w


# --------------------------------------------------------------- перенос строк


def _split_keep_nbsp(paragraph: str) -> list[str]:
    """Слова для переноса: режем по обычным пробелам/табам, NBSP НЕ делит."""
    return re.split(r"[ \t]+", paragraph.strip())


def wrap_lines(
    text: str,
    font_name: str | None,
    size_pt: float,
    max_width_in: float,
    *,
    calibrated: bool = True,
) -> list[str]:
    """Перенос ПО СЛОВАМ, как PowerPoint. Явные \\n — жёсткие переносы.
    Слово длиннее строки не делим по буквам (PowerPoint не делит без дефиса) —
    оно даёт строку с overflow; ловит лишь lint по ширине, не молчим."""
    out: list[str] = []
    for para in str(text).split("\n"):
        words = [w for w in _split_keep_nbsp(para) if w != ""]
        if not words:
            out.append("")
            continue
        cur = words[0]
        for w in words[1:]:
            cand = f"{cur} {w}"
            if text_width_in(cand, font_name, size_pt, calibrated=calibrated) <= max_width_in:
                cur = cand
            else:
                out.append(cur)
                cur = w
        out.append(cur)
    return out


def line_count(
    text: str, font_name: str | None, size_pt: float, max_width_in: float, *, calibrated: bool = True
) -> int:
    return len(wrap_lines(text, font_name, size_pt, max_width_in, calibrated=calibrated))


def widow_nbsp(
    text: str, font_name: str | None, size_pt: float, max_width_in: float,
    *, calibrated: bool = True,
) -> str:
    """Убрать висячее слово: если естественный перенос оставляет ОДНО слово на
    последней строке — склеить два последних слова неразрывным пробелом, чтобы
    к висяку спустилось соседнее (user-правило 2026-07). Склейку применяем, только
    если она не добавляет строк; иначе возвращаем текст как есть (поймает lint).
    Многострочный вход (явные \\n) и тексты <3 слов не трогаем."""
    # nbsp внутри текста (после коротких предлогов) НЕ мешает: split/wrap_lines
    # работают с ним корректно, а уже склеенную последнюю пару widow не перефлагует
    if not text or "\n" in text or len(text.split()) < 3:
        return text
    try:
        lines = wrap_lines(text, font_name, size_pt, max_width_in, calibrated=calibrated)
    except FontNotFound:
        return text
    if len(lines) < 2 or len(lines[-1].split()) != 1:
        return text
    # склейка = замена ПОСЛЕДНЕГО обычного пробела на nbsp РОВНО в реальном тексте
    # (модель совпадает с мутацией рунов _insert_last_nbsp; иначе существующий
    # nbsp после короткого предлога теряется в модели и склейка врёт).
    idx = text.rstrip().rfind(" ")
    if idx == -1:
        return text
    glued = text[:idx] + NBSP + text[idx + 1:]
    # ни одна неразрывная nbsp-цепочка не шире строки — иначе токен-overflow. Ловит
    # и случай, когда соседний зазор УЖЕ nbsp (после предлога) и склейка сцепляет
    # 3+ слова в один токен (регресс: «Что␣показал␣пилот» 13.9" в боксе 11.0").
    # Неисправимое склейкой остаётся висяком — его флагует lint widow (нужен реворд).
    for tok in glued.split(" "):        # split по ОБЫЧНОМУ пробелу — nbsp-цепочки целы
        if text_width_in(tok, font_name, size_pt, calibrated=calibrated) > max_width_in:
            return text
    if line_count(glued, font_name, size_pt, max_width_in, calibrated=calibrated) <= len(lines):
        return glued
    return text


def text_height_in(
    text: str,
    font_name: str | None,
    size_pt: float,
    max_width_in: float,
    *,
    line_spacing: float = 1.2,
    calibrated: bool = True,
) -> float:
    """Высота блока в дюймах: строки × кегль × межстрочный коэффициент."""
    n = line_count(text, font_name, size_pt, max_width_in, calibrated=calibrated)
    return n * size_pt * line_spacing / PT_PER_INCH


# ------------------------------------------------------------------ shrink/fit


def shrink_to_fit(
    text: str,
    font_name: str | None,
    max_size_pt: float,
    max_width_in: float,
    *,
    wrap: bool = False,
    max_height_in: float | None = None,
    min_size_pt: float = 8.0,
    line_spacing: float = 1.2,
    step_pt: float = 0.5,
    calibrated: bool = True,
) -> float:
    """Наибольший кегль ≤ max_size_pt, при котором текст влезает.

    wrap=False (слот значения/стата, plan §12): меряем одной строкой по ширине,
    перенос запрещён. wrap=True: проверяем число строк по ширине и, если задана
    max_height_in, высоту. Возвращает кегль (может вернуть min_size_pt, если даже
    он не влезает — тогда сработает overflow-lint)."""
    size = max_size_pt
    while size >= min_size_pt:
        if not wrap:
            ok = text_width_in(text, font_name, size, calibrated=calibrated) <= max_width_in
        else:
            ok = True
            if max_height_in is not None:
                ok = (
                    text_height_in(
                        text, font_name, size, max_width_in,
                        line_spacing=line_spacing, calibrated=calibrated,
                    )
                    <= max_height_in
                )
        if ok:
            return round(size, 2)
        size -= step_pt
    return round(min_size_pt, 2)


# ----------------------------------------------------------- пустые строки/чистка


def visible_glyph_run(text: str) -> str:
    """Текст без управляющих/zero-width символов — для проверки «строка пустая»."""
    return "".join(
        ch for ch in (text or "") if not unicodedata.category(ch).startswith("C")
    ).strip()


# Ячейка таблицы «числовая»: цифры + числовая пунктуация/знаки/валюта, БЕЗ слов
# («18 тыс» — текст). Правило вёрстки (user 2026-07-20): числовые колонки — вправо,
# текстовые — влево; заголовок — как его колонка. Используют fill (авто-algn
# datatable) и lint (чекер table-align).
_NUM_CELL_RE = re.compile(
    r"^[\s ]*[+±≈<>~-]?"
    r"[\d\s .,%+±≈<>~/:×–—-]*\d[\d\s .,%+±≈<>~/:×–—-]*"
    r"[₽$€%]?[\s ]*$")
NEUTRAL_CELLS = {"", "—", "–", "-", "н/д", "n/a"}   # прочерки типовость не решают


def is_numeric_cell(text: str) -> bool:
    return bool(_NUM_CELL_RE.match(text or ""))


if __name__ == "__main__":  # быстрый self-test на строках демо Еды
    import sys

    probes = [
        ("40%", "YS Geo Heavy", 138.0),
        ("193", "YS Geo Heavy", 138.0),
        ("32%", "YS Geo Heavy", 88.0),
        ("2×", "YS Geo Heavy", 88.0),
        ("4,8", "YS Geo Heavy", 88.0),
        ("меньше шагов до заказа", "YS Text Regular", 32.0),
    ]
    print(f"calibration: {load_calibration()}")
    print(f"{'text':12} {'font':16} {'pt':>5} {'raw_in':>8} {'calib_in':>9}")
    for t, f, s in probes:
        try:
            r = raw_width_in(t, f, s)
            c = text_width_in(t, f, s)
            print(f"{t!r:12} {f:16} {s:5.0f} {r:8.3f} {c:9.3f}")
        except FontNotFound as e:
            print(f"{t!r:12} {f:16} {s:5.0f}   FONT NOT FOUND: {e}", file=sys.stderr)
