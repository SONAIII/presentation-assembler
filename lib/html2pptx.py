#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""html2pptx.py — кастомные слайды через HTML-хэтч (ветка роутера fill → structure → custom).

Слайд, свёрстанный по дизайн-системе Еды скилла (assets/eda/deck.css —
завендорена, сиблинг НЕ нужен), обмеряется headless Chrome'ом
(scripts/html_measure.mjs) и собирается нативными
шейпами НА ХРОМЕ дизайнерского шаблона: заголовок → TITLE-плейсхолдер лэйаута,
номер страницы → от лэйаута, фон HTML выбирает лэйаут.

Канва 1920×1080 = слайд шаблона (1px = 1pt, датумы совпадают).

Ключевые факты (выведены спайком html2pptx 2026-07 с пиксельными сверками; архив спайка — у владельца скилла):
- PowerPoint НЕ применяет GPOS-кернинг (читает только legacy-kern; у YS-шрифтов
  её нет) → метрики PP = advance-ширины HarfBuzz БЕЗ кернинга (pp_width).
- Инлайн-плашки .mark несут горизонтальные паддинги в потоке строки — в тексте
  плейсхолдера они восстанавливаются межзнаковым интервалом spc.
- Плейсхолдер сужается до ширины h1 из обмера — иначе PP переносит строки не там,
  где браузер, а тело слайда разложено вокруг браузерных строк.
- a:gradFill вставлять строго после геометрии в spPr — иначе PP молча берёт
  стилевую заливку темы.

Пер-шаблонное (шрифты, лэйауты хрома по фону, калибровка заголовка, канва,
санкция фото-панели A15) вынесено в layouts/<tdir>/custom-profile.json и читается
через lib/template_profile.py — этот модуль от Еды не зависит, tdir прокидывается
входами (build_custom_slide / pick_layout / measure_html), дефолт 'eda'.

Калибровка (66pt / lnSpc 80%, замерена по PP-рендерам спайка 2026-07-10;
при смене шаблона или кегля заголовка — перемерить, см. README спайка):
  calibration.title_dy_px = -9 в custom-profile.json — PP ставит чернила
  заголовка выше браузера (модель строчного бокса).
"""
from __future__ import annotations

import glob
import os
import re
import subprocess
import tempfile
import urllib.parse

from lxml import etree
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

from images import cover_fit_crop, _set_srcrect
import chrome as CHROME
from html_security import validate_html_file
import template_profile as TP                    # пер-шаблонный профиль кастома

PX = 12700                      # EMU в 1px канвы (1px = 1pt); универсально, не пер-шаблонно
# Шрифты (имена/файлы/резолвер), лэйауты хрома, калибровка, канва, санкция A15 —
# в custom-profile.json (lib/template_profile.py). Ниже — только OS-уровневый
# поиск ttf по системным папкам (не пер-шаблонный).
_FONT_DIRS = [d for d in [
    os.path.expanduser("~/Library/Fonts"), "/Library/Fonts",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    if os.environ.get("LOCALAPPDATA") else "",
    r"C:\Windows\Fonts",
    os.path.expanduser("~/.local/share/fonts"),
    # Linux руками/системно (фидбек H 2026-07-16); подпапки — рекурсивный фолбэк
    os.path.join(os.environ.get("XDG_DATA_HOME", ""), "fonts")
    if os.environ.get("XDG_DATA_HOME") else "",
    os.path.expanduser("~/.fonts"),
    "/usr/share/fonts", "/usr/local/share/fonts",
] if d]

class Html2PptxError(RuntimeError):
    pass


def _validate_canvas(prs, tdir: str) -> None:
    """Слайд шаблона = канва дизайн-системы (1px=1pt). Грубое расхождение —
    честная ошибка, не молчаливый сдвиг/обрез тела custom-слайда. Допуск ~2px:
    датумы Еды дают слайд 1919.875px по ширине (0.125px слоп, не дефект)."""
    cw, ch = TP.canvas_px(tdir)
    tol = 2 * PX
    if abs(prs.slide_width - cw * PX) > tol or abs(prs.slide_height - ch * PX) > tol:
        raise Html2PptxError(
            f"слайд шаблона {prs.slide_width/PX:.1f}×{prs.slide_height/PX:.1f}pt ≠ "
            f"канва профиля {cw}×{ch}px — тело custom-слайда сдвинется/обрежется; "
            f"перемерь канву и custom-profile.json.canvas.px "
            f"(custom-profile шаблона готовит команда скилла)")


def _validate_measured_canvas(data: dict, tdir: str) -> None:
    """Обмер .slide = канва шаблона. Мис-размер .slide (автор сверстал не на
    канву дизайн-системы) разложил бы тело мимо якорей — ловим на обмере."""
    cw, ch = TP.canvas_px(tdir)
    sw, sh = data["slide"]["w"], data["slide"]["h"]
    if abs(sw - cw) > 2 or abs(sh - ch) > 2:
        raise Html2PptxError(
            f"канва HTML-слайда {sw:.0f}×{sh:.0f}px ≠ канва шаблона {cw}×{ch}px "
            f"(custom-profile.json.canvas.px) — .slide должен быть размером канвы "
            f"дизайн-системы (references/custom-html.md)")


def chrome_path() -> str:
    import chrome as _chrome                    # единый резолвер (env + pw-browsers)
    p = _chrome.find_chrome()
    if p:
        return p
    raise Html2PptxError(
        "Chrome/Chromium не найден — он нужен family:custom как измеритель "
        "(обмер HTML и растеризация SVG); остальной сборке не нужен. Три выхода: "
        "(1) поставь headless-хромиум одной командой: npx playwright install "
        "chromium (~150 МБ, резолвер найдёт его автоматически); (2) укажи путь "
        "явно: CHROME_PATH='/Applications/Google Chrome.app/Contents/MacOS/"
        "Google Chrome' (Яндекс.Браузер НЕ годится — его headless виснет, "
        "проверено 2026-07-23); (3) реворди слайд под готовое семейство "
        "(references/families-eda.md) — custom не понадобится")


# ---------------------------------------------------------------- PP-метрики
_HB = {}


def _font_file(name: str, tdir: str = TP.DEFAULT_TDIR) -> str:
    fname = TP.font_file_name(tdir, name)
    for d in _FONT_DIRS:
        p = os.path.join(d, fname)
        if os.path.exists(p):
            return p
    for d in _FONT_DIRS:                       # линуксовые подпапки (truetype/…)
        hits = glob.glob(os.path.join(d, "**", fname), recursive=True)
        if hits:
            return hits[0]
    raise Html2PptxError(f"шрифт {name!r} ({fname}) не установлен — "
                         "python3 scripts/fetch_template.py --fonts")


def pp_width(text: str, font: str, size_pt: float, tdir: str = TP.DEFAULT_TDIR) -> float:
    """Ширина текста, как её наберёт PowerPoint: advance-ширины БЕЗ кернинга
    (PP игнорирует GPOS-кернинг; браузер кернит — его обмер тут не годится)."""
    try:
        import uharfbuzz as hb
    except ImportError as e:
        raise Html2PptxError("нет uharfbuzz — python3 -m pip install uharfbuzz "
                             "(нужен для family:custom)") from e
    key = (tdir, font)                          # стиль → файл пер-шаблонный
    if key not in _HB:
        face = hb.Face(hb.Blob.from_file_path(_font_file(font, tdir)))
        _HB[key] = (hb.Font(face), face.upem)
    f, upm = _HB[key]
    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    hb.shape(f, buf, {"kern": False})
    return sum(pos.x_advance for pos in buf.glyph_positions) / upm * size_pt


def font_name(family: str, weight: float, tdir: str = TP.DEFAULT_TDIR) -> tuple[str, bool]:
    """CSS family+weight → (имя шрифта PP, synthetic_bold). Контракт и данные —
    custom-profile.json (fonts.resolve). Канон бренда (user-фидбек 2026-07-16):
    «жирный» = ФАЙЛОВОЕ семейство стиля (typeface «YS Text Bold»), а не синтетика
    b=1 поверх Regular — пишем полные имена стилей, как их пишет PowerPoint
    дизайнера (synthetic_bold=false у всех правил Еды)."""
    return TP.resolve_font(tdir, family, weight)


def rgb(v, opacity=1.0, bg=(255, 255, 255)) -> RGBColor:
    r, g, b = v[0], v[1], v[2]
    a = (v[3] if len(v) > 3 else 1.0) * opacity
    if a < 1.0:
        r = r*a + bg[0]*(1-a); g = g*a + bg[1]*(1-a); b = b*a + bg[2]*(1-a)
    return RGBColor(int(r), int(g), int(b))


def pick_layout(prs, bg, tdir: str = TP.DEFAULT_TDIR):
    name = TP.layout_for_bg(tdir, bg)
    for l in prs.slide_master.slide_layouts:
        if l.name == name:
            return l
    raise Html2PptxError(f"в шаблоне нет лэйаута {name!r} для хрома custom-слайда")


def _layout_canvas_cover(layout, tdir=TP.DEFAULT_TDIR):
    """Не-плейсхолдерный шейп лэйаута, накрывающий ≥95% канвы (дизайнерский
    фоновый прямоугольник — у Еды «Rectangle 3»), или None. Такой шейп рисуется
    ПОВЕРХ слайдового p:bg — честный фон слайда через него не виден (проба
    PP-рендером 2026-08-10)."""
    cw, ch = TP.canvas_px(tdir)
    for sh in layout.shapes:
        if sh.is_placeholder or sh.width is None:
            continue
        if sh.width >= cw * PX * 0.95 and sh.height >= ch * PX * 0.95:
            return sh
    return None


def _apply_offpalette_bg(slide, slide_bg, layout, warnings, tdir=TP.DEFAULT_TDIR):
    """Фон HTML-секции вне палитры хрома (ни одно правило by_bg) → нативный фон
    слайда. Раньше такой фон молча терялся — слайд собирался на фоне
    default-лэйаута (фидбек переноса HTML 2026-08-10: «тёмный фон не перенёсся,
    пришлось городить фоновую фигуру»). Два механизма: лэйаут без фоновой
    фигуры → p:bg override (лежит под хромом: номер страницы виден, дизайнер
    правит через «Формат фона»); лэйаут с фул-канва фигурой (Еда) → фигура-фон
    ПЕРВЫМ шейпом слайда (под контентом, но поверх лэйаута — его номер/футер
    уходят под фон, как под фото-панелью канона). Хром (цвет заголовка, номера)
    на такой фон НЕ рассчитан — предупреждаем, правда — PP-рендер."""
    if TP.bg_rule_matched(tdir, slide_bg):
        return
    dflt = TP.default_bg(tdir)
    if not dflt or max(abs(int(a) - int(b))
                       for a, b in zip(slide_bg[:3], dflt[:3])) <= 12:
        return                      # фон ≈ фон default-лэйаута — красить нечего
    tail = (f"; хром лэйаута {layout.name!r} на него не рассчитан: сверь "
            f"читаемость заголовка и номера на рендере (канонный путь — "
            f"фоны палитры, references/custom-html.md)")
    if _layout_canvas_cover(layout, tdir) is None:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(slide_bg)
        warnings.append(
            f"фон {tuple(int(c) for c in slide_bg[:3])} вне палитры хрома — "
            f"перенесён нативным фоном слайда{tail}")
    else:
        cw, ch = TP.canvas_px(tdir)
        sp = _add_rect(slide, 0, 0, cw, ch)
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(slide_bg)
        warnings.append(
            f"фон {tuple(int(c) for c in slide_bg[:3])} вне палитры хрома — "
            f"закрашен фигурой-фоном (лэйаут несёт свой фоновый прямоугольник, "
            f"p:bg через него не виден); номер страницы/футер лэйаута уходят "
            f"под фон{tail}")


# ---------------------------------------------------------------- сборка
def _role(item) -> str:
    """data-ppt-role элемента из обмера ('' — атрибута нет). Явная семантика
    ролей при переносе готового HTML (фидбек 2026-08-10): грамматика классов
    дизайн-системы (.h1) остаётся основной, роль — переносимый override."""
    return (item.get("role") or "").strip().lower()


def is_title_text(t: dict) -> bool:
    """Текст — заголовок слайда: data-ppt-role="title" ИЛИ класс .h1
    (грамматика дизайн-системы). Такой текст уходит в TITLE-плейсхолдер
    лэйаута, телом не рисуется."""
    return _role(t) == "title" or t["cls"].startswith("h1")


# A16: слайд без заголовка, но с крупным буквенным текстом у верха канвы —
# почти наверняка заголовок, не помеченный .h1/role (перенос чужого HTML).
# Пороги: >.sub (44px) по кеглю; ≥8 букв отсекает крупные значения («22,400»,
# «84%»); верхняя четверть канвы — датум заголовка (y=79 у Еды).
_TITLE_SUSPECT_MIN_PX = 48
_TITLE_SUSPECT_MIN_LETTERS = 8
_TITLE_SUSPECT_Y_FRAC = 0.25


def custom_content_violations(data: dict, tdir: str = TP.DEFAULT_TDIR) -> list[str]:
    """Содержательные гварды custom-слайда по обмеру (фидбек переноса
    HTML→PPTX 2026-08-10, дек «Еды»): построчные нарушения или пустой список.
    A16 — заголовок мимо TITLE-плейсхолдера: дизайнер не сможет править его
    штатным полем PowerPoint. A17 — график картинкой: SVG с числовыми
    подписями растеризуется, «Изменить данные» у такого графика нет, а продукт
    скилла — редактируемый .pptx (chart-selection.md §5)."""
    out = []
    texts = data.get("texts") or []
    if not any(is_title_text(t) for t in texts):
        H = data["slide"]["h"]
        for t in texts:
            if _role(t) in ("body", "eyebrow", "decor") \
                    or t["cls"].startswith("pagenum"):
                continue
            runs = [r for r in t.get("runs") or [] if not r.get("br")]
            size = max((r.get("sizePx") or 0 for r in runs), default=0)
            letters = sum(ch.isalpha() for r in runs for ch in r.get("text", ""))
            if size >= _TITLE_SUSPECT_MIN_PX \
                    and letters >= _TITLE_SUSPECT_MIN_LETTERS \
                    and t["bbox"]["y"] < H * _TITLE_SUSPECT_Y_FRAC:
                head = " ".join(r.get("text", "") for r in runs).strip()[:40]
                out.append(
                    f"A16: текст «{head}» ({size:.0f}px, y={t['bbox']['y']:.0f}) "
                    f"похож на заголовок, но не помечен — в TITLE-плейсхолдер "
                    f"не попадёт, дизайнер не отредактирует его штатным полем. "
                    f"Пометь .h1 или data-ppt-role=\"title\"; это НЕ заголовок → "
                    f"data-ppt-role=\"body\" (references/custom-html.md)")
    for s in data.get("svgs") or []:
        if _role(s) == "decor":
            continue
        labels = [txt.strip() for txt in
                  re.findall(r"<(?:text|tspan)[^>]*>([^<>]+)<", s.get("svg") or "")
                  if any(ch.isdigit() for ch in txt)]
        if len(labels) >= 3:
            cls = (s.get("cls") or "").split()[0] if s.get("cls") else "?"
            out.append(
                f"A17: SVG .{cls} несёт {len(labels)} числовых подписей "
                f"({', '.join(labels[:3])}…) — это график картинкой: он "
                f"растеризуется и НЕ редактируется через «Изменить данные». "
                f"Числовые ряды — только нативным чартом: chart-донор семейства "
                f"(chart-selection.md §2), нативный data-chart в custom "
                f"(custom-html.md) или таблица datatable; осознанный "
                f"декор без данных — data-ppt-role=\"decor\" на <svg>")
    return out


def font_fidelity_violation(data: dict, tdir: str = TP.DEFAULT_TDIR):
    """Шрифтовая проба обмера (user-фидбек 2026-08-07, дек drive-long: цифры
    слиплись, пилюли наехали на заголовки). Браузер, не нашедший шрифта бренда,
    молча меряет ПОДМЕННЫМ — getComputedStyle возвращает запрошенное имя, не
    фактическое, поэтому по стилям подмену не видно; вся геометрия занижена на
    15–30%, и в PP (который рисует честным шрифтом) текст рвёт плашки. Проба:
    браузерная ширина однострочных ранов vs PP-метрика тем же семейством
    (pp_width); медиана PP/браузер по семейству вне 0.90–1.06 (кернинг браузера
    даёт лишь ~1–2%) → строка-вердикт с именем семейства; None — проба пройдена.

    Ужесточение (хэндофф Ultima 2026-08-11, O6.2): на Windows имена стилей
    падали в serif, а проба МОЛЧАЛА — заголовочные семейства набирали <2
    однострочных образцов на слайд и не судились вовсе. Теперь (1) образцы
    берутся из ЛЮБОГО текста: ран, не перенёсшийся на вторую строку (высота
    его rect ≈ одна строка), меряем и внутри многострочных блоков; (2)
    семейство с единственным образцом судим по широкой полосе 0.85–1.15 —
    подмена на serif даёт 1.2–1.4 и ловится, шум одного рана — нет."""
    import statistics
    by_fam: dict[str, list[float]] = {}
    for t in data.get("texts") or []:
        for r in t.get("runs") or []:
            rect = r.get("rect") or {}
            txt = (r.get("text") or "").strip()
            if r.get("br") or len(txt) < 4 or rect.get("w", 0) < 30:
                continue
            if rect.get("h", 0) > (r.get("sizePx") or 0) * 1.6:
                continue    # ран перенёсся: union-rect не отражает ширину текста
            try:
                fname, _ = font_name(r["family"], r["weight"], tdir)
                wpp = pp_width(txt, fname, r["sizePx"], tdir)
            except Html2PptxError:
                raise
            except Exception:
                continue
            if wpp > 0:
                by_fam.setdefault(fname, []).append(wpp / rect["w"])
    bad = []
    for fam, ratios in by_fam.items():
        med = statistics.median(ratios)
        lo, hi = (0.90, 1.06) if len(ratios) >= 2 else (0.85, 1.15)
        if not (lo <= med <= hi):
            bad.append((fam, med))
    if not bad:
        return None
    fam, med = max(bad, key=lambda x: abs(x[1] - 1))
    return (f"шрифтовая проба: браузер мерил {fam!r} подменным шрифтом "
            f"(PP-ширина/браузерная = {med:.2f}, норма 0.90–1.06) — геометрия "
            f"вёрстки недостоверна, в PP текст порвёт плашки и наедет на соседей. "
            f"Шрифты бренда не видны браузеру: установи их "
            f"(python3 scripts/fetch_template.py --fonts) и пересобери; "
            f"family:custom без честной пробы не собирается")


def build_custom_slide(prs, data: dict, tdir: str = TP.DEFAULT_TDIR) -> tuple[int, dict]:
    """Добавить в prs слайд с хромом лэйаута и телом по обмеру data.
    Возвращает (индекс слайда, отчёт). tdir → профиль кастома (шрифты/лэйауты/
    калибровка/канва)."""
    warnings: list[str] = []
    _validate_canvas(prs, tdir)
    fidelity = font_fidelity_violation(data, tdir)
    if fidelity:
        raise Html2PptxError(fidelity)
    slide_bg = data["slide"]["bg"] or [255, 255, 255]
    layout = pick_layout(prs, slide_bg, tdir)
    slide = prs.slides.add_slide(layout)
    ci = len(prs.slides) - 1
    _apply_offpalette_bg(slide, slide_bg, layout, warnings, tdir)

    title_ph = slide.shapes.title
    titles = [t for t in data["texts"] if is_title_text(t)]
    if len(titles) > 1:
        warnings.append(f"{len(titles)} заголовков (.h1/role=title) — в TITLE-"
                        f"плейсхолдер уходит первый, остальные лягут телом")
    h1 = titles[0] if titles else None
    if h1 and title_ph is not None:
        _fill_title(title_ph, h1, data, warnings, tdir)
    elif h1 is None:
        warnings.append("в HTML нет .h1 — слайд без заголовка (плейсхолдер с заглушкой)")

    painted: list = []

    def bg_under(rect):
        cx, cy = rect["x"] + rect["w"]/2, rect["y"] + rect["h"]/2
        for bb, col in reversed(painted):
            if bb["x"] <= cx <= bb["x"]+bb["w"] and bb["y"] <= cy <= bb["y"]+bb["h"]:
                return col
        return tuple(slide_bg[:3])

    items = ([("box", b) for b in data["boxes"]]
             + [("img", i) for i in data.get("images", [])]
             + [("svg", s) for s in data.get("svgs", [])]
             + [("tbl", t) for t in data.get("tables", [])]
             + [("line", l) for l in data.get("lines", [])]
             + [("chart", c) for c in data.get("charts", [])])
    items.sort(key=lambda t: t[1].get("seq", 0))
    for kind, it in items:
        if kind == "box":
            _draw_box(slide, it, slide_bg, painted)
        elif kind == "img":
            _draw_image(slide, it, warnings)
        elif kind == "tbl":
            _draw_table(slide, it, slide_bg, warnings, tdir)
        elif kind == "line":
            _draw_data_line(slide, it, warnings)
        elif kind == "chart":
            _draw_chart(slide, it, warnings, tdir)
        else:
            _draw_svg(slide, it)

    # плашки/фото нарисованы ПОСЛЕ плейсхолдера → он ушёл бы под них
    if title_ph is not None:
        sp_tree = title_ph._element.getparent()
        sp_tree.remove(title_ph._element)
        sp_tree.append(title_ph._element)

    placed_title = h1 if (h1 and title_ph is not None) else None
    n_texts = _draw_texts(slide, data, bg_under, tdir, skip=placed_title)

    report = {"action": "CUSTOM", "layout": layout.name,
              "atoms": {"boxes": len(data["boxes"]), "texts": n_texts,
                        "images": len(data.get("images", [])),
                        "svgs": len(data.get("svgs", [])),
                        "tables": len(data.get("tables", [])),
                        "lines": len(data.get("lines", [])),
                        "charts": len(data.get("charts", []))},
              "warnings": warnings, "src": None}
    return ci, report


# ---------------------------------------------------------------- заголовок
def _fill_title(title_ph, h1, data, warnings, tdir=TP.DEFAULT_TDIR):
    title_dy = TP.title_dy_px(tdir)
    _strip_paragraph_periods(h1["runs"])   # заголовки без финальной точки (как fill)
    tf = title_ph.text_frame
    tf.clear()
    # отступы в ноль: текст стартует с датума, как h1 в браузере
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    # плейсхолдер шире h1 грамматики → PP переносил бы строки не там, где браузер,
    # а тело слайда разложено вокруг браузерных строк; xfrm пишем целиком
    l, t, hgt = title_ph.left, title_ph.top, title_ph.height
    title_ph.left, title_ph.top, title_ph.height = l, t, hgt
    title_ph.width = Emu(int(h1["bbox"]["w"] * PX))
    p = tf.paragraphs[0]

    if h1["nLines"] > 1 and any(_mark_for(r, data) for r in h1["runs"] if not r.get("br")):
        warnings.append("маркированные плашки в многострочном заголовке: курсор "
                        "PP-метрик не учитывает переносы — проверь глазом, "
                        "плашки не на первой строке могут уехать")

    marks_hit = []
    segs = []                       # [текст, spc_после_последнего_символа_pt]
    cursor = title_ph.left / PX
    for r in [x for x in h1["runs"] if not x.get("br")]:
        fname, _ = font_name(r["family"], r["weight"], tdir)
        m = _mark_for(r, data)
        if not m:
            segs.append([r["text"], 0])
            cursor += pp_width(r["text"], fname, r["sizePx"], tdir)
            continue
        rect, bb = r["rect"], m["bbox"]
        padL = rect["x"] - bb["x"]
        padR = (bb["x"] + bb["w"]) - (rect["x"] + rect["w"])
        if segs:
            pt_, ps = segs[-1]
            if len(pt_) > 1:        # отделить последний символ под spc
                segs[-1] = [pt_[:-1], 0]
                segs.append([pt_[-1], ps + padL])
            else:
                segs[-1][1] = ps + padL
        else:
            warnings.append("плашка .mark в самом начале заголовка: spc вешать "
                            "некуда, паддинг слева потерян")
        w_txt = pp_width(r["text"], fname, r["sizePx"], tdir)
        bb["x"] = cursor            # плашка — от PP-курсора, не от браузера
        bb["w"] = w_txt + padL + padR
        bb["y"] += title_dy
        marks_hit.append(m)
        cursor += padL + w_txt + padR
        if len(r["text"]) > 1:
            segs.append([r["text"][:-1], 0])
            segs.append([r["text"][-1], padR])
        else:
            segs.append([r["text"], padR])

    for text, spc in segs:
        run = p.add_run()
        run.text = text
        if spc:
            run._r.get_or_add_rPr().set("spc", str(int(round(spc * 100))))
        # гарнитуру/кегль НЕ трогаем — их даёт дизайнерский плейсхолдер


def _mark_for(r, data):
    rect = r.get("rect")
    if not rect:
        return None
    cx, cy = rect["x"] + rect["w"]/2, rect["y"] + rect["h"]/2
    for m in data["boxes"]:
        if m["cls"] != "mark":
            continue
        bb = m["bbox"]
        if bb["x"] <= cx <= bb["x"]+bb["w"] and bb["y"] <= cy <= bb["y"]+bb["h"]:
            return m
    return None


# ---------------------------------------------------------------- атомы тела
def _add_rect(slide, x, y, w, h, radius=0, radius_y=None):
    # радиусы закрывают ОБЕ полуоси (border-radius:50%) — эллипс/круг;
    # стадион-пилюля (rx = ry = h/2 < w/2) остаётся roundRect
    ry = radius if radius_y is None else radius_y
    if radius >= w / 2 - 1 and ry >= h / 2 - 1:
        st = MSO_SHAPE.OVAL
    elif radius > 0.5:
        st = MSO_SHAPE.ROUNDED_RECTANGLE
    else:
        st = MSO_SHAPE.RECTANGLE
    sp = slide.shapes.add_shape(st, Emu(int(x*PX)), Emu(int(y*PX)),
                                Emu(int(w*PX)), Emu(int(h*PX)))
    if st == MSO_SHAPE.ROUNDED_RECTANGLE:
        # PowerPoint for macOS can hang on PDF export when a slide combines a
        # picture with a roundRect whose adjustment is exactly the OOXML
        # boundary 50000.  Keep the geometry visually identical but strictly
        # inside the supported interval.
        sp.adjustments[0] = min(0.499, radius / max(1.0, min(w, h)))
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def _draw_data_line(slide, l, warnings):
    """<div data-line="…"> → нативный редактируемый коннектор (cxnSp).
    Геометрия: линия вдоль ДЛИННОЙ стороны бокса по его середине, толщина —
    короткая сторона (1px = 1pt), цвет — background-color элемента.
    Токены spec: arrow (наконечник в конце: вправо/вниз), arrow-start,
    arrow-both, dashed. Стиль наконечника — как у atom:arrow палитры (lg/med)."""
    from pptx.enum.shapes import MSO_CONNECTOR
    x, y, w, h = (l["bbox"][k] for k in "xywh")
    spec = set((l.get("spec") or "").split())
    horiz = w >= h
    thick = max(0.5, h if horiz else w)
    if thick > 12:
        warnings.append(f"data-line {l.get('cls') or ''!r}: бокс {w:.0f}×{h:.0f} не похож "
                        f"на линию (толщина {thick:.0f}px) — рисую вдоль длинной стороны")
    if horiz:
        p1, p2 = (x, y + h/2), (x + w, y + h/2)
    else:
        p1, p2 = (x + w/2, y), (x + w/2, y + h)
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Emu(int(p1[0]*PX)), Emu(int(p1[1]*PX)),
        Emu(int(p2[0]*PX)), Emu(int(p2[1]*PX)))
    conn.line.color.rgb = rgb(l["color"])
    conn.line.width = Pt(thick)
    conn.shadow.inherit = False
    ln = conn.line._get_or_add_ln()
    if "dashed" in spec:
        etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
    # порядок детей a:ln по схеме: fill, prstDash, headEnd, tailEnd
    for tag, on in (("a:headEnd", {"arrow-start", "arrow-both"} & spec),
                    ("a:tailEnd", {"arrow", "arrow-both"} & spec)):
        e = etree.SubElement(ln, qn(tag))
        e.set("type", "arrow" if on else "none")
        e.set("w", "lg")
        e.set("len", "med")
    return conn


def _split_series_colors(series, tag: str):
    """Пер-серийные опции (0.37.0 цвет, 0.38.0 метки): значение серии —
    список чисел ИЛИ {"values": [...], "color": "#RRGGBB", "labels": false}.
    Возвращает (series в прежнем виде для _norm_series, {имя: RGBColor},
    {имя: labels_off}). Кривой hex — честная ошибка ДО вставки чарта."""
    def strip(name, v):
        if isinstance(v, dict):
            col = v.get("color")
            if col:
                h = str(col).lstrip("#")
                if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
                    raise Html2PptxError(
                        f"{tag}: color {col!r} серии «{name}» — нужен hex "
                        f"#RRGGBB (цвет бери из палитры шаблона)")
                colors[str(name)] = RGBColor.from_string(h.upper())
            if v.get("labels") is False:       # серия-контекст без цифр
                labels_off[str(name)] = True
            return v.get("values")
        return v
    colors: dict[str, "RGBColor"] = {}
    labels_off: dict[str, bool] = {}
    if isinstance(series, dict):
        clean = {k: strip(k, v) for k, v in series.items()}
    else:
        clean = [[p[0], strip(p[0], p[1] if len(p) > 1 else None)]
                 for p in series]
    return clean, colors, labels_off


# канон: у плотной линии подписывать только ключевые точки (user-правило
# 2026-08-10 «если точек много — только ключевые»); порог — больше этого
_LABELS_SPARSE_FROM = 9


def _style_chart_canon(chart, t: str, spec: dict) -> None:
    """Канон стиля чартов брендов (user-правило 2026-08-10, «фундамент всех
    шаблонов»; база сверена с 6 дизайнерскими чартами Еды — #35/40/41/45:
    majorGridlines=0, valAx c:delete=1, showVal=1): БЕЗ сетки на фоне, БЕЗ
    левой оси значений — цифры читаются с меток прямо на графике; ось
    категорий остаётся с подписями и ВИДИМОЙ нулевой линией — тонкой, цветом
    accent3 ТЕМЫ шаблона (эталон-слайд Лики 2026-08-10 «нулевая линия должна
    быть видна»: schemeClr accent3, дефолтная толщина, тиков нет — тема даёт
    каждому бренду свой оттенок). spec["axis"]=true — осознанный возврат оси
    значений (плотный документ-режим). pie/doughnut осей не имеют."""
    if t in ("pie", "doughnut"):
        return
    from pptx.enum.chart import XL_TICK_MARK
    from pptx.enum.dml import MSO_THEME_COLOR
    val_ax, cat_ax = chart.value_axis, chart.category_axis
    val_ax.has_major_gridlines = False
    val_ax.has_minor_gridlines = False
    cat_ax.has_major_gridlines = False
    if not spec.get("axis"):
        val_ax.visible = False
    cat_ax.major_tick_mark = XL_TICK_MARK.NONE
    cat_ax.minor_tick_mark = XL_TICK_MARK.NONE
    cat_ax.format.line.color.theme_color = MSO_THEME_COLOR.ACCENT_3


def _key_point_indices(vals) -> list[int]:
    """Ключевые точки серии: первая, последняя, минимум, максимум (None —
    пропуски, не точки)."""
    real = [(i, v) for i, v in enumerate(vals) if v is not None]
    if not real:
        return []
    idxs = {real[0][0], real[-1][0],
            min(real, key=lambda p: p[1])[0], max(real, key=lambda p: p[1])[0]}
    return sorted(idxs)


def _dlbl_show_flags(parent, show_val: bool) -> None:
    """ПОЛНЫЙ набор show*-флагов метки (порядок схемы). Писать только showVal
    нельзя: отсутствующие флаги PowerPoint добивает СВОИМИ дефолтами — у
    каждой точки вылезали «Имя серии, категория» (живая проба 2026-08-10);
    дизайнерские чарты Еды тоже несут полный набор."""
    for tag, val in (("c:showLegendKey", "0"),
                     ("c:showVal", "1" if show_val else "0"),
                     ("c:showCatName", "0"), ("c:showSerName", "0"),
                     ("c:showPercent", "0"), ("c:showBubbleSize", "0")):
        etree.SubElement(parent, qn(tag)).set("val", val)


def _series_dlbls(ser_el, show: bool, key_idxs=(), pos: str | None = None):
    """Пер-серийные метки данных: c:dLbls на c:ser переопределяет плотовые
    настройки (порядок детей c:ser: … marker, dPt, dLbls, cat — вставляем
    перед c:cat/c:val). show=False + key_idxs — канон плотной линии: серия
    без меток, ключевые точки с меткой (dLbl idx + showVal=1)."""
    dlbls = etree.SubElement(ser_el, qn("c:dLbls"))
    for i in key_idxs:
        dlbl = etree.SubElement(dlbls, qn("c:dLbl"))
        etree.SubElement(dlbl, qn("c:idx")).set("val", str(i))
        if pos:
            etree.SubElement(dlbl, qn("c:dLblPos")).set("val", pos)
        _dlbl_show_flags(dlbl, show_val=True)
    if pos and show:
        etree.SubElement(dlbls, qn("c:dLblPos")).set("val", pos)
    _dlbl_show_flags(dlbls, show_val=show)
    anchor = ser_el.find(qn("c:cat"))
    if anchor is None:
        anchor = ser_el.find(qn("c:val"))
    if anchor is not None:
        anchor.addprevious(dlbls)
    return dlbls


def _paint_series(chart, t: str, colors, warnings, tag: str) -> None:
    """Покрасить серии заказанным цветом: line — линия (+маркеры у
    line_markers), column/bar/area — заливка. У pie/doughnut цвет СЕРИИ смысла
    не имеет (красятся точки-доли темой) — честное предупреждение."""
    if not colors:
        return
    if t in ("pie", "doughnut"):
        warnings.append(f"{tag}: color на {t} не применяется — доли красит "
                        f"accent-ряд темы по точкам; нужного акцента нет — "
                        f"смени тип или скажи владельцу скилла")
        return
    for s in chart.plots[0].series:
        col = colors.get(s.name)
        if col is None:
            continue
        if t in ("line", "line_markers"):
            s.format.line.color.rgb = col
            if t == "line_markers":
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = col
                s.marker.format.line.color.rgb = col
        else:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = col


# CSS-токен типа → XL_CHART_TYPE (атрибут по имени, чтобы не тянуть enum на импорте)
_CHART_TYPES = {
    "column": "COLUMN_CLUSTERED", "column_stacked": "COLUMN_STACKED",
    "bar": "BAR_CLUSTERED", "bar_stacked": "BAR_STACKED",
    "line": "LINE", "line_markers": "LINE_MARKERS",
    "area": "AREA", "pie": "PIE", "doughnut": "DOUGHNUT",
}


def _draw_chart(slide, ch, warnings, tdir=TP.DEFAULT_TDIR):
    """<div data-chart='{...}'> → НАТИВНЫЙ PowerPoint Chart (graphicFrame со
    встроенной Excel-книгой — python-pptx создаёт её сам): «Изменить данные»
    работает штатно, значения/серии/категории правятся в PowerPoint (фидбек
    переноса HTML 2026-08-10, P0). Спека:
      {"type": "column|column_stacked|bar|bar_stacked|line|line_markers|
                area|pie|doughnut",
       "categories": [...],
       "series": {"Имя": [числа|null],
                  "Цель": {"values": [...], "color": "#B8B5B2",
                           "labels": false}},
       "labels": true|false|"all", "number_format": "0%", "legend": true,
       "axis": true}
    Канон стиля (user-правило 2026-08-10, дефолт): без сетки, без оси
    значений — метки данных ВКЛЮЧЕНЫ по умолчанию (цифры прямо на графике),
    ось категорий без линии; плотная линия (>9 точек) подписывает только
    ключевые точки (первая/последняя/мин/макс; "labels": "all" — все);
    "axis": true возвращает ось значений, "labels": false глушит метки.
    Бокс элемента = рамка чарта (1px = 1pt); потомки элемента — заглушка для
    браузерного превью, в .pptx не идут (их скипает обмерщик). Цвета серий —
    от ТЕМЫ шаблона (accent-ряд дизайнера); точечный override — "color" на
    серии (0.37.0, цвет из палитры шаблона). Шрифт — фолбэк-стиль профиля.
    Комбо «столбцы + линия» одним чартом python-pptx не строит — два
    data-chart рядом (custom-html.md). После вставки — readback-верификация
    тем же python-pptx (мера, не глаз), как у chart-доноров (lib/charts.py)."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    import charts as CH                     # _norm_series — общий парсер серий

    tag = f"data-chart .{(ch.get('cls') or '?').split()[0]}"
    if ch.get("chartError"):
        raise Html2PptxError(f"{tag}: {ch['chartError']}")
    spec = ch.get("chart") or {}
    t = str(spec.get("type") or "").strip().lower()
    if t not in _CHART_TYPES:
        raise Html2PptxError(
            f"{tag}: type {t or '—'!r} не поддержан; доступны: "
            f"{', '.join(sorted(_CHART_TYPES))} (комбо — два data-chart рядом)")
    cats = [str(c) for c in spec.get("categories") or []]
    clean, colors, labels_off = _split_series_colors(spec.get("series") or {}, tag)
    try:
        sers = CH._norm_series(clean)
    except Exception as e:
        raise Html2PptxError(f"{tag}: series не разобрались ({e}) — нужен "
                             f"{{\"Имя\": [числа]}} или [[имя, [числа]], …]") from e
    if not cats or not sers:
        raise Html2PptxError(f"{tag}: нужны categories и series с данными "
                             f"пользователя (no-fabrication)")
    for name, vals in sers:
        if len(vals) != len(cats):
            raise Html2PptxError(f"{tag}: серия «{name}» несёт {len(vals)} "
                                 f"значений при {len(cats)} категориях")
    if t in ("pie", "doughnut") and len(sers) > 1:
        raise Html2PptxError(f"{tag}: {t} — одна серия (доли), а заказано "
                             f"{len(sers)}; сравнение серий — column/bar")
    if len(sers) > 3:
        warnings.append(f"{tag}: {len(sers)} серий — канон держит ≤3 "
                        f"(chart-selection.md §4): главная одна, прочие контекст")

    x, y, w, h = (ch["bbox"][k] for k in "xywh")
    data = CategoryChartData(number_format=spec.get("number_format") or "General")
    data.categories = cats
    for name, vals in sers:
        data.add_series(name, tuple(vals))
    gf = slide.shapes.add_chart(
        getattr(XL_CHART_TYPE, _CHART_TYPES[t]),
        Emu(int(x*PX)), Emu(int(y*PX)), Emu(int(w*PX)), Emu(int(h*PX)), data)
    chart = gf.chart
    chart.font.name = TP.load(tdir)["fonts"]["fallback"]   # текст чарта — бренд
    chart.font.size = Pt(16)
    # заголовок чарта — вывод живёт в заголовке СЛАЙДА (Rule 2), авто-титул
    # PP (имя серии) гасим; spec.title — осознанное включение
    if spec.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = str(spec["title"])
    else:
        chart.has_title = False
    want_legend = spec.get("legend")
    if want_legend is None:
        want_legend = len(sers) > 1 or t in ("pie", "doughnut")
    chart.has_legend = bool(want_legend)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    _style_chart_canon(chart, t, spec)         # без сетки и оси значений
    want_labels = spec.get("labels")
    if want_labels is None:
        want_labels = True     # канон: без оси цифры читаются только с меток
    if want_labels:
        plot = chart.plots[0]
        plot.has_data_labels = True
        plot.data_labels.show_value = True     # has_data_labels сам не включает
        if spec.get("number_format"):          # showVal (живая проба doughnut)
            plot.data_labels.number_format = spec["number_format"]
            plot.data_labels.number_format_is_linked = False
        pos = "t" if t in ("line", "line_markers") else None
        sparse = (t in ("line", "line_markers", "area")
                  and len(cats) > _LABELS_SPARSE_FROM and want_labels != "all")
        ser_els = chart._chartSpace.findall(f".//{qn('c:ser')}")
        for ser_el, (name, vals) in zip(ser_els, sers):
            if labels_off.get(name):
                _series_dlbls(ser_el, show=False)
            elif sparse:                       # плотная линия: ключевые точки
                _series_dlbls(ser_el, show=False,
                              key_idxs=_key_point_indices(vals), pos=pos)
            elif pos:                          # line: метки над точками
                _series_dlbls(ser_el, show=True, pos=pos)
    _paint_series(chart, t, colors, warnings, tag)

    # readback: то, что читается из вставленного чарта, == заказанному
    plot = chart.plots[0]
    got_cats = [str(c) for c in plot.categories]
    got = [(s.name, [None if v is None else float(v) for v in s.values])
           for s in plot.series]
    if got_cats != cats or got != [(n, v) for n, v in sers]:
        raise Html2PptxError(f"{tag}: readback вставленного чарта разошёлся "
                             f"с заказом — сообщи владельцу скилла "
                             f"(scripts/report_issue.py)")
    return gf


def _grad_fill(sp, grad: str):
    """CSS linear-gradient → a:gradFill (линейный, стопы с альфой)."""
    inner = grad[grad.index("(")+1:grad.rindex(")")]
    parts = re.split(r",(?![^()]*\))", inner)
    css_deg = 180.0                 # CSS-дефолт: to bottom
    if re.match(r"\s*[-\d.]+deg", parts[0]):
        css_deg = float(parts[0].strip().replace("deg", ""))
        parts = parts[1:]
    stops = []
    for i, p in enumerate(parts):
        m = re.match(r"\s*(rgba?\([^)]*\))\s*([\d.]+)?%?", p)
        if not m:
            continue
        v = [float(x) for x in re.findall(r"[\d.]+", m.group(1))]
        a = v[3] if len(v) > 3 else 1.0
        pos = float(m.group(2)) if m.group(2) is not None else (100.0 * i / max(1, len(parts)-1))
        stops.append((v[:3], a, pos))
    spPr = sp._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        e = spPr.find(qn(tag))
        if e is not None:
            spPr.remove(e)
    gf = etree.SubElement(spPr, qn("a:gradFill"))
    gf.set("rotWithShape", "1")
    gs_lst = etree.SubElement(gf, qn("a:gsLst"))
    for col, a, pos in stops:
        gs = etree.SubElement(gs_lst, qn("a:gs"))
        gs.set("pos", str(int(pos*1000)))
        c = etree.SubElement(gs, qn("a:srgbClr"))
        c.set("val", "%02X%02X%02X" % (int(col[0]), int(col[1]), int(col[2])))
        al = etree.SubElement(c, qn("a:alpha"))
        al.set("val", str(int(a*100000)))
    lin = etree.SubElement(gf, qn("a:lin"))
    lin.set("ang", str(int(((css_deg - 90) % 360) * 60000)))
    lin.set("scaled", "1")
    # схема spPr: xfrm, geom, fill, ln, effectLst — заливку СРАЗУ после геометрии,
    # иначе PP молча игнорирует её и откатывается на стилевую заливку темы
    anchor = spPr.find(qn("a:prstGeom"))
    if anchor is None:
        anchor = spPr.find(qn("a:custGeom"))
    if anchor is None:
        anchor = spPr.find(qn("a:xfrm"))
    if anchor is not None:
        anchor.addnext(gf)


def _draw_box(slide, b, slide_bg, painted):
    x, y, w, h = (b["bbox"][k] for k in "xywh")
    if b["bg"] or b.get("grad") or (b["borderC"] and b["borderW"] > 0):
        sp = _add_rect(slide, x, y, w, h, b["radius"], b.get("radiusY"))
        if b.get("grad"):
            _grad_fill(sp, b["grad"])
        elif b["bg"]:
            final = rgb(b["bg"], b.get("opacity", 1), slide_bg)
            sp.fill.solid()
            sp.fill.fore_color.rgb = final
            painted.append((b["bbox"], (final[0], final[1], final[2])))
        else:
            sp.fill.background()
        if b["borderC"] and b["borderW"] > 0:
            sp.line.color.rgb = rgb(b["borderC"])
            sp.line.width = Pt(b["borderW"])
    for side, rect in (("t", (x, y, w, 0)), ("b", (x, y+h, w, 0)),
                       ("l", (x, y, 0, h)), ("r", (x+w, y, 0, h))):
        s = (b.get("sides") or {}).get(side)
        if not s:
            continue                # рамка одной стороны → тонкий прямоугольник
        sx, sy, sw_, sh_ = rect
        if side in "tb":
            sy -= s["w"] if side == "b" else 0
            sh_ = s["w"]
        else:
            sx -= s["w"] if side == "r" else 0
            sw_ = s["w"]
        sp = _add_rect(slide, sx, sy, sw_ or w, sh_ or h)
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(s["c"], b.get("opacity", 1), slide_bg)


def _draw_image(slide, im, warnings):
    x, y, w, h = (im["frame"][k] for k in "xywh")
    src = urllib.parse.unquote(im["src"].replace("file://", ""))
    if not os.path.exists(src):
        raise Html2PptxError(f"фото custom-слайда не найдено: {src}")
    pic = slide.shapes.add_picture(src, Emu(int(x*PX)), Emu(int(y*PX)),
                                   Emu(int(w*PX)), Emu(int(h*PX)))
    if im.get("fit") == "cover" and im["natW"] and im["natH"]:
        _set_srcrect(pic._element, cover_fit_crop(im["natW"], im["natH"], w, h))
    elif im.get("fit") not in (None, "cover", "fill"):
        warnings.append(f"object-fit:{im['fit']} не поддержан (только cover) — "
                        f"картинка ляжет с искажением аспекта")
    if im.get("radius", 0) > 0.5:   # скругление рамки → prstGeom roundRect
        spPr = pic._element.spPr
        pg = spPr.find(qn("a:prstGeom"))
        if pg is not None:
            spPr.remove(pg)
        pg = etree.SubElement(spPr, qn("a:prstGeom"))
        pg.set("prst", "roundRect")
        av = etree.SubElement(pg, qn("a:avLst"))
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", "adj")
        gd.set("fmla", "val %d" % min(49900, int(im["radius"] / max(1.0, min(w, h)) * 100000)))
        ext = spPr.find(qn("a:xfrm"))   # prstGeom — после xfrm
        if ext is not None:
            ext.addnext(pg)
    pic.shadow.inherit = False
    eff = (im["natW"] / w) * 72 if w else 0
    if eff and eff < 96:
        warnings.append(f"фото {os.path.basename(src)}: ~{int(eff)} dpi на этом "
                        f"боксе — будет мыло; возьмите крупнее")


def _draw_svg(slide, s):
    x, y, w, h = (s["bbox"][k] for k in "xywh")
    scale = 4
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False) as f:
        f.write("<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>"
                "html,body{margin:0;background:transparent}"
                "svg{display:block;width:%dpx;height:%dpx}</style></head><body>%s</body></html>"
                % (round(w*scale), round(h*scale), s["svg"]))
        tmp_html = f.name
    png_handle = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    png = png_handle.name
    png_handle.close()
    # Флаги — эмпирика живых проб 2026-07-23/24 (0.21.0): свой --user-data-dir
    # headless=new НЕ давать (виснет); mock-keychain/basic гасят Keychain-диалог
    # macOS; тихие флаги глушат GoogleUpdater. Зеркало — html_preview.py,
    # html_measure.mjs.
    try:
        subprocess.run([chrome_path(), "--headless=new", "--disable-gpu",
                        "--no-first-run", "--no-default-browser-check",
                        "--use-mock-keychain", "--password-store=basic",
                        *CHROME.security_flags(),
                        "--default-background-color=00000000", "--hide-scrollbars",
                        "--window-size=%d,%d" % (round(w*scale), round(h*scale)),
                        "--screenshot=" + png, "file://" + tmp_html],
                       check=True, capture_output=True, timeout=60)
        slide.shapes.add_picture(png, Emu(int(x*PX)), Emu(int(y*PX)),
                                 Emu(int(w*PX)), Emu(int(h*PX)))
    finally:
        for path in (tmp_html, png):
            try:
                os.unlink(path)
            except OSError:
                pass


ALIGN = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT, "center": PP_ALIGN.CENTER,
         "start": PP_ALIGN.LEFT, "end": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}
VALIGN = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
NO_STYLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # встроенный стиль PP «без стиля, без сетки»


def _cell_borders(cell, sides):
    """CSS-границы ячейки по сторонам → a:lnL/R/T/B в tcPr. Порядок схемы
    CT_TableCellProperties: ln* строго ПЕРЕД заливкой — вставляем в начало."""
    tcPr = cell._tc.get_or_add_tcPr()
    made = []
    for key, tag in (("l", "a:lnL"), ("r", "a:lnR"), ("t", "a:lnT"), ("b", "a:lnB")):
        s = (sides or {}).get(key)
        if not s:
            continue
        ln = tcPr.makeelement(qn(tag), {})
        ln.set("w", str(int(round(s["w"] * PX))))         # 1px = 1pt = 12700 EMU
        ln.set("cap", "flat")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(sf, qn("a:srgbClr"))
        col = rgb(s["c"]) if s["c"] else RGBColor(0, 0, 0)
        clr.set("val", "%02X%02X%02X" % (col[0], col[1], col[2]))
        made.append(ln)
    for ln in reversed(made):
        tcPr.insert(0, ln)


def _draw_table(slide, t, slide_bg, warnings, tdir=TP.DEFAULT_TDIR):
    """<table> из обмера → НАТИВНЫЙ редактируемый a:tbl (не линии + текстбоксы).
    Стиль — «No Style, No Grid»: заливки/границы/шрифты каждой ячейки задаются
    явно из обмера; оверлеи (chip/mark в ячейках) рисует общий конвейер поверх,
    их текст из ранов ячейки исключён ещё обмерщиком."""
    x, y, w, h = (t["bbox"][k] for k in "xywh")
    rows, cols = t["nRows"], t["nCols"]
    if not rows or not cols:
        warnings.append(f"таблица .{t['cls']}: пустая сетка — пропущена")
        return
    gf = slide.shapes.add_table(rows, cols, Emu(int(x*PX)), Emu(int(y*PX)),
                                Emu(int(w*PX)), Emu(int(h*PX)))
    tbl = gf.table
    tbl.first_row = tbl.horz_banding = False
    sid = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = NO_STYLE_NO_GRID
    for i, cw in enumerate(t["colW"]):
        tbl.columns[i].width = Emu(int(cw*PX))
    for j, rh in enumerate(t["rowH"]):
        tbl.rows[j].height = Emu(int(rh*PX))
    for c in t["cells"]:
        cell = tbl.cell(c["r"], c["c"])
        if c["rowSpan"] > 1 or c["colSpan"] > 1:
            cell.merge(tbl.cell(c["r"] + c["rowSpan"] - 1,
                                c["c"] + c["colSpan"] - 1))
        if c["bg"]:
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(c["bg"], 1, slide_bg)
        else:
            cell.fill.background()
        _cell_borders(cell, c.get("sides"))
        cell.margin_left = Emu(int(c["padL"]*PX))
        cell.margin_right = Emu(int(c["padR"]*PX))
        cell.margin_top = Emu(int(c["padT"]*PX))
        cell.margin_bottom = Emu(int(c["padB"]*PX))
        cell.vertical_anchor = VALIGN.get(c["valign"], MSO_ANCHOR.TOP)
        runs = [r for r in c.get("runs") or []]
        if not [r for r in runs if not r.get("br")]:
            continue
        _strip_paragraph_periods(runs)   # ячейки — как значения fill-пути
        tf = cell.text_frame
        tf.word_wrap = True
        align = ALIGN.get(c["align"], PP_ALIGN.LEFT)
        under = tuple(rgb(c["bg"], 1, slide_bg)) if c["bg"] else tuple(slide_bg[:3])
        p = tf.paragraphs[0]
        p.alignment = align
        if c["lineHeightPx"]:
            p.line_spacing = Pt(c["lineHeightPx"])
        for r in runs:
            if r.get("br"):
                p = tf.add_paragraph()
                p.alignment = align
                if c["lineHeightPx"]:
                    p.line_spacing = Pt(c["lineHeightPx"])
                continue
            _add_run(p, r, under, tdir)
        # ячейку не расширить на PP-запас (+~0.8% без кернинга) — ловим впритык
        if c["nLines"] == 1:
            inner = c["bbox"]["w"] - c["padL"] - c["padR"]
            wpp = sum(pp_width(r["text"], font_name(r["family"], r["weight"], tdir)[0],
                               r["sizePx"], tdir) for r in runs if not r.get("br"))
            if wpp > inner + 0.5:
                txt = " ".join(r["text"] for r in runs if not r.get("br"))
                warnings.append(f"таблица .{t['cls']}: текст {txt[:40]!r} шире "
                                f"ячейки в PP-метрике — возможен лишний перенос; "
                                f"reword или шире колонка")


def _strip_paragraph_periods(runs):
    """Правило точек (правило user 2026-07-02, зеркало fill._typo) для
    custom-пути: самостоятельный абзац не заканчивается точкой — срезаем
    финальную точку КАЖДОГО абзаца (сегменты между <br> и конец текста);
    «..» и «…» живут, внутренние точки не трогаются. Fill-путь делает это в
    typo_walk по строкам deckspec; кастомные тексты приходят HTML-файлом
    МИМО спеки — до 0.40.1 деки расходились типографикой (вопрос Лики
    2026-08-13 по слайдам 4/11/12 дека Ultima: каждый абзац с точкой).
    NBSP-часть _typo сюда сознательно НЕ переносится: браузер мерил текст с
    обычными пробелами — замена на NBSP сдвигает точки переноса PP и рвёт
    пиксельный паритет; срез точки строку только укорачивает — безопасен."""
    last = None
    for r in runs:
        if r.get("br"):
            if last is not None:
                t = (last.get("text") or "").rstrip()
                if t.endswith(".") and not t.endswith(("..", "…")):
                    last["text"] = t[:-1]
            last = None
        else:
            last = r
    if last is not None:
        t = (last.get("text") or "").rstrip()
        if t.endswith(".") and not t.endswith(("..", "…")):
            last["text"] = t[:-1]


# Канон буллита — atom:block_title, слайд #42 шаблона Хурала (утверждённый
# дизайнером вид; снят с XML эталона — хэндофф Ultima 2026-08-11, O2):
# marL 468000 EMU (36.85pt) с висячей строкой, маркер «•» Arial 150% цветом
# текста, интервалы 6pt до / 16pt после абзаца.
_BULLET_MARL_EMU = 468000
_BULLET_HANG_PX = _BULLET_MARL_EMU / PX          # 36.85px (1px = 1pt)


def _apply_bullet(p, kind: str, idx: int = 0, cont: bool = False):
    """<li> → нативный абзац-буллит (раньше маркеры терялись вместе с ::marker,
    и перечисление в custom можно было показать только литеральным «—» — его
    дизайнер бракует, потому что рядом на донорах буллиты настоящие).
    kind: 'ul' — buChar «•», 'ol' — buAutoNum arabicPeriod (каждый li — свой
    текстбокс, нумерация продолжается через startAt=idx+1). cont — абзац-
    продолжение того же li (после <br>): отступ тот же, маркера нет."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(_BULLET_MARL_EMU))
    pPr.set("indent", str(-_BULLET_MARL_EMU))
    # порядок детей по схеме CT_TextParagraphProperties: lnSpc (его пишет
    # line_spacing раньше нас), spcBef, spcAft, bu* — append соблюдает его.
    # spcBef канона (6pt) — ТОЛЬКО абзацам-продолжениям: на первом абзаце
    # фрейма PP применил бы его тоже и увёл текст ниже браузерного обмера;
    # интервалы МЕЖДУ li даёт геометрия боксов (margin в CSS). spcAft — канон
    # на случай, когда дизайнер дописывает буллеты Enter'ом внутри бокса.
    if cont:
        sp = etree.SubElement(pPr, qn("a:spcBef"))
        etree.SubElement(sp, qn("a:spcPts")).set("val", "600")
    sp = etree.SubElement(pPr, qn("a:spcAft"))
    etree.SubElement(sp, qn("a:spcPts")).set("val", "1600")
    if cont:
        etree.SubElement(pPr, qn("a:buNone"))
        return
    etree.SubElement(pPr, qn("a:buClrTx"))
    etree.SubElement(pPr, qn("a:buSzPct")).set("val", "150000")
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", "Arial")
    bf.set("panose", "020B0604020202020204")
    bf.set("pitchFamily", "34")
    bf.set("charset", "0")
    if kind == "ol":
        an = etree.SubElement(pPr, qn("a:buAutoNum"))
        an.set("type", "arabicPeriod")
        if idx:
            an.set("startAt", str(idx + 1))
    else:
        etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def _draw_texts(slide, data, bg_under, tdir=TP.DEFAULT_TDIR, skip=None) -> int:
    """skip — заголовок, уже уложенный в TITLE-плейсхолдер (по идентичности).
    Прочие .h1/role=title больше НЕ глотаются молча — лягут текстбоксами тела
    (предупреждение выдал build_custom_slide)."""
    n = 0
    for t in data["texts"]:
        if t is skip or t["cls"].startswith("pagenum"):
            continue                # это даёт хром
        n += 1
        li = t.get("li") or None    # <li> идёт абзац-веткой даже одной строкой
        _strip_paragraph_periods(t["runs"])   # правило точек (зеркало fill._typo)
        runs = [r for r in t["runs"] if not r.get("br")] if t["nLines"] == 1 else t["runs"]
        if t["nLines"] == 1 and not li:
            for r in runs:
                rect = r.get("rect") or t["lines"][0]
                tb = slide.shapes.add_textbox(Emu(int(rect["x"]*PX)), Emu(int(rect["y"]*PX)),
                                              Emu(int((rect["w"]+6)*PX)), Emu(int(rect["h"]*PX)))
                tf = tb.text_frame
                tf.word_wrap = False
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf.auto_size = None
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                _add_run(p, r, bg_under(rect), tdir)
            continue
        first = t["lines"][0]
        # буллит: бокс сдвигается влево на висячую строку — текст остаётся на
        # браузерном датуме (marL), маркер живёт в полосе паддинга <ul>
        hang = _BULLET_HANG_PX if li else 0
        x = t["bbox"]["x"] + t["padL"] - hang
        y = first["y"]
        # +1.2%: PP не кернит (GPOS) → текст шире браузерного, при боксе впритык
        # к самой длинной строке PP перенёс бы лишнее слово вниз
        w = (t["bbox"]["w"] - t["padL"] - t["padR"]) * 1.012 + 2 + hang
        h = t["lines"][-1]["y"] + t["lines"][-1]["h"] - first["y"]
        tb = slide.shapes.add_textbox(Emu(int(x*PX)), Emu(int(y*PX)),
                                      Emu(int(w*PX)), Emu(int(h*PX)))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.auto_size = None
        align = ALIGN.get(t["align"], PP_ALIGN.LEFT)
        p = tf.paragraphs[0]
        p.alignment = align
        if t["lineHeightPx"]:
            p.line_spacing = Pt(t["lineHeightPx"])
        if li:
            _apply_bullet(p, li, t.get("liIdx", 0))
        under = bg_under({"x": x, "y": y, "w": w, "h": h})
        for r in t["runs"]:
            if r.get("br"):
                p = tf.add_paragraph()
                p.alignment = align
                if t["lineHeightPx"]:
                    p.line_spacing = Pt(t["lineHeightPx"])
                if li:                     # продолжение li: отступ без маркера
                    _apply_bullet(p, li, cont=True)
                continue
            _add_run(p, r, under, tdir)
    return n


def _add_run(p, r, under, tdir=TP.DEFAULT_TDIR):
    run = p.add_run()
    run.text = r["text"]
    name, bold = font_name(r["family"], r["weight"], tdir)
    run.font.name = name
    run.font.bold = bold
    run.font.size = Pt(r["sizePx"])
    run.font.italic = r.get("italic", False)
    if r["color"]:
        run.font.color.rgb = rgb(r["color"], r.get("effOp", 1), under)


# ---------------------------------------------------------------- обмер
def measure_html(html_path: str, slide_idx: int, skill_dir: str,
                 tdir: str = TP.DEFAULT_TDIR) -> dict:
    """Прогнать scripts/html_measure.mjs → dict обмера. Требует node + Chrome.
    tdir → сверка канвы обмера с профилем шаблона."""
    import json as _json
    mjs = os.path.join(skill_dir, "scripts", "html_measure.mjs")
    validate_html_file(html_path)
    out_handle = tempfile.NamedTemporaryFile(suffix=".json", delete=False)
    out = out_handle.name
    out_handle.close()
    # передаём резолвнутый путь в node — html_measure.mjs не дублирует поиск
    env = dict(os.environ, CHROME_PATH=chrome_path())
    try:
        # encoding явно: node пишет UTF-8, а text=True декодирует локалью ОС —
        # на русской Windows (cp1251) stderr приходил мохнатым «РёР·РјРµСЂ…»
        # (хэндофф Ultima 2026-08-11, O6.3)
        subprocess.run(["node", mjs, html_path, str(slide_idx), out],
                       check=True, capture_output=True, timeout=120, env=env,
                       encoding="utf-8", errors="replace")
    except FileNotFoundError as e:
        raise Html2PptxError("node не найден — нужен для family:custom "
                             "(обмер HTML)") from e
    except subprocess.CalledProcessError as e:
        raise Html2PptxError(f"обмер HTML упал: {e.stderr.strip()[:300]}") from e
    try:
        with open(out, encoding="utf-8") as f:
            data = _json.load(f)
    finally:
        try:
            os.unlink(out)
        except OSError:
            pass
    _validate_measured_canvas(data, tdir)
    return data
