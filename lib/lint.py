#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""lint.py — статический ИЗМЕРИТЕЛЬНЫЙ линт до рендера (plan §7 A4/A9/A12, §12).

Меряет через lib.fit (тот же измеритель, что у рендерера) и ловит то, что
визуальный подагент стабильно пропускает, потому что это НАДО МЕРИТЬ, а не «на глаз»:

  overflow   — текст шире/выше своей зоны:
               * wrap=off (слот значения/стата): берём фактический НАРИСОВАННЫЙ
                 экстент (с учётом выравнивания: лево/центр/право), проверяем,
                 не вылез ли он за холст или в соседний шейп. Сам факт «шире бокса»
                 для центрированного дисплей-числа — норма, поэтому меряем экстент,
                 а не сравниваем с шириной бокса.
               * wrap=on: переносим по ширине, сравниваем высоту блока с высотой
                 зоны (вертикальный overflow / обрезка); плюс одиночный токен шире
                 зоны = горизонтальный overflow (ровно кейс «40%» из §12, когда «%»
                 переносился в боксе под «193»).
  empty_lines — трейлинг <a:br/> и пустые абзацы (A12): в центрированном боксе
                пустая строка уводит текст вверх. paragraph.text их не показывает.
  collisions  — нарисованные рамки шейпов пересекаются (A9).
  containment — текст, сидящий на плашке (roundRect/rect с заливкой), не выходит
                за её края (машинный слой B10; кейс «180+/45 тыс/12 млн»
                2026-07-03, когда значение длиннее дизайнерской заглушки
                вырастало за левый край карточки).

Запуск:
    python3 lib/lint.py deck.pptx            # текст
    python3 lib/lint.py deck.pptx --json
Exit: 1, если есть HIGH.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fit  # noqa: E402

EMU = fit.EMU_PER_INCH
NBSP = fit.NBSP

# Дефолты OOXML, если атрибут не задан в bodyPr
_DEF_INS = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
_DEF_LINESPC = 1.2


@dataclass
class Finding:
    severity: str  # HIGH | MEDIUM | LOW
    code: str
    message: str
    slide: int = 0
    shape: str = ""


@dataclass
class LintReport:
    file: str
    findings: list[Finding] = field(default_factory=list)
    stats: dict = field(default_factory=dict)

    @property
    def high(self) -> int:
        return sum(f.severity == "HIGH" for f in self.findings)

    @property
    def medium(self) -> int:
        return sum(f.severity == "MEDIUM" for f in self.findings)

    @property
    def low(self) -> int:
        return sum(f.severity == "LOW" for f in self.findings)


# ------------------------------------------------------ резолв шрифта и кегля


_theme_cache: dict = {}


def _slide_theme_fonts(slide) -> tuple[str | None, str | None]:
    """(+mj-lt, +mn-lt) ТЕМЫ ЭТОГО слайда. В деке несколько мастеров с разными
    темами (у Еды theme1 = YS, theme2/3 = Calibri) — глобально брать первую тему
    нельзя, иначе +mn-lt резолвится в чужой шрифт. Идём слайд→мастер→его тема."""
    from lxml import etree

    try:
        master_part = slide.slide_layout.slide_master.part
    except Exception:
        return (None, None)
    key = str(master_part.partname)
    if key in _theme_cache:
        return _theme_cache[key]
    res = (None, None)
    try:
        for rel in master_part.rels.values():
            if "theme" not in rel.reltype:
                continue
            root = etree.fromstring(rel.target_part.blob)
            fs = root.find(f"{qn('a:themeElements')}/{qn('a:fontScheme')}")
            maj = fs.find(f"{qn('a:majorFont')}/{qn('a:latin')}") if fs is not None else None
            mnr = fs.find(f"{qn('a:minorFont')}/{qn('a:latin')}") if fs is not None else None
            res = (maj.get("typeface") if maj is not None else None,
                   mnr.get("typeface") if mnr is not None else None)
            break
    except Exception:
        pass
    _theme_cache[key] = res
    return res


def _autofit(txBody) -> str:
    """Режим автоподгонки бокса: 'sp' (растёт под текст), 'norm' (кегль ужимается),
    'no'/'none'. Определяет, возможен ли вообще overflow."""
    b = txBody.find(qn("a:bodyPr"))
    if b is None:
        return "none"
    if b.find(qn("a:spAutoFit")) is not None:
        return "sp"
    if b.find(qn("a:normAutofit")) is not None:
        return "norm"
    if b.find(qn("a:noAutofit")) is not None:
        return "no"
    return "none"


def _resolve_scheme_font(name: str | None, theme: tuple) -> str | None:
    if name == "+mj-lt":
        return theme[0]
    if name == "+mn-lt":
        return theme[1]
    return name


def _defRPr_of(pPr_parent, lvl: int):
    """Достаём (font, sz_pt) из <a:lvl{lvl}pPr><a:defRPr> элемента-контейнера
    (lstStyle / txStyle level). None-поля, если нет."""
    if pPr_parent is None:
        return (None, None)
    pPr = pPr_parent.find(qn(f"a:lvl{lvl}pPr"))
    if pPr is None:
        return (None, None)
    dr = pPr.find(qn("a:defRPr"))
    if dr is None:
        return (None, None)
    lat = dr.find(qn("a:latin"))
    sz = dr.get("sz")
    return (lat.get("typeface") if lat is not None else None,
            int(sz) / 100.0 if sz else None)


def _layout_master_props(shape, slide, lvl):
    """Для плейсхолдера: пройти layout-плейсхолдер -> master-плейсхолдер ->
    master txStyles по типу. Best-effort, возвращает (font, sz_pt)."""
    if not shape.is_placeholder:
        return (None, None)
    pf = shape.placeholder_format
    idx, ptype = pf.idx, pf.type
    font = sz = None
    for src in (slide.slide_layout, slide.slide_layout.slide_master):
        try:
            for ph in src.placeholders:
                if ph.placeholder_format.idx == idx:
                    f2, s2 = _defRPr_of(ph.text_frame._txBody.find(qn("a:lstStyle")), lvl)
                    font = font or f2
                    sz = sz or s2
                    break
        except Exception:
            pass
    if font and sz:
        return (font, sz)
    # master txStyles по типу плейсхолдера
    try:
        txs = slide.slide_layout.slide_master.element.find(qn("p:txStyles"))
        tag = {1: "p:titleStyle", 13: "p:titleStyle"}.get(int(ptype) if ptype else -1, "p:bodyStyle")
        node = txs.find(qn(tag)) if txs is not None else None
        if node is not None:
            f2, s2 = _defRPr_of(node, lvl)
            # txStyles используют lvlNpPr без префикса lstStyle — _defRPr_of подходит
            font = font or f2
            sz = sz or s2
    except Exception:
        pass
    return (font, sz)


def effective_props(shape, paragraph, run, prs, slide, theme, default_style):
    """(font_name, size_pt): первое заданное по цепочке наследования OOXML.
    Возвращает (None, size) если шрифт не нашёлся (тогда замерить нельзя)."""
    lvl = (paragraph.level or 0) + 1
    txBody = shape.text_frame._txBody

    font = run.font.name if run is not None else None
    size = (run.font.size.pt if run is not None and run.font.size else None)

    # paragraph defRPr (pPr)
    if font is None or size is None:
        try:
            if paragraph.font.name and font is None:
                font = paragraph.font.name
            if paragraph.font.size and size is None:
                size = paragraph.font.size.pt
        except Exception:
            pass
    # txBody lstStyle lvlN
    if font is None or size is None:
        f2, s2 = _defRPr_of(txBody.find(qn("a:lstStyle")), lvl)
        font = font or f2
        size = size or s2
    # placeholder -> layout/master
    if font is None or size is None:
        f2, s2 = _layout_master_props(shape, slide, lvl)
        font = font or f2
        size = size or s2
    # presentation defaultTextStyle
    if font is None or size is None:
        f2, s2 = default_style.get(lvl, (None, None))
        font = font or f2
        size = size or s2

    font = _resolve_scheme_font(font, theme)
    return (font, size)


def _default_text_style(prs) -> dict:
    out = {}
    try:
        dts = prs.part._element.find(qn("p:defaultTextStyle"))
        if dts is not None:
            for lvl in range(1, 10):
                f, s = _defRPr_of(dts, lvl)
                if f or s:
                    out[lvl] = (f, s)
    except Exception:
        pass
    return out


# ----------------------------------------------------------- геометрия шейпа


def _bodypr(txBody):
    b = txBody.find(qn("a:bodyPr"))
    wrap = "square"
    ins = dict(_DEF_INS)
    if b is not None:
        wrap = b.get("wrap") or "square"
        for k in ins:
            v = b.get(k)
            if v is not None:
                ins[k] = int(v)
    return wrap, ins


def _v_anchor(txBody) -> str:
    """Вертикальный якорь бокса: 't' (верх, дефолт), 'ctr' (центр), 'b' (низ).
    Определяет, куда растёт многострочный текст при переполнении бокса."""
    b = txBody.find(qn("a:bodyPr"))
    a = b.get("anchor") if b is not None else None
    return a if a in ("t", "ctr", "b") else "t"


def _align(paragraph) -> str:
    a = paragraph.alignment
    if a is None:
        return "l"
    s = str(a)
    if "CENTER" in s:
        return "ctr"
    if "RIGHT" in s:
        return "r"
    if "JUSTIFY" in s:
        return "just"
    return "l"


def _line_spacing(paragraph, txBody, lvl) -> float:
    """Межстрочный коэффициент: сперва lnSpc на самом абзаце (pPr), затем
    lstStyle lvlN, иначе дефолт 1.2. На обложке Еды интервал 80% задан именно
    на абзаце — если смотреть только lstStyle, посчитаешь высоту завышенной."""
    pPr = paragraph._p.find(qn("a:pPr"))
    if pPr is not None:
        sp = pPr.find(f"{qn('a:lnSpc')}/{qn('a:spcPct')}")
        if sp is not None and sp.get("val"):
            return int(sp.get("val")) / 100000.0
    ls = txBody.find(qn("a:lstStyle"))
    if ls is not None:
        pPr = ls.find(qn(f"a:lvl{lvl}pPr"))
        if pPr is not None:
            sp = pPr.find(f"{qn('a:lnSpc')}/{qn('a:spcPct')}")
            if sp is not None and sp.get("val"):
                return int(sp.get("val")) / 100000.0
    return _DEF_LINESPC


# ------------------------------------------------------------------- линты


def _measurable_shapes(slide):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            yield sh


def lint_overflow(slide, si, prs, theme, default_style, canvas_w, canvas_h, rep,
                  extents=None):
    """extents (dict) — сюда складываются нарисованные экстенты текста по
    sh.shape_id для containment/collisions. Именно dict: python-pptx отдаёт
    ОДНОРАЗОВЫЕ прокси шейпов, атрибут на sh испаряется между пассами — до
    2026-07-07 containment из-за этого всегда мерил по сырому боксу и был слеп
    к вылету право-выровненных wrap=off значений (user-скриншот «~35 мин»)."""
    for sh in _measurable_shapes(slide):
        tf = sh.text_frame
        txBody = tf._txBody
        # авто-поля (номер слайда/дата) — значение задаёт PowerPoint, бокс
        # шаблонно-фиксированный; мерить нечего (иначе «38» в узком боксе = ложный overflow)
        if txBody.find(f".//{qn('a:fld')}") is not None:
            continue
        wrap, ins = _bodypr(txBody)
        left = sh.left / EMU
        top = sh.top / EMU
        bw = sh.width / EMU
        bh = sh.height / EMU
        inner_l = left + ins["lIns"] / EMU
        inner_r = left + bw - ins["rIns"] / EMU
        inner_w = max(inner_r - inner_l, 0.01)
        inner_t = top + ins["tIns"] / EMU
        inner_b = top + bh - ins["bIns"] / EMU
        inner_h = max(inner_b - inner_t, 0.01)

        para = tf.paragraphs[0]
        run = para.runs[0] if para.runs else None
        font, size = effective_props(sh, para, run, prs, slide, theme, default_style)
        label = f"{tf.text.strip()[:22]!r}"
        if not font or not size:
            rep.findings.append(Finding(
                "LOW", "unmeasured",
                f"не определён шрифт/кегль ({font}/{size}) — замер пропущен", si, label))
            continue

        lvl = (para.level or 0) + 1
        linespc = _line_spacing(para, txBody, lvl)
        af = _autofit(txBody)
        text = tf.text

        try:
            if wrap == "none":
                # одна строка, бокс по тексту не переносит. Меряем нарисованный
                # экстент с учётом выравнивания и проверяем выход за холст.
                w = max(fit.text_width_in(line, font, size)
                        for line in text.split("\n"))
                align = _align(para)
                if align == "ctr":
                    cx = (inner_l + inner_r) / 2
                    x0, x1 = cx - w / 2, cx + w / 2
                elif align == "r":
                    x0, x1 = inner_r - w, inner_r
                else:
                    x0, x1 = inner_l, inner_l + w
                if x1 > canvas_w + 0.02 or x0 < -0.02:
                    rep.findings.append(Finding(
                        "HIGH", "overflow:off-canvas",
                        f"строка {w:.2f}\" (wrap=off, {align}) рисуется в "
                        f"[{x0:.2f}..{x1:.2f}]\", выходит за холст 0..{canvas_w:.2f}\"",
                        si, label))
                if extents is not None:
                    extents[sh.shape_id] = (x0, inner_t, x1, inner_b)
            else:
                # горизонталь: одиночный токен шире зоны. normAutofit ужимает кегль
                # под бокс, поэтому там это не клипается — не флагуем (ровно §12-кейс
                # «40%» возможен при noAutofit/sp, где кегль не ужимается).
                if af != "norm":
                    worst, worst_w = "", 0.0
                    for tok in text.replace("\n", " ").split(" "):
                        tw = fit.text_width_in(tok, font, size)
                        if tw > worst_w:
                            worst_w, worst = tw, tok
                    if worst_w > inner_w + 0.02:
                        rep.findings.append(Finding(
                            "HIGH", "overflow:token-wider-than-box",
                            f"токен {worst!r} = {worst_w:.2f}\" шире зоны {inner_w:.2f}\" "
                            f"→ перенос/обрезка внутри слова", si, label))
                # вертикаль: нарисованная высота многострочного текста. Когда бокс
                # НЕ автоподгоняется (sp растёт под текст, norm ужимает кегль — оба
                # не клипают), текст ниже бокса = обрезка/вылет. Тон. допуск на лидинг.
                nlines = fit.line_count(text, font, size, inner_w)
                draw_t, draw_b = inner_t, inner_b
                if af in ("no", "none") and nlines >= 2:
                    h = fit.text_height_in(text, font, size, inner_w, line_spacing=linespc)
                    # куда рисуется блок высотой h внутри/за боксом — по якорю; экстент
                    # РАСШИРЯЕМ (max/min), но не ужимаем ниже бокса (бокс = минимум).
                    anchor = _v_anchor(txBody)
                    if anchor == "ctr":
                        cy = (inner_t + inner_b) / 2
                        draw_t, draw_b = min(inner_t, cy - h / 2), max(inner_b, cy + h / 2)
                    elif anchor == "b":
                        draw_t = min(inner_t, inner_b - h)
                    else:
                        draw_b = max(inner_b, inner_t + h)
                    if h > inner_h * 1.12 + 0.02:
                        rep.findings.append(Finding(
                            "MEDIUM", "overflow:vertical",
                            f"{nlines} строк ≈ {h:.2f}\" (lnSpc={linespc:.2f}) не влезают "
                            f"в высоту зоны {inner_h:.2f}\" → обрезка/вылет", si, label))
                if extents is not None:
                    # нарисованный экстент (низ может уйти за бокс при переполнении) —
                    # парные проверки видят «выросший» 2-строчный заголовок в чужой зоне
                    extents[sh.shape_id] = (inner_l, draw_t, inner_r, draw_b)
        except fit.FontNotFound as e:
            rep.findings.append(Finding(
                "LOW", "unmeasured", f"нет файла шрифта {e} — замер пропущен", si, label))


def lint_empty_lines(slide, si, rep):
    """Трейлинг <a:br/> и пустые абзацы (A12). Чистый XML-скан: paragraph.text
    не показывает <a:br/>, поэтому идём по элементам."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        if not tf.text.strip():
            continue
        paras = tf._txBody.findall(qn("a:p"))
        label = f"{tf.text.strip()[:22]!r}"
        # пустые абзацы (нет видимого текста), кроме случая единственного абзаца
        nonempty = [p for p in paras
                    if fit.visible_glyph_run("".join(
                        t.text or "" for t in p.findall(f"{qn('a:r')}/{qn('a:t')}")))]
        empty_count = len(paras) - len(nonempty)
        if empty_count and len(paras) > 1:
            rep.findings.append(Finding(
                "MEDIUM", "empty-lines:empty-paragraph",
                f"{empty_count} пустой абзац(ев) в боксе — в центрированном боксе "
                f"уводит текст; стрипать (A12)", si, label))
        # трейлинг <a:br/> после последнего непустого run в абзаце
        for p in paras:
            kids = list(p)
            last_text = -1
            for i, ch in enumerate(kids):
                if ch.tag == qn("a:r"):
                    txt = "".join(t.text or "" for t in ch.findall(qn("a:t")))
                    if txt.strip():
                        last_text = i
            trailing_br = sum(1 for ch in kids[last_text + 1:] if ch.tag == qn("a:br"))
            if trailing_br:
                rep.findings.append(Finding(
                    "MEDIUM", "empty-lines:trailing-br",
                    f"{trailing_br} трейлинг <a:br/> в абзаце — стрипать (A12)", si, label))


def _bbox(sh, extents=None):
    ext = (extents or {}).get(sh.shape_id)
    if ext:
        return ext
    return (sh.left / EMU, sh.top / EMU,
            (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)


def _overlap_area(a, b):
    ox = max(0, min(a[2], b[2]) - max(a[0], b[0]))
    oy = max(0, min(a[3], b[3]) - max(a[1], b[1]))
    return ox * oy


_PLACEHOLDER_RE = None
# универсальные (не пер-шаблонные) маркеры черновика — общие для любого шаблона
_PLACEHOLDER_UNIVERSAL = [r"lorem ipsum", r"click to add",
                          r"\btodo\b", r"\btbd\b", r"\bplaceholder\b"]


def _placeholder_patterns() -> list[str]:
    """Пер-шаблонные тексты-заглушки (placeholder_texts) из brand-tokens ВСЕХ
    шаблонов + универсальные маркеры. Union по шаблонам: линт бежит по готовому
    pptx без tdir, лишние паттерны ложно не срабатывают (реальный контент не
    пишет «заголовок слайда»). Пер-шаблонные — впереди (порядок как до Г.4.3)."""
    import glob as _glob
    import json as _json
    pats: list[str] = []
    root = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "layouts")
    for bt in sorted(_glob.glob(os.path.join(root, "*", "brand-tokens.json"))):
        try:
            with open(bt, encoding="utf-8") as f:
                pats += _json.load(f).get("placeholder_texts", [])
        except (OSError, ValueError):
            pass
    return pats + _PLACEHOLDER_UNIVERSAL


def lint_placeholder_leaks(slide, si, rep):
    """Остаточные тексты-заглушки шаблона/черновика в финальном деке (идея —
    у аудита community/presentation; здесь адаптирована под русские заглушки
    Еды). Точный пер-слотовый контроль делает build_deck (предупреждение «слот
    не заполнен»); этот линт — общая сеть на GENERATE-пути и ручные правки."""
    global _PLACEHOLDER_RE
    if _PLACEHOLDER_RE is None:
        import re
        _PLACEHOLDER_RE = re.compile("|".join(_placeholder_patterns()), re.IGNORECASE)
    for sh in slide.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        m = _PLACEHOLDER_RE.search(sh.text_frame.text)
        if m:
            rep.findings.append(Finding(
                "MEDIUM", "placeholder-leak",
                f"похоже на заглушку ({m.group(0)!r}) — демо-тексту шаблона "
                f"в финальном деке не место: заполни слот или очисти", si,
                f"{sh.text_frame.text.strip()[:22]!r}"))


def _card_shapes(slide):
    """Плашки-контейнеры слайда: автофигуры roundRect/rect с заливкой (solidFill
    или тема через p:style/fillRef) и без собственного текста. Картинки НЕ
    контейнеры (текст на фото — штатная слоистость: обложка, hero); микро-плашки
    < 1 in² (иконки, номерные бейджи) — тоже нет."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    cards = []
    for sh in slide.shapes:
        if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            continue
        try:
            if abs(sh.rotation) > 0.5:      # повёрнутая фигура — декор-оверлей (напр. градиент-
                continue                     # полоса closing Городских), НЕ плашка-контейнер: её
        except Exception:                    # осепараллельный bbox не отражает границ → ложный B10
            pass
        try:
            spPr = sh._element.spPr
            pg = spPr.find(qn("a:prstGeom"))
            if pg is None or pg.get("prst") not in ("roundRect", "rect"):
                continue
            has_fill = (spPr.find(qn("a:solidFill")) is not None
                        or sh._element.find(f"{qn('p:style')}/{qn('a:fillRef')}")
                        is not None)
            if not has_fill:
                continue
            b = (sh.left / EMU, sh.top / EMU,
                 (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)
        except Exception:
            continue
        if (b[2] - b[0]) * (b[3] - b[1]) < 1.0:
            continue
        cards.append(b)
    return cards


def lint_containment(slide, si, rep, extents=None, eps=0.06):
    """Машинный слой чеклиста B10: нарисованный экстент текста (для wrap=off —
    фактические глифы с учётом выравнивания, из _draw_extent) лежит внутри
    своей плашки ПО ГОРИЗОНТАЛИ. «Своей» = центр текста внутри плашки и ≥60%
    площади на ней — текст рядом/выше плашки (заголовки, подписи над чартами)
    ничьим не считается и не проверяется. Вылет > eps — HIGH: лечение по
    контентному тиру (value → shrink группой, проза → reword), геометрию
    плашки не трогать (quality-gates.md).

    Вертикаль сознательно НЕ меряем: у клонов дизайнерских слайдов текст-боксы
    легитимно длиннее компактных генеративных карточек (нижний паддинг бокса —
    не глифы), а честная глиф-вертикаль требует anchor-математики. Мотивирующий
    кейс B10 (значение длиннее заглушки → вылет ЛЕВЕЕ/ПРАВЕЕ карточки) —
    горизонтальный; вертикаль карточек стерегут verify_tiles/composition-чеки
    и глазной B10."""
    cards = _card_shapes(slide)
    if not cards:
        return
    for sh in _measurable_shapes(slide):
        if sh.text_frame._txBody.find(f".//{qn('a:fld')}") is not None:
            continue                       # авто-поля (номер слайда) не меряем
        try:                               # заголовок B10 не чек по замыслу
            if sh.is_placeholder and sh.placeholder_format.idx == 0:
                continue                   # (он «ничей»); на custom-слайдах он
        except (AttributeError, KeyError):  # легитимно перекрывает пилюли .mark,
            pass                           # иначе пилюля ложно становится «домом»
        try:
            t = _bbox(sh, extents)         # нарисованный экстент из overflow-пасса
        except Exception:
            continue
        # «своей» = якорная точка текста в плашке по X и ≥60% ВЫСОТЫ текста на
        # плашке. Якорные точки: центр ИЛИ левая/правая кромка (0.32.2: при
        # ЭКСТРЕМАЛЬНОМ вылете центр покидает плашку и текст становился «ничьим»
        # — дыра того же рода, что и до 2026-07-07, когда вылет >40% ширины
        # «освобождал» значение от чека; кромка старта текста остаётся в плашке
        # и держит его домом). Вертикаль — по центру, как раньше.
        cx, cy = (t[0] + t[2]) / 2, (t[1] + t[3]) / 2
        th = max(t[3] - t[1], 1e-6)
        xa = (cx, t[0] + 0.05, t[2] - 0.05)     # центр, левая и правая кромки
        home, best_ov = None, 0.0
        for b in cards:
            if not (b[1] <= cy <= b[3] and any(b[0] <= x <= b[2] for x in xa)):
                continue
            ov = (min(t[3], b[3]) - max(t[1], b[1])) / th
            if ov >= 0.6 and ov > best_ov:
                home, best_ov = b, ov
        if home is None:
            continue
        spills = {"левый": home[0] - t[0], "правый": t[2] - home[2]}
        bad = {k: v for k, v in spills.items() if v > eps}
        if bad:
            side, v = max(bad.items(), key=lambda kv: kv[1])
            label = f"{sh.text_frame.text.strip()[:22]!r}"
            rep.findings.append(Finding(
                "HIGH", "containment:text-outside-card",
                f"текст вылезает за {side} край своей плашки на {v:.2f}\" "
                f"(B10; value → shrink группой до общего кегля, проза → reword; "
                f"геометрию плашки не трогать)", si, label))


def lint_collisions(slide, si, canvas_w, canvas_h, rep, extents=None,
                    min_overlap_in2=0.25):
    """A9: пересечение ТЕКСТ-vs-ТЕКСТ. Текст поверх карточки/картинки — штатная
    слоистость (не коллизия), поэтому пары с не-текстом пропускаем: ловим именно
    нечитаемое наложение надписей. Экстент текста берём с учётом overflow
    (_draw_extent), так что разъехавшееся wrap=off-значение поймается тут."""
    items = []
    for sh in slide.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        try:
            x0, y0, x1, y1 = _bbox(sh, extents)
        except Exception:
            continue
        if x1 - x0 <= 0 or y1 - y0 <= 0:
            continue
        items.append((sh, (x0, y0, x1, y1)))
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            shi, bi = items[i]
            shj, bj = items[j]
            area = _overlap_area(bi, bj)
            if area <= min_overlap_in2:
                continue
            li = shi.text_frame.text.strip()[:16]
            lj = shj.text_frame.text.strip()[:16]
            rep.findings.append(Finding(
                "MEDIUM", "collision:text-text",
                f"наложение надписей ≈{area:.2f} in²: [{li!r}] × [{lj!r}]", si))


def lint_pills(slide, si, rep, extents=None, eps=0.06, foreign_overlap_in2=0.05):
    """Пилюли — МЕЛКИЕ залитые roundRect/rect (0.02–1.0 in²) без собственного
    текста. Контейнер-чек B10 их не видит (фильтр «микро-плашки < 1 in²»), а
    B13 требует накрытия ≥60% площади — частичный наезд молчал. Мотив — дек
    drive-long s11 (user-фидбек 2026-08-07): пилюли-статусы легли на хвосты
    заголовков карточек, а их лейблы вылезли из пилюль белым по белому. Два
    дефекта, оба HIGH:
      (а) свой текст (центр в пилюле) ШИРЕ пилюли → обрезка/невидимые буквы;
      (б) пилюля и ЧУЖОЙ текст наезжают друг на друга ≥ foreign_overlap_in2
          (независимо от z-порядка: сверху — прячет буквы, снизу — грязный слой).
    Пилюли с собственным текстом (тег-доноры Хурала) меряются штатными чеками
    своего бокса и сюда не попадают."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pills = []
    for sh in slide.shapes:
        if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            continue
        try:
            if abs(sh.rotation) > 0.5:
                continue
            spPr = sh._element.spPr
            pg = spPr.find(qn("a:prstGeom"))
            if pg is None or pg.get("prst") not in ("roundRect", "rect"):
                continue
            has_fill = (spPr.find(qn("a:solidFill")) is not None
                        or sh._element.find(f"{qn('p:style')}/{qn('a:fillRef')}") is not None)
            if not has_fill:
                continue
            b = (sh.left / EMU, sh.top / EMU,
                 (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)
        except Exception:
            continue
        if not (0.02 <= (b[2] - b[0]) * (b[3] - b[1]) <= 1.6):
            continue                     # верхний порог 1.6: пилюля «Maintain»
        pills.append(b)                  # drive-long = 1.08 in² (0.32.2)
    if not pills:
        return
    for sh in _measurable_shapes(slide):
        if sh.text_frame._txBody.find(f".//{qn('a:fld')}") is not None:
            continue
        try:
            t = _bbox(sh, extents)
        except Exception:
            continue
        cx, cy = (t[0] + t[2]) / 2, (t[1] + t[3]) / 2
        label = f"{sh.text_frame.text.strip()[:22]!r}"
        for b in pills:
            if b[0] <= cx <= b[2] and b[1] <= cy <= b[3]:      # (а) свой лейбл
                spill = max(b[0] - t[0], t[2] - b[2])
                if spill > eps:
                    rep.findings.append(Finding(
                        "HIGH", "containment:text-outside-pill",
                        f"текст шире своей пилюли на {spill:.2f}\" — обрезка/"
                        f"невидимые буквы; пилюлю шире (паддинг) или лейбл короче",
                        si, label))
                    break
            elif _overlap_area(t, b) >= foreign_overlap_in2:   # (б) чужой текст
                rep.findings.append(Finding(
                    "HIGH", "collision:pill-over-text",
                    f"пилюля и надпись наезжают друг на друга "
                    f"(≈{_overlap_area(t, b):.2f} in²) — буквы прячутся под "
                    f"плашкой; разнести (пилюля в поток, не поверх текста)",
                    si, label))
                break


# --------------------------------------------------------------------- API


def lint_hidden_under_plate(slide, si, rep, extents=None, cover=0.6):
    """B13: надпись, скрытая ЗАЛИТОЙ плашкой поверх неё (по z-order документа).
    Плашка стоит ПОЗЖЕ текста в порядке отрисовки и накрывает ≥cover его площади
    → текст физически не читается. Мотив — hero: жёлтая плашка широкого значения
    легла на подпись-контекст, а gate молчал (фидбек медиакит РБ, 2026-07).
    Фон-карточки рисуются ПОД своим текстом (меньший z) и сюда не попадают —
    поэтому штатная слоистость «текст на карточке» не флагуется."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    order = list(slide.shapes)
    zi = {sh.shape_id: i for i, sh in enumerate(order)}   # shape_id стабилен между пассами
    plates = []
    for sh in order:
        if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            continue
        try:
            spPr = sh._element.spPr
            pg = spPr.find(qn("a:prstGeom"))
            if pg is None or pg.get("prst") not in ("roundRect", "rect"):
                continue
            has_fill = (spPr.find(qn("a:solidFill")) is not None
                        or sh._element.find(f"{qn('p:style')}/{qn('a:fillRef')}") is not None)
            if not has_fill:
                continue
            b = (sh.left / EMU, sh.top / EMU,
                 (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)
        except Exception:
            continue
        if (b[2] - b[0]) * (b[3] - b[1]) < 1.0:          # микро-плашки не прячут
            continue
        plates.append((zi[sh.shape_id], b))
    if not plates:
        return
    for sh in _measurable_shapes(slide):
        if sh.text_frame._txBody.find(f".//{qn('a:fld')}") is not None:
            continue
        try:
            t = _bbox(sh, extents)
        except Exception:
            continue
        area = max((t[2] - t[0]) * (t[3] - t[1]), 1e-6)
        tz = zi[sh.shape_id]
        for pz, b in plates:
            if pz <= tz:                                 # плашка под текстом = фон
                continue
            if _overlap_area(t, b) / area >= cover:
                label = f"{sh.text_frame.text.strip()[:22]!r}"
                rep.findings.append(Finding(
                    "HIGH", "collision:text-under-plate",
                    f"надпись накрыта залитой плашкой поверх неё "
                    f"(≥{int(cover * 100)}% площади) — текст не виден (B13)", si, label))
                break


def lint_widows(slide, si, prs, theme, default_style, rep):
    """B13/widow (user-правило 2026-07): одно слово на последней строке заголовка
    недопустимо — к нему спускают ещё одно (nbsp между двумя последними словами).
    Меряем фактический перенос заголовочных боксов (кегль ≥ 28pt — heading/
    subheading/caption/statement) и ловим 1-словный хвост при ≥2 строках.
    Тело/буллеты не трогаем: там висячее слово нормально."""
    for sh in _measurable_shapes(slide):
        txBody = sh.text_frame._txBody
        if txBody.find(f".//{qn('a:fld')}") is not None:
            continue
        wrap, ins = _bodypr(txBody)
        if wrap == "none":                     # значения/статы не переносятся
            continue
        para = sh.text_frame.paragraphs[0]
        if len(sh.text_frame.paragraphs) > 1:  # многоабзацный блок = тело, не заголовок
            continue
        run = para.runs[0] if para.runs else None
        font, size = effective_props(sh, para, run, prs, slide, theme, default_style)
        if not font or not size or size < 28:  # только заголовочный кегль
            continue
        text = sh.text_frame.text.strip()
        if len(text.split()) < 3:              # 1-2 слова хвостом не висят
            continue
        inner_w = max(sh.width / EMU - (ins["lIns"] + ins["rIns"]) / EMU, 0.01)
        try:
            lines = fit.wrap_lines(text, font, size, inner_w)
        except fit.FontNotFound:
            continue
        if len(lines) >= 2 and len(lines[-1].split()) == 1 and NBSP not in lines[-1]:
            rep.findings.append(Finding(
                "MEDIUM", "widow:single-word-last-line",
                f"висячее слово {lines[-1]!r} одно на последней строке заголовка "
                f"— склеить nbsp с предыдущим или перефразировать (B13)", si,
                f"{text[:22]!r}"))


def _insert_last_nbsp(para) -> bool:
    """Заменить последний обычный пробел в ранах абзаца на nbsp (склейка двух
    последних слов). Работает и при разбиении на раны (инлайн-жирный): ищем
    пробел с конца по ранам."""
    for r in reversed(para.runs):
        t = r.text or ""
        idx = t.rstrip().rfind(" ")
        if idx != -1:
            r.text = t[:idx] + NBSP + t[idx + 1:]
            return True
    return False


def defuse_widows(prs) -> int:
    """Пост-пасс сборки: убрать висячее слово в заголовках — вставить nbsp между
    двумя последними словами, если это снимает 1-словный хвост без роста числа
    строк (user-правило 2026-07: «к висяку спускают ещё одно слово»). Меряем тем
    же fit, что и lint, по фактической геометрии собранного бокса; покрывает все
    пути сразу (FILL, GENERATE-плитки, обложка). Неисправимое склейкой (два
    последних слова шире строки) остаётся и флагуется lint widow → нужен реворд.
    Возвращает число исправленных боксов."""
    _theme_cache.clear()
    default_style = _default_text_style(prs)
    fixed = 0
    for slide in prs.slides:
        theme = _slide_theme_fonts(slide)
        for sh in _measurable_shapes(slide):
            txBody = sh.text_frame._txBody
            if txBody.find(f".//{qn('a:fld')}") is not None:
                continue
            wrap, ins = _bodypr(txBody)
            if wrap == "none":
                continue
            paras = sh.text_frame.paragraphs
            if len(paras) != 1:                # многоабзацный блок = тело, не заголовок
                continue
            para = paras[0]
            run = para.runs[0] if para.runs else None
            font, size = effective_props(sh, para, run, prs, slide, theme, default_style)
            if not font or not size or size < 28:
                continue
            text = sh.text_frame.text.strip()
            if len(text.split()) < 3:
                continue
            inner_w = max(sh.width / EMU - (ins["lIns"] + ins["rIns"]) / EMU, 0.01)
            try:
                if fit.widow_nbsp(text, font, size, inner_w) != text and _insert_last_nbsp(para):
                    fixed += 1
            except fit.FontNotFound:
                continue
    return fixed


def lint_table_alignment(slide, si, rep):
    """Правило вёрстки таблиц (user 2026-07-20): текстовая колонка — влево,
    колонка из чисел — вправо (заголовок — как его колонка). Смотрим data-строки
    (r ≥ 1); прочерки/пустые нейтральны; ловит все ветки разом — fill, custom
    (нативный a:tbl), перевёрстку. MEDIUM: типы — эвристика, решает глаз."""
    from pptx.enum.text import PP_ALIGN
    for sh in slide.shapes:
        if not getattr(sh, "has_table", False):
            continue
        tbl = sh.table
        nr, nc = len(tbl.rows), len(tbl.columns)
        if nr < 2:
            continue
        for c in range(nc):
            texts, aligns = [], set()
            for r in range(1, nr):
                try:
                    cell = tbl.cell(r, c)
                except Exception:        # span-зоны мержей пропускаем
                    continue
                t = (cell.text_frame.text or "").strip()
                if not t or t in fit.NEUTRAL_CELLS:
                    continue
                texts.append(t)
                for p in cell.text_frame.paragraphs:
                    aligns.add(p.alignment)
            if not texts:
                continue                 # пустая/декоративная колонка донора
            numeric = all(fit.is_numeric_cell(t) for t in texts)
            if numeric and PP_ALIGN.RIGHT not in aligns:
                rep.findings.append(Finding(
                    "MEDIUM", "table-align:numeric-column-not-right",
                    f"колонка {c + 1}: все значения числовые "
                    f"({texts[0]!r}…) — по правилу вёрстки выравнивать вправо",
                    si, sh.name))
            elif not numeric and PP_ALIGN.RIGHT in aligns:
                rep.findings.append(Finding(
                    "MEDIUM", "table-align:text-column-right",
                    f"колонка {c + 1}: текстовые значения ({texts[0]!r}…) "
                    f"выровнены вправо — текст в таблице выравнивается влево",
                    si, sh.name))


def run_lints(pptx_path: str) -> LintReport:
    rep = LintReport(file=pptx_path)
    _theme_cache.clear()
    prs = Presentation(pptx_path)
    canvas_w = prs.slide_width / EMU
    canvas_h = prs.slide_height / EMU
    default_style = _default_text_style(prs)
    rep.stats["canvas_in"] = f"{canvas_w:.3f}x{canvas_h:.3f}"
    rep.stats["calibration"] = dict(fit.load_calibration())
    si = 0
    themes_seen = set()
    for si, slide in enumerate(prs.slides, 1):
        theme = _slide_theme_fonts(slide)
        themes_seen.add(theme)
        extents = {}   # нарисованные экстенты wrap=off по shape_id (см. lint_overflow)
        lint_overflow(slide, si, prs, theme, default_style, canvas_w, canvas_h,
                      rep, extents)
        lint_containment(slide, si, rep, extents)
        lint_placeholder_leaks(slide, si, rep)
        lint_empty_lines(slide, si, rep)
        lint_collisions(slide, si, canvas_w, canvas_h, rep, extents)
        lint_hidden_under_plate(slide, si, rep, extents)
        lint_pills(slide, si, rep, extents)
        lint_widows(slide, si, prs, theme, default_style, rep)
        lint_table_alignment(slide, si, rep)
    rep.stats["slides"] = si
    rep.stats["theme_fonts"] = [{"major": t[0], "minor": t[1]} for t in themes_seen]
    return rep


def render_text(rep: LintReport) -> str:
    out = [f"LINT: {rep.file}", "-" * 64]
    for k, v in rep.stats.items():
        out.append(f"  {k}: {v}")
    out.append("")
    if not rep.findings:
        out.append("✅ нет измерительных проблем.")
    else:
        out.append(f"{rep.high} HIGH / {rep.medium} MEDIUM / {rep.low} LOW:")
        for f in sorted(rep.findings, key=lambda x: {"HIGH": 0, "MEDIUM": 1, "LOW": 2}[x.severity]):
            loc = f" s{f.slide}" + (f" [{f.shape}]" if f.shape else "")
            out.append(f"  [{f.severity}] {f.code}{loc}: {f.message}")
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    rep = run_lints(a.pptx)
    if a.json:
        print(json.dumps({
            "file": rep.file, "stats": rep.stats,
            "findings": [vars(f) for f in rep.findings],
            "summary": {"high": rep.high, "medium": rep.medium, "low": rep.low},
        }, indent=2, ensure_ascii=False))
    else:
        print(render_text(rep))
    return 1 if rep.high else 0


if __name__ == "__main__":
    raise SystemExit(main())
