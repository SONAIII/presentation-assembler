#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fill.py — заполнение готовых слайдов шаблона текстом (ветка FILL роутера §3).

Вёрстку/геометрию НЕ трогаем — только текст по месту, плюс детерминированные
правила §12, которые иначе ловятся «на глаз» и пропускаются:
  * слот значения/стата → wrap=off (число не переносится посимвольно), кегль
    designer'а сохраняем, при необходимости shrink-to-fit как fallback;
  * пустые строки (трейлинг <a:br/>, пустые абзацы) стрипаем.

Все замеры — через lib.fit (тот же измеритель, что у lint и рендерера).
"""
from __future__ import annotations

import copy
import os
import re
import sys

from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fit  # noqa: E402
import template_profile as TP  # noqa: E402  (bold-карта — из custom-profile.json)


_MD_BOLD = re.compile(r"\*\*(.+?)\*\*")
# Потолок роста кегля при stretch table-слота БЕЗ абсолютных потолков в спеке:
# ×4/3 от дизайнерского кегля рана. Шаблон-агностичный дефолт; 4/3 = одобренный
# владельцем потолок Городских (body 24pt → 32pt, слайд-образец 2026-07-29).
STRETCH_MAX_FACTOR = 4 / 3
# «Жирный» бренда = ФАЙЛОВОЕ семейство стиля, не синтетика b=1 поверх Regular
# (user-фидбек 2026-07-16: YS-стили — отдельные шрифты с полными именами).
# Раны БЕЗ явного rFonts (наследование от темы: None / +mj-lt / +mn-lt) резолвятся
# в фактические семейства темы шаблона (user-фидбек Хурал 2026-08-07: акцент на
# таких ранах давал синтетику, потому что имя не находилось в bold-карте);
# build_deck снимает major/minor latin с темы мастера и кладёт сюда set_theme_fonts.
# Карта имя-стиля → bold-семейство живёт в custom-profile.json (fonts.bold_family),
# ОДИН источник со шрифтами кастома; tdir — ambient (TP.active(), ставится
# build_deck.main() один раз на дек; дефолт 'eda').

_THEME_FONTS = {"major": None, "minor": None}


def set_theme_fonts(major, minor):
    """major/minor latin темы активного шаблона (ставит build_deck после открытия prs)."""
    _THEME_FONTS["major"], _THEME_FONTS["minor"] = major, minor


def _effective_font_name(name):
    """Явное имя рана как есть; наследуемое (None) и темовые +mj-lt/+mn-lt —
    в фактическое семейство темы (для bold-карты). Без темы — как было."""
    if name == "+mj-lt":
        return _THEME_FONTS["major"] or name
    if not name or name == "+mn-lt":
        return _THEME_FONTS["minor"] or name
    return name


def _fill_tf(tf, lines):
    """Заполнить text_frame построчно, сохраняя формат первого run'а каждого абзаца.
    Строки СВЕРХ дизайнерских абзацев наследуют формат последнего абзаца (клон
    его <a:p> с pPr и rPr) — голый add_paragraph терял типографику, из-за чего
    несколько абзацев desc приходилось склеивать в один (кейс плиток 2026-07-15).
    Работает и для шейпа, и для ячейки таблицы (у обоих есть text_frame).
    Инлайн-акцент (v0.16.0): `**фрагмент**` в строке → отдельный ран жирным
    ФАЙЛОВЫМ семейством (YS Text Regular → YS Text Bold, карта fonts.bold_family
    в custom-profile.json); синтетический b=1 — только fallback вне карты."""
    def _set_line(p, line):
        segs = _MD_BOLD.split(line)
        if len(segs) > 1:
            _set_line_md(p, segs)
            return
        if p.runs:
            p.runs[0].text = line
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
        else:
            p.add_run().text = line

    def _set_line_md(p, segs):
        pieces = [(seg, i % 2 == 1) for i, seg in enumerate(segs) if seg != ""]
        if not p.runs:
            p.add_run()
        base = copy.deepcopy(p.runs[0]._r)
        for r in list(p.runs[1:]):
            r._r.getparent().remove(r._r)
        anchor = p.runs[0]._r
        for _ in range(len(pieces) - 1):
            el = copy.deepcopy(base)
            anchor.addnext(el)
            anchor = el
        for (seg, bold), run in zip(pieces, p.runs):
            run.text = seg
            if not bold:
                continue
            name = _effective_font_name(run.font.name)
            fam = TP.bold_family(TP.active()).get(name or "")
            if fam:
                run.font.name = fam
            elif not name or not re.search(r"Bold|Heavy|Black", name):
                run.font.bold = True      # вне карты (и темы) — честный fallback

    # шаблон для строк сверх дизайнерских абзацев — ПРИСТИННЫЙ последний абзац
    # (снятый до правок: md-жирный ран текущей строки не должен наследоваться
    # следующими абзацами)
    pristine_last = copy.deepcopy(tf.paragraphs[-1]._p)
    for k, line in enumerate(lines):
        if k < len(tf.paragraphs):
            _set_line(tf.paragraphs[k], line)
        else:
            tf.paragraphs[-1]._p.addnext(copy.deepcopy(pristine_last))
            _set_line(tf.paragraphs[-1], line)
    for extra in list(tf.paragraphs[len(lines):]):
        extra._p.getparent().remove(extra._p)


def fill_lines(shape, lines):
    """Заполнить текст-бокс построчно, сохраняя формат первого run'а абзаца.
    Годится и для одной строки, и для «лид-ин + тело»."""
    _fill_tf(shape.text_frame, lines)


def strip_empty_lines(tf):
    """A12: убрать трейлинг <a:br/> и пустые абзацы — в центрированном боксе
    пустая строка уводит текст вверх. Намеренные переносы между непустыми
    строками («спасибо\\nза внимание») сохраняем."""
    for par in list(tf.paragraphs):
        p = par._p
        kids = list(p)
        last_text = -1
        for idx, ch in enumerate(kids):
            if ch.tag == qn("a:r"):
                txt = "".join((t.text or "") for t in ch.findall(qn("a:t")))
                if txt.strip():
                    last_text = idx
        for ch in kids[last_text + 1:]:
            if ch.tag in (qn("a:br"), qn("a:r")):
                p.remove(ch)
    for par in list(tf.paragraphs):
        if not par.text.strip() and len(tf.paragraphs) > 1:
            par._p.getparent().remove(par._p)


def _run_font_size(shape):
    """(font_name, size_pt) первого run'а — для замера. Берём то, что реально
    уйдёт в рендер (run override важнее наследования)."""
    p0 = shape.text_frame.paragraphs[0]
    run = p0.runs[0] if p0.runs else None
    name = run.font.name if run else None
    size = run.font.size.pt if run and run.font.size else None
    # наследование шрифта из lstStyle, если на run'е нет
    if name is None:
        ls = shape.text_frame._txBody.find(qn("a:lstStyle"))
        if ls is not None:
            lat = ls.find(f"{qn('a:lvl1pPr')}/{qn('a:defRPr')}/{qn('a:latin')}")
            if lat is not None:
                name = lat.get("typeface")
    return name, size


def set_value_slot(shape, text, *, max_width_in=None, min_pt=24.0, force_pt=None):
    """Слот значения/стата (§12): wrap=off, кегль designer'а сохраняем. Если
    задан max_width_in и значение шире — shrink-to-fit как fallback (не даём
    числу уехать в соседа). force_pt — ЕДИНЫЙ кегль группы однородных значений
    (правило user 2026-07-02: ужали одно — ужимаем все; см. value_group_pt).
    strip пустых строк включён."""
    fill_lines(shape, [text])
    shape.text_frame.word_wrap = False
    if force_pt is not None:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(force_pt)
        strip_empty_lines(shape.text_frame)
        return
    name, size = _run_font_size(shape)
    if max_width_in and name and size:
        try:
            w = fit.text_width_in(text, name, size)
            if w > max_width_in:
                new = fit.shrink_to_fit(text, name, size, max_width_in,
                                        wrap=False, min_size_pt=min_pt)
                run = shape.text_frame.paragraphs[0].runs[0]
                run.font.size = Pt(new)
        except fit.FontNotFound:
            pass
    strip_empty_lines(shape.text_frame)


def value_group_pt(shapes, values, slot) -> float | None:
    """ЕДИНЫЙ кегль для группы однородных значений (метрики одного ранга всегда
    одного размера — user 2026-07-02): считаем shrink-кегль каждого значения в
    его боксе и берём МИНИМУМ. None = мерить нечем (нет шрифта/ширины).
    Шрифт/кегль без явного рана добираем из спеки (slot.font / slot.size_pt):
    у дизайнерских плашек #11 кегль 2-й/3-й карточки живёт в lstStyle, не на
    ране — до фикса 2026-07-07 такие значения МОЛЧА выпадали из группы, и
    самое длинное («6 городов») не влияло на общий кегль."""
    bw = slot.get("box_w_in")
    sizes = []
    for sh, v in zip(shapes, values):
        name, size = _run_font_size(sh)
        name = name or slot.get("font")
        size = size or slot.get("size_pt")
        if not size:
            continue
        if bw and name:
            try:
                sizes.append(fit.shrink_to_fit(str(v), name, size, bw, wrap=False,
                                               min_size_pt=slot.get("min_pt", 24.0)))
                continue
            except fit.FontNotFound:
                pass
        sizes.append(size)
    return min(sizes) if sizes else None


def value_max_width(vslot: dict, vfit: dict) -> float:
    """Потолок роста значения (right-align, wrap=off — глифы растут влево):
    от правого края дизайнерского бокса до левого края плашки + паддинг
    pad_left_in (default 0.2"). Прежний люфт «бокс + 1.0"» разрешал вылет
    на дюйм ЗА плашку — отменён user-фидбеком 2026-07-07 («~35 мин» уезжало
    на фото)."""
    pad = vfit.get("pad_left_in", 0.2)
    return vslot["rel_left_in"] + vslot["width_in"] - pad


def value_overflows(shape, text, max_width_in: float, eps=0.02) -> bool:
    """Пост-верификация значения ПОСЛЕ применения кегля (в т.ч. группового
    force_pt): глифы шире потолка даже на флоре → честный overflow-флаг
    (REWORD), молча не пропускаем. Не меряется — False (подстрахует линт)."""
    name, size = _run_font_size(shape)
    if not (name and size):
        return False
    try:
        return fit.text_width_in(str(text), name, size) > max_width_in + eps
    except fit.FontNotFound:
        return False


# ====================================================================
# Структурный тир (обложка/раздел/содержание/закрытие): FILL по fill-map.
# Готовый дизайнерский слайд заполняем по тексту-заглушке; вёрстку/фон/декор не
# трогаем. Проза фитится как СТРАХОВКА — переформулировка (reword) это забота
# операционного контракта, здесь только защитная сетка от вылета за бокс.
# ====================================================================


def find_shape_by_text(slide, prefix, *, ci=False):
    """Найти шейп, чей текст (stripped) начинается с prefix. ci — без регистра.
    Стабильный ключ §5 (текст-заглушка), не индекс."""
    pref = prefix.lower() if ci else prefix
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        hay = t.lower() if ci else t
        if hay.startswith(pref):
            return sh
    return None


def find_shape_by_name(slide, name, *, recurse=False):
    """Первый текст-шейп с данным именем. Для контентных слайдов имя шейпа стабильнее
    и точнее lorem-префикса (когда имя на слайде уникально); повторяемые роли — через
    find_shapes_by_name. recurse — искать и внутри групп (event-блоки closing Хурала)."""
    def _scan(shapes):
        for sh in shapes:
            if sh.name == name and sh.has_text_frame:
                return sh
            if recurse and sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                got = _scan(sh.shapes)
                if got is not None:
                    return got
        return None
    return _scan(slide.shapes)


def clear_soft_breaks(tf):
    """Убрать ВСЕ <a:br/> во всех абзацах — для слотов с зашитыми в заглушке
    мягкими переносами (обложка «Название ⤶ в три строки»). Намеренные переносы
    приходят отдельными строками в fill_lines, не из заглушки. Сам <a:pPr>
    (с lnSpc дизайнера) не трогаем."""
    for par in tf.paragraphs:
        p = par._p
        for br in p.findall(qn("a:br")):
            p.remove(br)


def empty_text_frame(tf):
    """Очистить текст (оставить один пустой абзац, pPr сохранён): опциональный слот
    без контента — не показываем инструктивную заглушку и НЕ выдумываем текст."""
    first = tf.paragraphs[0]._p
    for ch in list(first):
        if ch.tag in (qn("a:r"), qn("a:br")):
            first.remove(ch)
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)


def _para_line_spacing(par, default=1.0):
    """Межстрочный коэффициент абзаца из <a:lnSpc><a:spcPct>. Дизайнерские title-боксы
    часто 80% — без этого высоту переоцениваем и ложно флагуем overflow."""
    pPr = par._p.find(qn("a:pPr"))
    if pPr is not None:
        ln = pPr.find(qn("a:lnSpc"))
        if ln is not None:
            pct = ln.find(qn("a:spcPct"))
            if pct is not None and pct.get("val"):
                return int(pct.get("val")) / 100000.0
    return default


def slot_overflows(shape, slot, *, default_line_spacing=1.0):
    """НЕ мутирует. Меряет, влезает ли контент в бокс на ДИЗАЙНЕРСКИХ кеглях,
    уважая кегль И межстрочный КАЖДОГО абзаца (у содержания par0=115pt активный,
    par1..n=72pt — иерархию сохраняем; title-боксы часто lnSpc 80%). Структурный
    тир = точь-в-точь: типографику дизайнера НЕ трогаем; переполнение = сигнал
    ПЕРЕФОРМУЛИРОВАТЬ (работа контракта) или флаг, а не смена шрифта. True при
    переполнении по высоте."""
    name = slot.get("font") or _run_font_size(shape)[0]
    bw, bh = slot.get("box_w_in"), slot.get("box_h_in")
    if not (name and bw and bh):
        return False
    tf = shape.text_frame
    total_h = 0.0
    first = True
    for par in tf.paragraphs:
        if not par.text.strip():
            continue
        r = par.runs[0] if par.runs else None
        # композитный слот (head_* — коробка «заголовок + тело» в два кегля):
        # первый абзац меряем шрифтом/кеглем заголовка, остальные — телом
        fname, default_sz = name, float(slot.get("size_pt") or 0)
        if first and slot.get("head_size_pt"):
            fname = slot.get("head_font") or name
            default_sz = float(slot["head_size_pt"])
        first = False
        sz = (r.font.size.pt if r and r.font.size else None) or default_sz
        if not sz:
            continue
        try:
            total_h += fit.text_height_in(par.text, fname, sz, bw,
                                          line_spacing=_para_line_spacing(par, default_line_spacing))
        except fit.FontNotFound:
            return False
    return total_h > bh + 0.05


def _locate_slot_shape(slide, slot):
    """Единая адресация одиночного слота: match_name(+occ) ИЛИ текст-заглушка."""
    if slot.get("match_name") and "occ" in slot:      # k-й из одноимённых по позиции (#23 mixed)
        shs = find_shapes_by_name(slide, slot["match_name"], sort=slot.get("sort", "doc"),
                                  recurse=slot.get("recurse", False), skip_regex=slot.get("skip_regex"))
        return shs[slot["occ"]] if slot["occ"] < len(shs) else None
    if slot.get("match_name"):
        return find_shape_by_name(slide, slot["match_name"],
                                  recurse=slot.get("recurse", False))
    return find_shape_by_text(slide, slot["find"], ci=slot.get("ci", False))


def _widen_centered(sh, w_in: float):
    """Спека-объявленное расширение бокса с сохранением центра — лечит
    template-inherent дефект (узкий дата-бокс таймлайна 1.42" уже 4-значного
    года 1.51", #28): тот же фикс, что генерация применяет своим датам
    (date_w_in). ТОЛЬКО для шейпов с собственным xfrm — placeholder без него
    от set-width вырождается в (0,0) (урок 2026-07-02)."""
    from pptx.oxml.ns import qn as _qn
    if sh._element.spPr.find(_qn("a:xfrm")) is None:
        return                      # placeholder с наследуемой геометрией — не трогаем
    new_w = Inches(w_in)
    if new_w <= sh.width:
        return
    sh.left = sh.left - Emu(int((new_w - sh.width) / 2))
    sh.width = new_w


def _reposition(sh, top_in, h_in):
    """Спека-объявленный сдвиг/высота бокса (set_top_in/set_h_in) — для
    user-санкционированных отступлений от дизайнерской геометрии (напр.
    комментарий #35 выравнивается по верху чарта — user-fix 2026-07-03).
    Только шейпы с собственным xfrm (см. _widen_centered)."""
    from pptx.oxml.ns import qn as _qn
    if sh._element.spPr.find(_qn("a:xfrm")) is None:
        return
    if top_in is not None:
        sh.top = Inches(float(top_in))
    if h_in is not None:
        sh.height = Inches(float(h_in))


def fill_text_slot(slide, slot, value):
    """Заполнить один структурный слот по fill-map — ТОЛЬКО текст. value — строка
    ИЛИ список строк (многострочный слот, напр. содержание: одна строка = один
    абзац). Типографику дизайнера (кегли, иерархию, выравнивание) НЕ меняем —
    структурный тир точь-в-точь. Гасит зашитые заглушечные переносы, чистит пустые
    строки. value пустой + clear_if_empty → очистить бокс (не оставлять заглушку).
    Переполнение на дизайнерских кеглях НЕ ужимаем, а флагуем (reword = контракт)."""
    sh = _locate_slot_shape(slide, slot)
    if sh is None:
        return {"found": False, "find": slot.get("match_name") or slot.get("find")}
    tf = sh.text_frame
    if slot.get("set_w_in"):
        _widen_centered(sh, float(slot["set_w_in"]))
    if slot.get("set_top_in") or slot.get("set_h_in"):
        _reposition(sh, slot.get("set_top_in"), slot.get("set_h_in"))
    if slot.get("clear_soft_breaks"):
        clear_soft_breaks(tf)
    if not value:
        if slot.get("clear_if_empty"):
            empty_text_frame(tf)
            return {"found": True, "cleared": True, "_shape": sh}
        return {"found": True, "skipped": True, "_shape": sh}
    if slot.get("fit_policy") == "value_shrink":       # число: wrap=off + shrink к боксу
        set_value_slot(sh, str(value), max_width_in=slot.get("box_w_in"),
                       min_pt=slot.get("min_pt", 24.0))
        return {"found": True, "filled": True, "value_shrink": True,
                "_shape": sh}
    lines = [str(x) for x in value] if isinstance(value, (list, tuple)) else [str(value)]
    single_thesis = len(lines) == 1  # ДО wrap: визуальные строки ≠ новые тезисы
    if slot.get("wrap_to_visible") and len(lines) == 1:
        lines = _wrap_to_visible(sh, slot, lines[0])
    fill_lines(sh, lines)            # только текст; формат абзацев дизайнера сохранён
    strip_empty_lines(tf)
    rep = {"found": True, "filled": True, "lines": len(lines), "_shape": sh}
    if single_thesis:
        # один тезис — не список (user 2026-07-29): буллет дизайнерского
        # спискового слота снимается, когда пункт остался единственным
        stripped = strip_bullets_shape(sh, only_bulleted=True)
        if stripped:
            rep["solo_bullet_stripped"] = stripped
    if slot.get("fit_policy") == "text_shrink" and slot.get("min_pt"):
        shrunk = _shrink_text_to_box(sh, slot)   # санкционированное спекой ужатие
        if shrunk:
            rep["text_shrink"] = shrunk
    if slot_overflows(sh, slot):
        rep["overflow"] = True       # сигнал переформулировать; шрифт НЕ трогаем
    return rep


def _shrink_text_to_box(sh, slot):
    """fit_policy «text_shrink» (v0.16.0, фидбек F медиакита): прозаический слот
    с САНКЦИОНИРОВАННЫМ спекой ужатием (объявленный min_pt — политика дизайнера,
    как у value-слотов) — весь слот жмётся ЕДИНЫМ кеглем от size_pt вниз до
    min_pt, пока текст не влезет в бокс по высоте; иерархия не ломается (кегль
    один на слот). Типовой пример: 2-строчный CTA контактов на 130pt.
    Возвращает применённый кегль или None (не ужимали)."""
    name = slot.get("font")
    bw, bh = slot.get("box_w_in"), slot.get("box_h_in")
    size, mn = float(slot.get("size_pt") or 0), float(slot["min_pt"])
    if not (name and bw and bh and size):
        return None
    tf = sh.text_frame

    def _fits(pt):
        total = 0.0
        for par in tf.paragraphs:
            if not par.text.strip():
                continue
            try:
                # ширина с запасом: PowerPoint не кернит GPOS и набирает шире
                # браузерных/HB-метрик — без запаса PP переносит РАНЬШЕ замера
                total += fit.text_height_in(
                    par.text, name, pt, bw * 0.985,
                    line_spacing=_para_line_spacing(par, 1.0))
            except fit.FontNotFound:
                return True          # шрифта нет — не мерим и не жмём
        return total <= bh + 0.02

    pt = size
    while pt > mn and not _fits(pt):
        pt = max(mn, pt - 2)
    if pt >= size or not _fits(pt) and pt > mn:
        return None
    for par in tf.paragraphs:
        for r in par.runs:
            r.font.size = Pt(pt)
    return pt


def strip_bullets_shape(sh, only_bulleted=False):
    """Снять текстовые буллеты с ОДНОГО шейпа: в pPr абзаца вычищаются
    buChar/buAutoNum/bu-стили и ставится a:buNone. only_bulleted=True —
    трогать лишь абзацы с ЯВНЫМ буллетом в pPr (правило «один тезис — не
    список», user 2026-07-29): дизайнерские marL/indent небуллетных абзацев
    не задеваются. Буллеты-ШЕЙПЫ (точки-кружки) не трогаются."""
    if not getattr(sh, "has_text_frame", False):
        return 0
    n = 0
    for par in sh.text_frame.paragraphs:
        if only_bulleted:
            pPr0 = par._p.find(qn("a:pPr"))
            if pPr0 is None or (pPr0.find(qn("a:buChar")) is None
                                and pPr0.find(qn("a:buAutoNum")) is None):
                continue
        pPr = par._p.get_or_add_pPr()
        for tag in ("a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct",
                    "a:buSzPts", "a:buFontTx", "a:buFont", "a:buNone",
                    "a:buAutoNum", "a:buChar"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        bu = pPr.makeelement(qn("a:buNone"), {})
        ref = pPr.find(qn("a:defRPr"))
        if ref is None:
            ref = pPr.find(qn("a:extLst"))
        if ref is not None:
            ref.addprevious(bu)
        else:
            pPr.append(bu)
        # буллетный висячий отступ (marL + отрицательный indent) без буллета
        # оставляет переносы строк с «лесенкой» — обнуляем вместе с буллетом
        if pPr.get("marL") or pPr.get("indent"):
            pPr.set("marL", "0")
            pPr.set("indent", "0")
        n += 1
    return n


def strip_bullets_slide(slide):
    """slots.bullets=false (v0.16.0, фидбек D медиакита «desc рендерится
    буллетами»): снять текстовые буллеты со всех шейпов слайда."""
    return sum(strip_bullets_shape(sh) for sh in slide.shapes)


def decolor_accent_cards(slide, names):
    """slots.accent=false (v0.16.0, фидбек D медиакита): жёлтая
    приоритет-карточка донора перекрашивается в серую (accent1 → accent2),
    когда карточки равноправны. Трогаем ТОЛЬКО шейпы из fill-map списка
    accent_cards и только со схемной заливкой accent1."""
    out = []
    for nm in names or []:
        sh = find_shape_by_name(slide, nm)
        if sh is None:
            continue
        try:
            f = sh.fill
            if (f.type == MSO_FILL.SOLID
                    and f.fore_color.type == MSO_COLOR_TYPE.SCHEME
                    and f.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1):
                f.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
                out.append(nm)
        except Exception:
            continue
    return out


def _para_style(par):
    """Снять стиль абзаца дизайнерского слайда: (pPr, rPr первого рана), deepcopy."""
    pPr = par._p.find(qn("a:pPr"))
    rPr = par.runs[0]._r.find(qn("a:rPr")) if par.runs else None
    return (copy.deepcopy(pPr) if pPr is not None else None,
            copy.deepcopy(rPr) if rPr is not None else None)


def _apply_para_style(par, style):
    """Переназначить абзацу снятый дизайнерский стиль (pPr + rPr каждого рана)."""
    pPr, rPr = style
    p = par._p
    old = p.find(qn("a:pPr"))
    if old is not None:
        p.remove(old)
    if pPr is not None:
        p.insert(0, copy.deepcopy(pPr))
    for run in par.runs:
        r = run._r
        old_r = r.find(qn("a:rPr"))
        if old_r is not None:
            r.remove(old_r)
        if rPr is not None:
            r.insert(0, copy.deepcopy(rPr))


def _style_sig(el):
    """Рекурсивная сигнатура элемента: тег ({ns}имя) + атрибуты + дети.
    НЕ через tostring: у отсоединённой копии и элемента в дереве сериализуются
    разные xmlns-декларации → ложные несовпадения."""
    if el is None:
        return None
    return (el.tag, tuple(sorted(el.attrib.items())),
            tuple(_style_sig(ch) for ch in el))


def _check_agenda_styles(tf, active_k, active, regular):
    """Машинный чекер перестановки (процесс-урок 2026-07-02: каждая новая
    композиция получает геометрический/стилевой чек ДО показа): rPr каждого
    рана байт-в-байт равен снятому дизайнерскому стилю своей роли."""
    problems = []
    for i, par in enumerate(tf.paragraphs):
        want = active if i == active_k else regular
        for run in par.runs:
            got = run._r.find(qn("a:rPr"))
            if _style_sig(got) != _style_sig(want[1]):
                problems.append(f"para {i}: rPr ≠ дизайнерскому "
                                f"{'активному' if i == active_k else 'обычному'}")
                break
    return problems


def _para_spc_in(par) -> float:
    """spcBef + spcAft абзаца в дюймах (<a:spcPts val=сотые-pt>). У агенды Еды
    18+18pt на абзац — на 4-5 пунктах это дюймы, без них замер льстит."""
    total = 0.0
    pPr = par._p.find(qn("a:pPr"))
    if pPr is not None:
        for tag in ("a:spcBef", "a:spcAft"):
            el = pPr.find(qn(tag))
            if el is not None:
                pts = el.find(qn("a:spcPts"))
                if pts is not None and pts.get("val"):
                    total += int(pts.get("val")) / 100.0 / 72.0
    return total


def _agenda_overflows(sh, slot, active_k) -> bool:
    """Замер агенды с РОЛЕВОЙ типографикой: активный/обычный пункт меряются
    своим шрифтом и кеглем (slot.measure; у #7 обычные пункты наследуют
    66pt YS Text Regular — слотовый 96pt дал бы ложный overflow)."""
    m = slot.get("measure")
    if not m:
        return slot_overflows(sh, slot)
    bw, bh = slot.get("box_w_in"), slot.get("box_h_in")
    if not (bw and bh):
        return False
    total = 0.0
    for i, par in enumerate(sh.text_frame.paragraphs):
        if not par.text.strip():
            continue
        role = m["active"] if i == active_k else m["regular"]
        try:
            total += fit.text_height_in(
                par.text, role["font"], float(role["size_pt"]), bw,
                line_spacing=_para_line_spacing(par, 1.0)) + _para_spc_in(par)
        except fit.FontNotFound:
            return False
    return total > bh + 0.05


def fill_agenda_slot(slide, slot, items, active_index=0):
    """Агенда-per-section (user-правило 2026-07-02): в длинном деке на каждый
    раздел свой агенда-слайд, выделен СВОЙ пункт. Стили АКТИВНОГО (пара 0:
    #2 = 115pt Heavy / #7 = solid) и ОБЫЧНОГО (пара 1: 72pt / полупрозрачный
    alpha 50%) снимаются С САМОГО дизайнерского слайда ДО заливки и
    переназначаются по active_index — перестановка двух дизайнерских стилей,
    НЕ мутация типографики. Каждый абзац получает явный стиль, поэтому пунктов
    может быть и больше, чем у дизайнера (add_paragraph больше не теряет кегль)."""
    sh = _locate_slot_shape(slide, slot)
    if sh is None:
        return {"found": False, "find": slot.get("match_name") or slot.get("find")}
    tf = sh.text_frame
    active = _para_style(tf.paragraphs[0])
    regular = _para_style(tf.paragraphs[1]) if len(tf.paragraphs) > 1 else active
    if slot.get("clear_soft_breaks"):
        clear_soft_breaks(tf)
    items = [str(x) for x in (items or [])]
    if not items:
        return {"found": True, "skipped": True, "_shape": sh}
    k = max(0, min(int(active_index or 0), len(items) - 1))
    fill_lines(sh, items)
    for i, par in enumerate(tf.paragraphs):
        _apply_para_style(par, active if i == k else regular)
    strip_empty_lines(tf)
    rep = {"found": True, "filled": True, "lines": len(items),
           "active_index": k, "_shape": sh}
    if int(active_index or 0) != k:
        rep["active_index_clamped"] = True
    prob = _check_agenda_styles(tf, k, active, regular)
    if prob:
        rep["style_problems"] = prob
    if _agenda_overflows(sh, slot, k):
        rep["overflow"] = True       # сигнал сократить пункты; шрифт НЕ трогаем
    return rep


def _wrap_to_visible(sh, slot, text) -> list:
    """Заголовок шире ВИДИМОЙ ширины (фото перекрывает бокс справа, #13/#15) →
    ЯВНЫЙ перенос строк по box_w_in, ГЕОМЕТРИЮ НЕ ТРОГАЕМ (урок 2026-07-02:
    placeholder без своего xfrm от set width улетает в (0,0); user-fix = перенос
    текста при нетронутом боксе). При 2 строках выбираем СБАЛАНСИРОВАННУЮ точку
    разрыва (короткого хвоста не оставляем), больше max_lines строк не делаем."""
    font = slot.get("font") or _run_font_size(sh)[0]
    size = _run_font_size(sh)[1] or float(slot.get("size_pt") or 0)
    bw = slot.get("box_w_in")
    if not (font and size and bw):
        return [text]
    try:
        if fit.text_width_in(text, font, size) <= bw:
            return [text]
        words = [w for w in text.split(" ") if w]   # НЕ .split(): он рвёт и по NBSP,
        best = None                                  # разклеивая «на\xa0высоте»
        for k in range(1, len(words)):          # один разрыв → 2 строки, баланс
            l1, l2 = " ".join(words[:k]), " ".join(words[k:])
            w1, w2 = fit.text_width_in(l1, font, size), fit.text_width_in(l2, font, size)
            if w1 <= bw and w2 <= bw:
                cand = (abs(w1 - w2), [l1, l2])
                best = min(best, cand) if best else cand
        if best:
            return best[1]
        lines = fit.wrap_lines(text, font, size, bw)  # >2 строк — жадный перенос
        return lines[: int(slot.get("max_lines", len(lines)))]
    except fit.FontNotFound:
        return [text]


def _collect_named(shapes, name, recurse, anc):
    """(shape, key_top, key_left) для всех текст-шейпов с именем name. Вложенные в
    группу шейпы сортируются по позиции ВНЕШНЕЙ группы (anc), т.к. их собственные
    координаты — в дочерней системе группы (chOff/chExt); для #10 «4 факта в 4
    группах» это даёт порядок групп сверху-вниз."""
    out = []
    for sh in shapes:
        if sh.name == name and sh.has_text_frame:
            kt, kl = anc if anc else (sh.top or 0, sh.left or 0)
            out.append((sh, kt, kl))
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP and recurse:
            child_anc = anc if anc else (sh.top or 0, sh.left or 0)
            out.extend(_collect_named(sh.shapes, name, recurse, child_anc))
    return out


def find_shapes_by_name(slide, name, *, sort="doc", recurse=False, skip_regex=None,
                        max_w_in=None):
    """Все текст-шейпы с данным именем (повторяемая роль: колонки/карточки/шаги/факты).
    По умолчанию — порядок документа (у дизайнерских слайдов Еды = порядок чтения,
    проверено рендером); можно пересортировать по позиции (top/left/top_left/left_top).
    recurse — заходить в группы (сорт по позиции внешней группы, Tier B #10). skip_regex —
    выкинуть шейпы, чей текст матчит regex (dual-role #19: одно имя «Title 3» у блоков И
    у номеров «01/02/03» → skip_regex=r'^\\d\\d$' оставляет только блоки). max_w_in —
    выкинуть шейпы ШИРЕ порога (dual-role по геометрии, #21: одно имя «Text Placeholder 3»
    у баннера-интро 15.98" И у колонок 5.74" → max_w_in=10 оставляет колонки)."""
    items = _collect_named(slide.shapes, name, recurse, None)
    if skip_regex:
        rx = re.compile(skip_regex)
        items = [it for it in items if not rx.match((it[0].text_frame.text or "").strip())]
    if max_w_in:
        items = [it for it in items if (it[0].width or 0) <= max_w_in * 914400]
    if sort == "top":
        items.sort(key=lambda it: it[1])
    elif sort == "left":
        items.sort(key=lambda it: it[2])
    elif sort == "top_left":
        items.sort(key=lambda it: (it[1], it[2]))
    elif sort == "left_top":
        items.sort(key=lambda it: (it[2], it[1]))
    return [it[0] for it in items]


def find_shapes_by_text(slide, prefix, *, ci=False, sort="doc", recurse=False):
    """Все текст-шейпы, чей текст начинается с prefix (повторяемая роль, когда имена
    РАЗНЫЕ, а заглушка одна — напр. 10 карточек команды «Костя Константинов»/«Роль…»).
    Стабильный ключ §5 по тексту. Сорт как в find_shapes_by_name."""
    pref = prefix.lower() if ci else prefix
    out = []

    def walk(shapes, anc):
        for sh in shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                hay = t.lower() if ci else t
                if hay.startswith(pref):
                    kt, kl = anc if anc else (sh.top or 0, sh.left or 0)
                    out.append((sh, kt, kl))
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP and recurse:
                walk(sh.shapes, anc if anc else (sh.top or 0, sh.left or 0))

    walk(slide.shapes, None)
    if sort == "top":
        out.sort(key=lambda it: it[1])
    elif sort == "left":
        out.sort(key=lambda it: it[2])
    elif sort == "top_left":
        out.sort(key=lambda it: (it[1], it[2]))
    elif sort == "left_top":
        out.sort(key=lambda it: (it[2], it[1]))
    return [it[0] for it in out]


def find_shapes_by_names(slide, names, *, sort="left", recurse=False):
    """Повторяемая роль, когда шейпы ИНДИВИДУАЛЬНО именованы И тексты разные — их не
    собрать ни match_name (имена разные), ни find_shapes_by_text (заглушки разные).
    Дизайн Городских: заголовки колонок Rectangle 8/9/10 + тексты TextBox 12/13/14.
    Собираем по ЯВНОМУ списку имён и сортируем по позиции (дефолт left = порядок чтения
    ряда). Еда этот путь не использует (её повторы одноимённы) → поведение не меняется."""
    items = []
    for nm in names:
        items.extend(_collect_named(slide.shapes, nm, recurse, None))
    if sort == "top":
        items.sort(key=lambda it: it[1])
    elif sort == "left":
        items.sort(key=lambda it: it[2])
    elif sort == "top_left":
        items.sort(key=lambda it: (it[1], it[2]))
    elif sort == "left_top":
        items.sort(key=lambda it: (it[2], it[1]))
    return [it[0] for it in items]


def fill_repeated_slot(slide, slot, values):
    """Заполнить N одинаковых по роли шейпов (match_name / match_names / find) списком
    значений в порядке чтения (Tier B/§3Б: колонки, шаги, карточки, пункты). Число
    значений СОГЛАСУЕМ с числом слотов дизайнера (контракт: reword/rebalance под N, а не
    генерим); лишние слоты чистим (не оставляем заглушку), лишних значений быть не должно —
    это сигнал резать/сплитить (FILL не генерит). fit_policy: value_shrink (число) /
    reword_first|text_only (проза → флаг overflow, типографику не трогаем)."""
    if slot.get("match_names"):                      # явный список индивидуальных имён (Городские)
        shapes = find_shapes_by_names(slide, slot["match_names"], sort=slot.get("sort", "left"),
                                      recurse=slot.get("recurse", False))
        key = "+".join(slot["match_names"])
    elif slot.get("match_name"):
        shapes = find_shapes_by_name(slide, slot["match_name"], sort=slot.get("sort", "doc"),
                                     recurse=slot.get("recurse", False), skip_regex=slot.get("skip_regex"),
                                     max_w_in=slot.get("max_w_in"))
        key = slot["match_name"]
    else:                                            # повторяемый по тексту-заглушке (разные имена)
        shapes = find_shapes_by_text(slide, slot["find"], ci=slot.get("ci", False),
                                     sort=slot.get("sort", "doc"), recurse=slot.get("recurse", False))
        key = slot["find"]
    if not shapes:
        return {"found": False, "key": key}
    vals = list(values) if isinstance(values, (list, tuple)) else [values]
    policy = slot.get("fit_policy", "reword_first")
    rep = {"found": True, "shapes": len(shapes), "values": len(vals),
           "filled": 0, "cleared": 0, "overflow": []}
    group_pt = value_group_pt(shapes, vals, slot) if policy == "value_shrink" else None
    if group_pt is not None:
        rep["group_pt"] = group_pt
    for i, sh in enumerate(shapes):
        if slot.get("clear_soft_breaks"):
            clear_soft_breaks(sh.text_frame)
        if slot.get("force_wrap"):
            # вернуть режим переноса донора-эталона (пилюли тегов Хурала: wrap=none —
            # бокс растягивается под текст; пересохранение v2.2 сбросило в square)
            sh.text_frame._txBody.bodyPr.set("wrap", slot["force_wrap"])
        if i >= len(vals):
            empty_text_frame(sh.text_frame)       # лишний дизайнерский слот → чистим
            rep["cleared"] += 1
            continue
        v = vals[i]
        if policy == "value_shrink":
            set_value_slot(sh, str(v), max_width_in=slot.get("box_w_in"),
                           min_pt=slot.get("min_pt", 24.0), force_pt=group_pt)
        else:
            lines = [str(x) for x in v] if isinstance(v, (list, tuple)) else [str(v)]
            fill_lines(sh, lines)
            strip_empty_lines(sh.text_frame)
            if len(lines) == 1 and strip_bullets_shape(sh, only_bulleted=True):
                # один тезис — не список (user 2026-07-29): дизайнерский буллет
                # спискового слота снимается, когда пункт остался единственным
                rep["solo_bullet_stripped"] = rep.get("solo_bullet_stripped", 0) + 1
            if slot_overflows(sh, slot):
                rep["overflow"].append(i)
        rep["filled"] += 1
    if len(vals) > len(shapes):
        rep["too_many"] = len(vals) - len(shapes)   # FILL-only: сигнал резать/сплитить
    return rep


# O9 (хэндофф №2 2026-08-13, слова дизайнера): «зелёный майлстоун — это
# выделение: важный пункт сторитела красим, не важный оставляем серым».
# Акцент — семантика КОНТЕНТА, не свойство донора: в доноре зелёная всегда
# третья веха (у t7 — ни одной), и без ручки акцент врал.
_TIMELINE_ACCENT_HEX = "71FFB5"     # сырой акцент дизайнера (в теме зелёного нет)


def apply_timeline_accent(slide, accent):
    """Перекрасить пилюли вех (atom:milestone_date, порядок = слева направо,
    как у слота dates): веха №accent (1-based) → #71FFB5, все прочие → серый
    филл донора (снимается с любой неакцентной пилюли — тема dk2/tx2).
    accent=None → ВСЯ линейка серая: донорскую зелёную третью гасим — дефолт
    без ключа не должен утверждать акцент, которого нет в контенте.
    Текст, стрелки, коннекторы и геометрию не трогаем.
    Ошибки честные (ValueError): пилюль нет (вариант/шаблон без milestone_date —
    акцент не поддержан), accent не одиночный int 1..N."""
    pills = find_shapes_by_name(slide, "atom:milestone_date", sort="left",
                                recurse=True)
    if not pills:
        raise ValueError(
            "accent: на слайде нет пилюль atom:milestone_date — акцент вехи "
            "поддержан только линейкой таймлайн-доноров с этим атомом "
            "(Хурал Райдтеха t4–t7)")
    n = len(pills)
    if accent is not None:
        if isinstance(accent, bool) or not isinstance(accent, int):
            raise ValueError(
                f"accent={accent!r}: нужен ОДИН номер вехи (int, 1-based); "
                f"списки не поддержаны — два акцента = ни одного (канон "
                f"дизайнера: одно выделение на слайд)")
        if not 1 <= accent <= n:
            raise ValueError(
                f"accent={accent}: вне линейки 1..{n} этого варианта — укажи "
                f"номер существующей вехи (счёт слева направо)")

    def _fill_of(sh):
        return sh._element.spPr.find(qn("a:solidFill"))

    def _is_green(sf):
        c = sf.find(qn("a:srgbClr")) if sf is not None else None
        return c is not None and (c.get("val") or "").upper() == _TIMELINE_ACCENT_HEX

    grey = next((_fill_of(p) for p in pills
                 if _fill_of(p) is not None and not _is_green(_fill_of(p))), None)
    recolored = 0
    for i, p in enumerate(pills, 1):
        spPr = p._element.spPr
        old = _fill_of(p)
        if i == accent:
            if _is_green(old):
                continue                      # донорская зелёная и так на месте
            new = spPr.makeelement(qn("a:solidFill"), {})
            clr = new.makeelement(qn("a:srgbClr"), {})
            clr.set("val", _TIMELINE_ACCENT_HEX)
            new.append(clr)
        else:
            if old is None or not _is_green(old):
                continue                      # уже серая — не трогаем
            if grey is None:
                continue                      # эталона серого нет (не бывает: N≥4)
            new = copy.deepcopy(grey)
        # заливка встаёт РОВНО на место старой (порядок детей spPr сохраняется);
        # пилюля без явной заливки (стилевая) — после геометрии, по схеме spPr
        if old is not None:
            old.addnext(new)
            spPr.remove(old)
        else:
            geom = spPr.find(qn("a:prstGeom"))
            if geom is None:
                geom = spPr.find(qn("a:xfrm"))
            if geom is not None:
                geom.addnext(new)
            else:
                spPr.insert(0, new)
        recolored += 1
    return {"pills": n, "accent": accent, "recolored": recolored}


def fill_table_slot(slide, slot, rows):
    """Заполнить ячейки таблицы из 2D-списка (rows[r][c] = строка ИЛИ список строк для
    многострочной ячейки; None/за пределами — ячейку не трогаем). Таблица по match_name.
    Базовый режим (гантт, контракт): размеры/стиль НЕ меняем — только текст ячеек
    (формат первого run'а сохранён). Tier C §3Б: `kind:"table"`, `cell(r,c).text`.

    Режим `"trim": true` (datatable #36, user-фидбек 2026-07-16): сетка ужимается
    под контент — лишние строки/колонки удаляются ФИЗИЧЕСКИ (a:tr / gridCol+tc,
    не бланчинг), фрейм пересчитывается по фактическим строкам (высоты донорские,
    без «distribute rows»), незаполненные ячейки в оставшейся сетке чистятся от
    болванок донора, выравнивание колонок — по типу контента (правило вёрстки,
    user 2026-07-20): текст → left, числа → right, заголовок — как его колонка."""
    frame, tbl = None, None
    for sh in slide.shapes:
        if sh.name == slot["match_name"] and sh.has_table:
            frame, tbl = sh, sh.table
            break
    if tbl is None:
        return {"found": False, "match_name": slot["match_name"]}
    nr0, nc0 = len(tbl.rows), len(tbl.columns)
    rows = rows or []
    rep = {"found": True, "filled": 0, "dims": [nr0, nc0]}

    if slot.get("trim") and rows:
        rep["donor_dims"] = [nr0, nc0]
        keep_r = min(len(rows), nr0)
        keep_c = min(max((len(r) for r in rows if r is not None), default=nc0), nc0)
        t_el = tbl._tbl
        donor_h = sum(int(tr.get("h") or 0) for tr in t_el.tr_lst)  # ДО удаления
        donor_w = sum(int(gc.get("w") or 0)
                      for gc in t_el.tblGrid.gridCol_lst)           # ДО удаления
        for tr in list(t_el.tr_lst[keep_r:]):
            tr.getparent().remove(tr)
        if keep_c < nc0:
            grid = t_el.tblGrid
            for gc in list(grid.gridCol_lst[keep_c:]):
                grid.remove(gc)
            for tr in t_el.tr_lst:
                for tc in list(tr.tc_lst[keep_c:]):
                    tr.remove(tc)
        rep["trimmed_rows"] = nr0 - keep_r
        rep["trimmed_cols"] = nc0 - keep_c
        hs = [int(tr.get("h") or 0) for tr in t_el.tr_lst]
        ws = [int(gc.get("w") or 0) for gc in t_el.tblGrid.gridCol_lst]
        if slot.get("stretch", True) and all(hs) and donor_h > sum(hs):
            # «растянуть как дизайнер» (user 2026-07-29, слайд-образец владельца;
            # с 2026-07-29 режим ОБЩИЙ для всех шаблонов — "stretch": false в
            # спеке = осознанный отказ): усечённая сетка распределяется на
            # ИСХОДНУЮ высоту донорской таблицы (строки равномерно × k), кегль
            # ячеек растёт с k до потолка — компактная таблица не «липнет»
            # к заголовку на полупустом слайде. Потолок кегля: абсолютные
            # stretch_max_header_pt/stretch_max_body_pt спеки, а без них —
            # ×STRETCH_MAX_FACTOR от дизайнерского кегля рана (шаблон-агностично)
            k = donor_h / sum(hs)
            for tr in t_el.tr_lst:
                tr.set("h", str(int(int(tr.get("h") or 0) * k)))
            hs = [int(tr.get("h") or 0) for tr in t_el.tr_lst]
            # "stretch_font": false (слот fill-map/спеки) — кегль НЕ растёт с k:
            # строки тянутся, кегль остаётся донорским. Для доноров, где ячейки
            # набраны В ПАРИТЕТЕ со спутниками варианта (sowhat Хурала: ячейки
            # 28pt = atom:so_what_body 28pt — решение дизайнера), рост кегля
            # таблицы обгонял плашку в полтора раза, и дизайнер читала это как
            # «в плашке вечно мелко» (хэндофф Ultima 2026-08-11, O3). Дефолт
            # true: рост с потолком — одобренное поведение Городских (24→32pt).
            if slot.get("stretch_font", True):
                cap_hdr = slot.get("stretch_max_header_pt")
                cap_body = slot.get("stretch_max_body_pt")
                for ri in range(len(tbl.rows)):
                    cap_abs = cap_hdr if ri == 0 else cap_body
                    for ci in range(len(tbl.columns)):
                        for par in tbl.cell(ri, ci).text_frame.paragraphs:
                            for run in par.runs:
                                if run.font.size is None:
                                    continue
                                cap = (float(cap_abs) if cap_abs else
                                       run.font.size.pt * STRETCH_MAX_FACTOR)
                                run.font.size = Pt(min(run.font.size.pt * k, cap))
            else:
                rep["stretch_font"] = False
            rep["stretched"] = round(k, 3)
        if slot.get("stretch_width", True) and all(ws) and donor_w > sum(ws):
            # ширинный аналог stretch (user 2026-08-05: «таблица с меньшим
            # числом колонок растягивается на донорскую ширину» — общий режим
            # всех шаблонов, "stretch_width": false — осознанный отказ):
            # оставшиеся колонки равномерно × kw до ИСХОДНОЙ ширины донорской
            # таблицы (у fullwidth-доноров это поля слайда, у таблиц с
            # сайдбаром — граница сайдбара). Кегль от ширины не растёт —
            # им управляет высотный stretch.
            kw = donor_w / sum(ws)
            for gc in t_el.tblGrid.gridCol_lst:
                gc.set("w", str(int(int(gc.get("w") or 0) * kw)))
            ws = [int(gc.get("w") or 0) for gc in t_el.tblGrid.gridCol_lst]
            rep["stretched_width"] = round(kw, 3)
        # фрейм = сумма фактических строк/колонок
        if all(hs):
            frame.height = Emu(sum(hs))
        if all(ws):
            frame.width = Emu(sum(ws))

    nr, nc = len(tbl.rows), len(tbl.columns)   # в trim-режиме nr ≤ len(rows)
    blanked = 0
    for r in range(nr if slot.get("trim") else min(len(rows), nr)):
        row = rows[r]
        if row is None:                 # None-строка = не трогаем (напр. r0 месяцы Ганта)
            continue
        for c in range(nc if slot.get("trim") else min(len(row), nc)):
            val = row[c] if c < len(row) else None
            if val is None:
                if slot.get("trim"):    # болванка донора в оставшейся сетке — чистим
                    _fill_tf(tbl.cell(r, c).text_frame, [""])
                    blanked += 1
                continue
            lines = [str(x) for x in val] if isinstance(val, (list, tuple)) else [str(val)]
            _fill_tf(tbl.cell(r, c).text_frame, lines)
            rep["filled"] += 1
    if blanked:
        rep["blanked"] = blanked

    if slot.get("trim") and rows:
        # выравнивание по типу колонки: смотрим data-строки (r ≥ 1), прочерки
        # нейтральны; числовая колонка (все значения — числа) → right, включая
        # заголовок; иначе → left
        aligned = {"r": [], "l": []}
        for c in range(nc):
            vals = []
            for r in range(1, len(rows)):
                row = rows[r]
                if row is None or c >= len(row) or row[c] is None:
                    continue
                v = row[c]
                for t in (v if isinstance(v, (list, tuple)) else [v]):
                    t = str(t).strip()
                    if t and t not in fit.NEUTRAL_CELLS:
                        vals.append(t)
            if not vals:
                continue
            algn = "r" if all(fit.is_numeric_cell(t) for t in vals) else "l"
            aligned[algn].append(c + 1)
            for r in range(nr):
                for p in tbl.cell(r, c).text_frame.paragraphs:
                    p.alignment = (PP_ALIGN.RIGHT if algn == "r" else PP_ALIGN.LEFT)
        rep["aligned"] = aligned
        rep["dims"] = [nr, nc]
    return rep


def _compose_slot_values(slotvals, content, keys):
    """Композитный слот «одна коробка, роли абзацами» (v2.2 Хурала: карточка =
    заголовок + тело/буллеты в ОДНОМ шейпе, стили абзацев-заглушек сохраняет
    _fill_tf). Собирает vals[i] = [заголовок_i, *строки_тела_i] из deckspec-ключей
    keys — контракт спеки (titles/bodies и т.п.) НЕ меняется. У других шаблонов
    ключа compose нет → ветка мертва (eda-safe)."""
    parts = [slotvals.get(k, getattr(content, k, None)) or [] for k in keys]
    n = max((len(p) for p in parts), default=0)
    vals = []
    for i in range(n):
        lines = []
        for p in parts:
            x = p[i] if i < len(p) else None
            if x is None or x == "":
                continue
            if isinstance(x, (list, tuple)):
                lines.extend(str(y) for y in x)
            else:
                lines.append(str(x))
        vals.append(lines)
    return vals


def fill_structural_slide(slide, variant, content):
    """FILL-исполнитель fill-map (§3Б: структурный тир И Tier-A/B контент — одна машинка,
    всегда FILL, вёрстку/типографику дизайнера не трогаем). Заполнить ВСЕ слоты варианта:
    значение из content.slots[name] (иначе атрибут content.<name>). Слот с "repeat":true
    → повторяемая роль (N шейпов по match_name), иначе — одиночный по тексту-заглушке.
    Слот с "compose": [ключи] — композитный (значения собираются из нескольких
    deckspec-ключей в одну коробку). Возвращает {slot: отчёт}."""
    slotvals = getattr(content, "slots", None) or {}
    report = {}
    for name, slot in variant["slots"].items():
        val = slotvals.get(name, getattr(content, name, None))
        if slot.get("compose"):
            val = _compose_slot_values(slotvals, content, slot["compose"])
        if slot.get("kind") == "table":
            report[name] = fill_table_slot(slide, slot, val or [])
        elif slot.get("agenda_hierarchy"):
            report[name] = fill_agenda_slot(slide, slot, val,
                                            active_index=slotvals.get("active_index", 0))
        elif slot.get("repeat"):
            report[name] = fill_repeated_slot(slide, slot, val or [])
        else:
            report[name] = fill_text_slot(slide, slot, val)
        # Некоторые дизайнерские доноры намеренно накладывают текст на chart.
        # Для контента иной длины/геометрии layout может безопасно развести их,
        # не меняя сам шаблон: координаты задаются явно и только для найденного
        # одиночного slot-shape. Размер и типографика остаются донорскими.
        rep = report.get(name)
        sh = rep.get("_shape") if isinstance(rep, dict) else None
        if sh is not None:
            if slot.get("set_left_in") is not None:
                sh.left = Inches(float(slot["set_left_in"]))
                rep["left_in"] = float(slot["set_left_in"])
            if slot.get("set_top_in") is not None:
                sh.top = Inches(float(slot["set_top_in"]))
                rep["top_in"] = float(slot["set_top_in"])
    # Связанный с контентом декор (например, стрелка рядом с focus-value)
    # можно сдвинуть вместе со slot'ом. Только явные имена и координаты;
    # размер, стиль и содержимое shape не меняются.
    for move in variant.get("move_shapes") or []:
        sh = find_shape_by_name(slide, move.get("match_name") or "",
                                recurse=move.get("recurse", False))
        if sh is None:
            report.setdefault("_move_missing", []).append(move.get("match_name"))
            continue
        if move.get("set_left_in") is not None:
            sh.left = Inches(float(move["set_left_in"]))
        if move.get("set_top_in") is not None:
            sh.top = Inches(float(move["set_top_in"]))
        report.setdefault("_moved", []).append(move.get("match_name"))
    # v0.23.2: focus_fill — «содержание как разделитель» (Городские #51, решение
    # владельца 2026-07-29): у донора залит номер АКТИВНОГО пункта (в файле —
    # пункт 1); движок переносит эту заливку на пункт slots.active_index
    # (0-based, дефолт 0 = как нарисовано), остальным — noFill. Ключа нет у
    # прочих шаблонов → ветка мертва (eda-safe).
    ff = variant.get("focus_fill")
    if ff:
        shapes = find_shapes_by_names(slide, ff["match_names"],
                                      sort=ff.get("sort", "top"))
        k = int(slotvals.get(ff.get("active_slot", "active_index"), 0) or 0)
        if shapes and 0 <= k < len(shapes):
            FILLS = ("solidFill", "gradFill", "blipFill", "pattFill", "noFill")
            def _fill_el(sh):
                sp = sh._element.spPr
                return next((ch for ch in sp if ch.tag.split("}")[-1] in FILLS), None)
            donor_sh = next((s for s in shapes
                             if _fill_el(s) is not None
                             and _fill_el(s).tag.split("}")[-1] != "noFill"), None)
            if donor_sh is not None and shapes[k] is not donor_sh:
                donor_fill = _fill_el(donor_sh)
                # цвет ЦИФРЫ свапается вместе с заливкой: у актива дизайнер
                # держит контраст (чёрная цифра на белом), у рядовых — наследуемый
                # (белый на тёмном); иначе цифра экс-актива пропадает на фоне
                def _txt_fill(sh):
                    for para in sh.text_frame.paragraphs:
                        for r in para.runs:
                            rPr = r._r.find(qn("a:rPr"))
                            if rPr is not None:
                                return rPr.find(qn("a:solidFill"))
                    return None
                def _set_txt_fill(sh, tmpl):
                    for para in sh.text_frame.paragraphs:
                        for r in para.runs:
                            rPr = r._r.get_or_add_rPr()
                            old = rPr.find(qn("a:solidFill"))
                            if old is not None:
                                rPr.remove(old)
                            if tmpl is not None:
                                ln = rPr.find(qn("a:ln"))
                                pos = (list(rPr).index(ln) + 1) if ln is not None else 0
                                rPr.insert(pos, copy.deepcopy(tmpl))
                ref_plain = next((s for s in shapes if s is not donor_sh), None)
                donor_txt = copy.deepcopy(_txt_fill(donor_sh)) if _txt_fill(donor_sh) is not None else None
                plain_txt = (copy.deepcopy(_txt_fill(ref_plain))
                             if ref_plain is not None and _txt_fill(ref_plain) is not None else None)
                for s in shapes:
                    old = _fill_el(s)
                    new = (copy.deepcopy(donor_fill) if s is shapes[k]
                           else s._element.spPr.makeelement(qn("a:noFill"), {}))
                    if old is not None:
                        s._element.spPr.replace(old, new)
                    else:
                        ln = s._element.spPr.find(qn("a:ln"))
                        (s._element.spPr.insert(list(s._element.spPr).index(ln), new)
                         if ln is not None else s._element.spPr.append(new))
                    _set_txt_fill(s, donor_txt if s is shapes[k] else plain_txt)
                # парные текстбоксы пунктов (companion_names): у донора яркость
                # текста привязана к активному пункту — свапаем и её
                comp = ff.get("companion_names")
                if comp:
                    cshapes = find_shapes_by_names(slide, comp,
                                                   sort=ff.get("sort", "top"))
                    if len(cshapes) == len(shapes):
                        di = shapes.index(donor_sh)
                        c_donor = (copy.deepcopy(_txt_fill(cshapes[di]))
                                   if _txt_fill(cshapes[di]) is not None else None)
                        c_ref = next((c for i, c in enumerate(cshapes) if i != di), None)
                        c_plain = (copy.deepcopy(_txt_fill(c_ref))
                                   if c_ref is not None and _txt_fill(c_ref) is not None else None)
                        for i, c in enumerate(cshapes):
                            _set_txt_fill(c, c_donor if i == k else c_plain)
                report["_focus_fill"] = {"active": k, "of": len(shapes)}
    # v0.15.0: залипшие статики донора (текстбокс с номером слайда-донора и т.п.) —
    # вычищаются всегда, объявляются в fill-map варианта списком strip_shapes
    for nm in variant.get("strip_shapes") or []:
        sh = find_shape_by_name(slide, nm)
        if sh is not None:
            sh._element.getparent().remove(sh._element)
            report.setdefault("_stripped", []).append(nm)
    # v0.15.0: декоративные подсветки донора (жёлтые пилюли datatable #36) прибиты
    # к позициям ячеек — живут только над ПОЛНОЙ дизайнерской сеткой. Сетка ужата
    # или ячейки пустые → снять; явный slots.highlights=false → снять всегда.
    dec = variant.get("decor_highlights") or []
    if dec:
        trep = next((v for v in report.values()
                     if isinstance(v, dict) and v.get("donor_dims")), None)
        partial = bool(trep and (trep.get("trimmed_rows") or trep.get("trimmed_cols")
                                 or trep.get("blanked")))
        if slotvals.get("highlights") is False or partial:
            gone = []
            for nm in dec:
                sh = find_shape_by_name(slide, nm)
                if sh is not None:
                    sh._element.getparent().remove(sh._element)
                    gone.append(nm)
            if gone:
                report["_decor_removed"] = gone
    _apply_push_below(variant, report)
    _check_multiline_conflicts(variant, report)
    broken = []
    for name, rep in report.items():         # вырожденная геометрия = класс бага
        sh = rep.get("_shape") if isinstance(rep, dict) else None
        if sh is None:
            continue
        try:
            if (sh.width or 0) <= 0 or (sh.height or 0) <= 0:
                broken.append(f"{name}: шейп выродился (w/h=0 — placeholder "
                              f"потерял наследуемый xfrm?)")
        except Exception:
            pass
    if broken:
        report["_geometry_broken"] = broken
    return report


def _check_multiline_conflicts(variant, report):
    """Слот со спековым "multiline_requires_empty": [слоты]: если заполненный
    текст лёг БОЛЬШЕ чем в designer_lines строк, перечисленные слоты обязаны
    быть пусты — сдвигать некуда (ниже плашка/чарт), а молчаливый наезд текстов
    хуже честной ошибки (user-фидбек по чарт-слайду 2026-07-03). Нарушение
    кладётся в report["_slot_conflicts"] — build_deck превращает в DeckError."""
    problems = []
    for name, slot in variant["slots"].items():
        others = slot.get("multiline_requires_empty")
        if not others:
            continue
        rep = report.get(name) or {}
        sh = rep.get("_shape")
        if sh is None or not rep.get("filled"):
            continue
        font = slot.get("font") or _run_font_size(sh)[0]
        size = _run_font_size(sh)[1] or float(slot.get("size_pt") or 0)
        bw = slot.get("box_w_in")
        if not (font and size and bw):
            continue
        try:
            nlines = fit.line_count(sh.text_frame.text, font, size, bw)
        except fit.FontNotFound:
            continue
        if nlines <= int(slot.get("designer_lines", 1)):
            continue
        filled = [o for o in others if (report.get(o) or {}).get("filled")]
        if filled:
            problems.append(
                f"«{name}» занял {nlines} строки — по вёрстке дизайнера слоты "
                f"{filled} обязаны быть пусты: сократите «{name}» до одной строки "
                f"или уберите {filled}")
    if problems:
        report["_slot_conflicts"] = problems


def _apply_push_below(variant, report):
    """Многострочный заголовок → сдвиг блоков ниже (user-вопрос 2026-07-02:
    «заголовок в несколько строк и опустить текстовый блок вниз»). Слот со
    "push_below": [имена слотов] и max_lines > 1: меряем ФАКТИЧЕСКИЕ строки
    заполненного текста (глифы, той же мерой, что и overflow) и сдвигаем шейпы
    целевых слотов вниз на высоту добавочных строк. Разрешено ТОЛЬКО там, где
    спека это объявила (лёгкая генерация поверх FILL, geometry-часть контентного
    тира); дизайнерский однострочный случай — ноль сдвига, точь-в-точь."""
    for name, slot in variant["slots"].items():
        targets = slot.get("push_below")
        if not targets:
            continue
        rep = report.get(name) or {}
        sh = rep.get("_shape")
        if sh is None or not rep.get("filled"):
            continue
        font = slot.get("font") or _run_font_size(sh)[0]
        size = _run_font_size(sh)[1] or float(slot.get("size_pt") or 0)
        bw = slot.get("box_w_in")
        if not (font and size and bw):
            continue
        try:
            nlines = fit.line_count(sh.text_frame.text, font, size, bw)
        except fit.FontNotFound:
            continue
        base = int(slot.get("designer_lines", 1))
        if nlines <= base:
            continue
        spacing = _para_line_spacing(sh.text_frame.paragraphs[0], 1.0)
        extra_in = (nlines - base) * size * spacing / 72.0
        for tgt in targets:
            tsh = (report.get(tgt) or {}).get("_shape")
            if tsh is not None:
                tsh.top = tsh.top + Emu(int(extra_in * 914400))
        rep["pushed_below_in"] = round(extra_in, 3)
