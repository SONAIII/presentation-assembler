#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""generate.py — ветка GENERATE роутера (§3): «иначе собери».

Clone-elastic: для числа карточек N, которого нет среди готовых слайдов, берём
готовую дизайнерскую тройку {фон+значение+подпись}, КЛОНИРУЕМ её до N и
перераскладываем по зоне. Геометрия/заливка/типографика остаются byte-exact «из
мастера» (§3) — параметрически считаем только позицию, число, акцент-по-индексу и
shrink. Заполнение и замеры — через lib.fill / lib.fit (тот же измеритель, что у
рендерера и линта).

Безопасность файла: каждый клон получает свежий <p:cNvPr id> и теряет
a16:creationId (дубли id/creationId роняют PowerPoint — см. crash_check). Удаление
лишних карточек — это lxml-remove УЗЛА в памяти, не файла (deny-правило на rm не
затрагивается).
"""
from __future__ import annotations

import copy
import math
import os
import re
import sys
from collections import namedtuple

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fill as F  # noqa: E402
import fit  # noqa: E402

EMU = 914_400
A16_CREATION = "{http://schemas.microsoft.com/office/drawing/2014/main}creationId"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

Card = namedtuple("Card", "bg value label")


class GenerateError(Exception):
    pass


def _in(emu) -> float:
    return emu / EMU


# ------------------------------------------------------------ поиск карточек


def locate_cards(slide, spec) -> list[Card]:
    """Найти тройки ГЕОМЕТРИЧЕСКИ (имена подписей дублируются): bg = roundRect в
    зоне, сортируем по top; value/label = текст-боксы по x-якорю внутри y-диапазона
    своего bg. Возвращает карточки сверху вниз."""
    zone, cu = spec["zone"], spec["card_unit"]
    prst = cu["bg"]["geom_prst"]
    bgs = []
    for sh in slide.shapes:
        pg = sh._element.find(f".//{qn('a:prstGeom')}")
        if pg is not None and pg.get("prst") == prst and abs(_in(sh.left) - zone["left_in"]) < 0.25:
            bgs.append(sh)
    bgs.sort(key=lambda s: s.top)

    boxes = [sh for sh in slide.shapes if sh.has_text_frame]

    def find_box(bg, match_left_in):
        y0, y1 = bg.top, bg.top + bg.height
        pad = int(0.35 * EMU)
        for sh in boxes:
            cy = sh.top + sh.height // 2
            if y0 - pad <= cy <= y1 + pad and abs(_in(sh.left) - match_left_in) < 0.3:
                return sh
        return None

    cards = []
    for bg in bgs:
        v = find_box(bg, cu["value"]["match_left_in"])
        l = find_box(bg, cu["label"]["match_left_in"])
        cards.append(Card(bg, v, l))
    return cards


# ------------------------------------------------------- клонирование (crash-safe)


def next_id_allocator(prs):
    """Аллокатор свежих cNvPr id, заведомо выше всех существующих (≥1000)."""
    mx = 0
    for s in prs.slides:
        for cnv in s._element.findall(f".//{qn('p:cNvPr')}"):
            try:
                mx = max(mx, int(cnv.get("id")))
            except (TypeError, ValueError):
                pass
    state = {"n": max(1000, mx + 1)}

    def alloc():
        v = state["n"]
        state["n"] += 1
        return v

    return alloc


def _renumber_ids(el, alloc) -> int:
    """Назначить свежие id всем cNvPr в поддереве. Вернуть id самого шейпа (первый)."""
    own = None
    for cnv in el.findall(f".//{qn('p:cNvPr')}"):
        new = alloc()
        cnv.set("id", str(new))
        if own is None:
            own = new
    return own


def _strip_a16(el) -> None:
    for c in el.findall(f".//{A16_CREATION}"):
        c.getparent().remove(c)


def assert_clone_safe(el) -> None:
    """Клонируемая тройка не должна тащить rel-ссылки (иначе orphan r:id)."""
    s = etree.tostring(el).decode()
    if "r:embed" in s or "r:id" in s or "r:link" in s:
        raise GenerateError("клон карточки содержит rel-ссылку (r:embed/r:id) — небезопасно")


def _shape_by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    return None


def clone_triplet(slide, template: Card, alloc) -> Card:
    """Склонировать тройку template, дать свежие id, снять a16, вернуть как Card
    из живых Shape-объектов (для дальнейшей правки через API)."""
    spTree = slide.shapes._spTree
    ids = []
    for src in (template.bg, template.value, template.label):
        el = copy.deepcopy(src._element)
        assert_clone_safe(el)
        own = _renumber_ids(el, alloc)
        _strip_a16(el)
        spTree.append(el)
        ids.append(own)
    return Card(_shape_by_id(slide, ids[0]), _shape_by_id(slide, ids[1]), _shape_by_id(slide, ids[2]))


# --------------------------------------------------------- раскладка / акцент


def set_card_accent(card: Card, scheme: str) -> None:
    sc = card.bg._element.find(f"{qn('p:spPr')}/{qn('a:solidFill')}/{qn('a:schemeClr')}")
    if sc is not None:
        sc.set("val", scheme)


def place_card(card: Card, left_in: float, top_in: float, card_h_in: float, spec: dict) -> None:
    """Поставить bg/value/label в слот: значение по вертикали по центру карточки,
    подпись — на фиксированном rel_top от верха карточки."""
    zone, cu, rf = spec["zone"], spec["card_unit"], spec["reflow"]
    v, l = cu["value"], cu["label"]
    card.bg.left = Inches(left_in)
    card.bg.top = Inches(top_in)
    card.bg.width = Inches(zone["width_in"])
    card.bg.height = Inches(card_h_in)
    if card.value is not None:
        card.value.left = Inches(left_in + v["rel_left_in"])
        card.value.width = Inches(v["width_in"])
        card.value.height = Inches(v["height_in"])
        card.value.top = Inches(top_in + (card_h_in - v["height_in"]) / 2.0)
    if card.label is not None:
        card.label.left = Inches(left_in + l["rel_left_in"])
        card.label.width = Inches(l["width_in"])
        card.label.top = Inches(top_in + rf["label_rel_top_in"])


def card_height_for(n: int, spec: dict) -> float:
    """Высота карточки при N (GENERATE): заполнить зону, не выше дизайнерской, не
    ниже минимума (иначе raise → 2-колоночный вариант, отложен)."""
    rf, zone = spec["reflow"], spec["zone"]
    zone_h = zone["bottom_in"] - zone["top_in"]
    gap = rf["gen_gap_in"]
    card_h = (zone_h - gap * (n - 1)) / n
    card_h = min(card_h, rf["designer_card_h_in"])
    if card_h < rf["min_card_h_in"]:
        raise GenerateError(
            f"N={n}: card_h={card_h:.2f}\" < min {rf['min_card_h_in']}\" — нужен "
            f"2-колоночный вариант (отложено)"
        )
    return card_h


# ------------------------------------------------------------------ API


def generate_statgrid(prs, spec, content, base_slide_index: int):
    """Вырастить готовый 3-карточный слайд до N стат-карточек и заполнить.
    content = {form, items:[{value,label}], ...}. Возвращает slide."""
    slide = prs.slides[base_slide_index]
    cards = locate_cards(slide, spec)
    if len(cards) < 1:
        raise GenerateError("на базовом слайде не найдено стат-карточек")

    items = content["items"] if isinstance(content, dict) else content.items
    n = len(items)
    card_h = card_height_for(n, spec)
    zone, rf = spec["zone"], spec["reflow"]
    gap = rf["gen_gap_in"]

    alloc = next_id_allocator(prs)
    template = cards[0]  # полностью укомплектованная тройка; акцент выставим по слоту

    work = list(cards)
    while len(work) < n:
        work.append(clone_triplet(slide, template, alloc))
    while len(work) > n:  # N<3: убрать лишние тройки (remove узла, не файла)
        extra = work.pop()
        for sh in (extra.bg, extra.value, extra.label):
            if sh is not None:
                sh._element.getparent().remove(sh._element)

    v, vfit = spec["card_unit"]["value"], spec["fit"]["value"]
    # потолок роста значения (right align, растёт влево) = внутренняя зона
    # плашки; люфт «бокс + 1"» отменён — user-фидбек 2026-07-07 (вылет на фото)
    max_val_w = F.value_max_width(v, vfit)

    # однородные значения — ЕДИНЫЙ кегль и на GENERATE-ветке (user 2026-07-02;
    # до 2026-07-07 каждая клон-карточка ужималась поодиночке → разнокегельно)
    vals = [str(it["value"]) for it in items]
    vshapes = [c.value for c in work[:len(vals)] if c.value is not None]
    group_pt = F.value_group_pt(vshapes, vals,
                                {"box_w_in": max_val_w, "min_pt": vfit["min_pt"],
                                 "font": v.get("font"),
                                 "size_pt": v.get("size_pt")})

    overflow = []
    for i, (card, item) in enumerate(zip(work, items)):
        top = zone["top_in"] + i * (card_h + gap)
        place_card(card, zone["left_in"], top, card_h, spec)
        set_card_accent(card, "accent1" if i == 0 else "accent2")
        if card.value is not None:
            F.set_value_slot(card.value, str(item["value"]),
                             max_width_in=max_val_w, min_pt=vfit["min_pt"],
                             force_pt=group_pt)
            if F.value_overflows(card.value, item["value"], max_val_w):
                overflow.append(f"value_{i + 1}")
        if card.label is not None:
            F.fill_lines(card.label, [str(item["label"])])
            F.strip_empty_lines(card.label.text_frame)
    return overflow


# ============================================================================
# Локатор горизонтальных юнитов + crash-safe клон шейпа. Живой потребитель —
# atoms.clone_atom (clone_shape_el). Плиточный layout.generate_tiles выпилен
# (0.21.0, решение владельца 2026-07-23); прежние clone-elastic hstrip-генераторы
# удалены ещё раньше (§4).
# ============================================================================

Unit = namedtuple("Unit", "bg members")   # members: list[(role, shape)]
_NUM_RE = re.compile(r"^\d{1,3}$")


def _prst(sh) -> str:
    pg = sh._element.find(f".//{qn('a:prstGeom')}")
    return pg.get("prst") if pg is not None else ""


def _classify_member(sh) -> str:
    """Роль шейпа внутри карточки: freeform → бейдж-декор; текст-цифра → номер
    (01/02/03, перегенерим по индексу); прочий текст → блок; иначе — декор."""
    if sh.shape_type == MSO_SHAPE_TYPE.FREEFORM:
        return "badge"
    if sh.has_text_frame:
        t = sh.text_frame.text.strip()
        if _NUM_RE.match(t):
            return "number"
        if t:
            return "block"
    return "decor"


def locate_hstrip_units(slide, gspec) -> list[Unit]:
    """Найти горизонтальные юниты ГЕОМЕТРИЧЕСКИ. bg-карточка = roundRect-автофигура в
    зоне, чья ширина < половины зоны (отсекает полноширинный баннер/заголовок). Члены
    юнита = шейпы, чей центр по X внутри карточки, не шире её и в вертикальной полосе
    карточки (отсекает заголовок сверху и футер снизу). Сортируем карточки слева
    направо = порядок чтения."""
    zone = gspec["zone"]
    zl, zw = zone["left_in"], zone["width_in"]
    prst = gspec.get("unit", {}).get("bg_prst", "roundRect")
    max_bg_w = zw * gspec.get("unit", {}).get("bg_max_w_frac", 0.5)
    bgs = []
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and _prst(sh) == prst:
            if zl - 0.3 <= _in(sh.left) < zl + zw and _in(sh.width) < max_bg_w:
                bgs.append(sh)
    bgs.sort(key=lambda s: s.left)

    pad_x, pad_top, pad_bot = 0.15, 0.55, 0.35
    units = []
    for bg in bgs:
        bl, br = _in(bg.left), _in(bg.left) + _in(bg.width)
        bt, bb = _in(bg.top), _in(bg.top) + _in(bg.height)
        members = []
        for sh in slide.shapes:
            if sh is bg or sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and _prst(sh) == prst:
                continue  # не втягиваем другие карточки как членов
            cx = _in(sh.left) + _in(sh.width) / 2.0
            cy = _in(sh.top) + _in(sh.height) / 2.0
            if (bl - pad_x <= cx <= br + pad_x and _in(sh.width) <= _in(bg.width) * 1.05
                    and bt - pad_top <= cy <= bb + pad_bot):
                members.append((_classify_member(sh), sh))
        units.append(Unit(bg, members))
    return units


def clone_shape_el(slide, src_shape, alloc) -> object:
    """Склонировать один шейп (crash-safe), вернуть живой Shape."""
    el = copy.deepcopy(src_shape._element)
    assert_clone_safe(el)
    own = _renumber_ids(el, alloc)
    _strip_a16(el)
    slide.shapes._spTree.append(el)
    return _shape_by_id(slide, own)


# _set_bg_scheme — общий помощник schemeClr-заливки bg-шейпа (потребитель —
# set_card_accent статгрида; плиточный GENERATE выпилен 0.21.0).


def _set_bg_scheme(bg_shape, scheme: str) -> None:
    """Выставить schemeClr-заливку bg-шейпа (accent1 жёлтый / accent2 серый). Если у
    шейпа нет solidFill (у дизайнерских «рядовых» колонок фон прозрачный) — добавляем
    solidFill со schemeClr, чтобы плитки были единообразно залиты."""
    spPr = bg_shape._element.find(qn("p:spPr"))
    if spPr is None:
        return
    sf = spPr.find(qn("a:solidFill"))
    if sf is not None:
        sc = sf.find(qn("a:schemeClr"))
        if sc is not None:
            sc.set("val", scheme)
            for ch in list(sc):        # снять alpha/оттенки, чистый акцент
                sc.remove(ch)
            return
        for ch in list(sf):
            sf.remove(ch)
    else:
        for tag in ("a:noFill", "a:gradFill", "a:blipFill", "a:pattFill"):
            e = spPr.find(qn(tag))
            if e is not None:
                spPr.remove(e)
        sf = etree.SubElement(spPr, qn("a:solidFill"))
        # порядок: solidFill должен идти после geom/ln? python-pptx терпит; вставим сразу
    sc = etree.SubElement(sf, qn("a:schemeClr"))
    sc.set("val", scheme)


# ============================================================================
# Таймлайн clone-elastic (§9). Дизайнерский таймлайн раскладывает этапы вдоль
# ДЕКОРАТИВНОЙ freeform-кривой (custGeom). Репараметризовать эту кривую под
# произвольное N (arc-length размещение пилюль на безье) — вне scope, как и
# внутренности чартов (решение §3Б). Поэтому GENERATE даёт ЧИСТЫЙ ПРЯМОЙ таймлайн:
# клонируем этап {пилюля-дата + описание} до N, ставим пилюли на одну базовую линию,
# описания через одно сверху/снизу, прямой коннектор вместо кривой. Кривые дизайнера
# остаются доступны через FILL для N=3 (#26/#27) и N=4 (#28).
# ============================================================================

Stage = namedtuple("Stage", "pill date desc")


def locate_timeline_stages(slide, gspec) -> list[Stage]:
    """Этапы таймлайна: пилюля = roundRect (на таймлайн-слайде другие roundRect'ы —
    только пилюли), дата = TextBox с центром внутри пилюли ([месяц/№, год]), описание =
    Text Placeholder, чей левый край совпадает с пилюлей. Слева направо = порядок."""
    u = gspec["unit"]
    pill_prst = u.get("pill_prst", "roundRect")
    date_pref = u.get("date_name_prefix", "TextBox")
    desc_name = u.get("desc_name", "Text Placeholder 9")
    pills = [sh for sh in slide.shapes if _prst(sh) == pill_prst]
    pills.sort(key=lambda s: s.left)
    dates = [sh for sh in slide.shapes if sh.has_text_frame and sh.name.startswith(date_pref)]
    descs = [sh for sh in slide.shapes if sh.name == desc_name and sh.has_text_frame]
    stages = []
    for pl in pills:
        pcx, pcy = _in(pl.left) + _in(pl.width) / 2, _in(pl.top) + _in(pl.height) / 2
        date = next((sh for sh in dates
                     if abs(_in(sh.left) + _in(sh.width) / 2 - pcx) < 1.6
                     and abs(_in(sh.top) + _in(sh.height) / 2 - pcy) < 1.5), None)
        desc = min(descs, key=lambda s: abs(_in(s.left) - _in(pl.left)), default=None)
        stages.append(Stage(pl, date, desc))
    return stages


def clone_stage(slide, template: Stage, alloc) -> Stage:
    pill2 = clone_shape_el(slide, template.pill, alloc)
    date2 = clone_shape_el(slide, template.date, alloc) if template.date is not None else None
    desc2 = clone_shape_el(slide, template.desc, alloc) if template.desc is not None else None
    return Stage(pill2, date2, desc2)


def _send_to_back(shape) -> None:
    """Отправить шейп на задний план: переместить его <p:sp> в начало spTree (после
    nvGrpSpPr+grpSpPr, индекс 2) — тогда пилюли/текст рисуются поверх коннектора."""
    sp = shape._element
    spTree = sp.getparent()
    spTree.remove(sp)
    spTree.insert(2, sp)


def _add_connector(slide, x1_in, x2_in, y_in, gspec, prs=None, alloc=None, tdir="eda"):
    """Прямая базовая линия вместо дизайнерской кривой (user-правило 2026-07-02:
    таймлайн 5+ = простая прямая). ПЕРВЫЙ выбор — клон дизайнерского атома
    `atom:connector` с палитры (стиль линии — её), растянутый от первой пилюли
    до последней; фолбэк (атома нет в карте) — прежний тонкий серый прямоугольник.
    В обоих случаях — на задний план (пилюли и текст поверх)."""
    conn = gspec.get("connector", {})
    if not conn.get("draw", True):
        return None
    if prs is not None and alloc is not None:
        try:
            import atoms as A
            amap = A.load_atoms(os.path.join(ROOT, "layouts", tdir, "atoms.json"))
            ln = A.clone_atom(prs, slide, amap, "connector", alloc)
            ln.left, ln.top = Inches(x1_in), Inches(y_in)
            ln.width, ln.height = Inches(x2_in - x1_in), 0
            _send_to_back(ln)
            return ln
        except (KeyError, FileNotFoundError):
            pass                          # роли/карты нет — честный фолбэк ниже
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    h = conn.get("thickness_in", 0.045)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x1_in), Inches(y_in - h / 2),
                                 Inches(x2_in - x1_in), Inches(h))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor.from_string(conn.get("color", "C9C7C6"))
    bar.line.fill.background()
    bar.shadow.inherit = False
    _send_to_back(bar)
    return bar


def _remove_decor_curve(slide, gspec):
    """Убрать дизайнерскую freeform-кривую (декор) — прямой таймлайн её заменяет.
    lxml-remove УЗЛА в памяти (не файла) — deny-правило на rm не затрагивается."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for sh in list(slide.shapes):
        if sh.shape_type == MSO_SHAPE_TYPE.FREEFORM:
            sh._element.getparent().remove(sh._element)


def timeline_slot_width(n: int, gspec: dict) -> float:
    rf, zone = gspec["reflow"], gspec["zone"]
    slot = zone["width_in"] / n
    if slot < rf["min_slot_w_in"]:
        raise GenerateError(
            f"N={n}: slot={slot:.2f}\" < min {rf['min_slot_w_in']}\" — нужен density guard "
            f"(2 ряда / сплит, отложено)"
        )
    return slot


def _timeline_pill_lefts(n: int, pill_w: float, gspec: dict) -> list[float]:
    """Лево пилюль. edge_to_edge (user-fix 2026-07-02): первая на левом краю полосы,
    последняя правым краем на правом, шаг равный; guard по зазору между пилюлями."""
    zone, rf = gspec["zone"], gspec["reflow"]
    if not rf.get("edge_to_edge"):
        slot = timeline_slot_width(n, gspec)
        return [zone["left_in"] + i * slot + slot / 2.0 - pill_w / 2.0 for i in range(n)]
    if n == 1:
        return [zone["left_in"] + (zone["width_in"] - pill_w) / 2.0]
    step = (zone["width_in"] - pill_w) / (n - 1)
    gap = step - pill_w
    if gap < rf.get("min_pill_gap_in", 0.3):
        raise GenerateError(
            f"N={n}: зазор пилюль {gap:.2f}\" < {rf.get('min_pill_gap_in', 0.3)}\" — "
            f"дели таймлайн на два слайда (ребаланс)")
    return [zone["left_in"] + i * step for i in range(n)]


def generate_timeline(prs, gspec, content, base_idx: int, tdir: str = "eda"):
    """Прямой таймлайн из N этапов (clone-elastic). content.items = [{date:[метка,год],
    desc:строка}]. Пилюли на базовой линии, описания через одно сверху/снизу, прямой
    коннектор. Дизайнерская кривая удаляется (её роль берёт коннектор)."""
    slide = prs.slides[base_idx]
    stages = locate_timeline_stages(slide, gspec)
    if not stages:
        raise GenerateError("на базовом слайде не найдено этапов таймлайна (пилюль)")

    items = content.items if not isinstance(content, dict) else content["items"]
    n = len(items)
    zone, rf, u = gspec["zone"], gspec["reflow"], gspec["unit"]
    baseline = rf["baseline_y_in"]
    pill_h = _in(stages[0].pill.height)
    pill_w0 = _in(stages[0].pill.width)
    lefts = _timeline_pill_lefts(n, pill_w0, gspec)      # guard внутри
    slot = (lefts[1] - lefts[0]) if n > 1 else zone["width_in"]
    gap = rf["desc_gap_in"]
    fit_d = gspec.get("fit", {}).get("desc", {})
    date_rl = u.get("date_rel_left_in", 0.5)
    date_rt = u.get("date_rel_top_in", 0.414)

    alloc = next_id_allocator(prs)
    template = stages[-1]                          # этапы одинаковы (без акцента)
    work = list(stages)
    while len(work) < n:
        work.append(clone_stage(slide, template, alloc))
    while len(work) > n:
        extra = work.pop()
        for sh in (extra.pill, extra.date, extra.desc):
            if sh is not None:
                sh._element.getparent().remove(sh._element)

    _remove_decor_curve(slide, gspec)

    pill_w = _in(work[0].pill.width)
    centers = []
    for i, (st, item) in enumerate(zip(work, items)):
        cx = lefts[i] + pill_w / 2.0
        centers.append(cx)
        # пилюля на базовой линии (edge_to_edge: крайние на краях полосы)
        st.pill.left = Inches(lefts[i])
        st.pill.top = Inches(baseline - pill_h / 2.0)
        # дата: центрируем в пилюле; при желании расширяем бокс, чтобы 4-значный год
        # влезал в одну строку (дизайнерский бокс 1.42\" уже года 1.55\" — свой лейаут,
        # можно дать боксу правильную ширину, а не наследовать overflow)
        if st.date is not None:
            if u.get("date_w_in"):
                st.date.width = Inches(u["date_w_in"])
            st.date.left = Inches(cx - _in(st.date.width) / 2.0)
            st.date.top = Inches((baseline - pill_h / 2.0) + date_rt)
            dv = item.get("date") if isinstance(item, dict) else None
            if dv:
                F.fill_lines(st.date, list(dv) if isinstance(dv, (list, tuple)) else [str(dv)])
                F.strip_empty_lines(st.date.text_frame)
        # описание (user-fix 2026-07-02): ШИРИНОЙ С ПИЛЮЛЮ, по её ЛЕВОМУ краю,
        # высота ПОД КОНТЕНТ, текст 32pt regular; через одно сверху/снизу,
        # якорь текста — К ПИЛЮЛЕ (иначе текст в боксах висит на разной высоте)
        if st.desc is not None:
            from pptx.enum.text import MSO_ANCHOR
            from pptx.util import Pt
            dv = item.get("desc") if isinstance(item, dict) else str(item)
            text = dv if isinstance(dv, str) else "\n".join(dv)
            dfont = fit_d.get("font", "YS Text Regular")
            dpt = float(fit_d.get("size_pt", 32))
            dw = pill_w
            try:
                dh = fit.text_height_in(text, dfont, dpt, dw,
                                        line_spacing=1.15) + fit_d.get("pad_h_in", 0.12)
            except fit.FontNotFound:
                dh = rf.get("desc_h_in", 1.6)
            st.desc.width = Inches(dw)
            st.desc.left = Inches(lefts[i])
            st.desc.height = Inches(dh)
            above = (i % 2 == 1)                    # 0-й снизу, дальше через один
            if above:
                st.desc.top = Inches(baseline - pill_h / 2.0 - gap - dh)
                st.desc.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
            else:
                st.desc.top = Inches(baseline + pill_h / 2.0 + gap)
                st.desc.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            F.clear_soft_breaks(st.desc.text_frame)
            F.fill_lines(st.desc, [text])
            F.strip_empty_lines(st.desc.text_frame)
            for par in st.desc.text_frame.paragraphs:   # типографика user-fix явно
                for run in par.runs:
                    run.font.size = Pt(dpt)
                    run.font.bold = False

    if gspec.get("connector", {}).get("full_bleed"):
        x1, x2 = 0.0, _in(prs.slide_width)      # user-fix: линия во всю ширину слайда
    else:
        x1, x2 = centers[0], centers[-1]
    _add_connector(slide, x1, x2, baseline, gspec, prs=prs, alloc=alloc, tdir=tdir)
    return slide


def verify_timeline(slide, gspec, n, *, eps=0.03) -> list[str]:
    """Геометрический composition-check прямого таймлайна (процесс-урок: каждая
    генеративная композиция обязана получить машинный чекер ДО показа): N пилюль
    на ОДНОЙ базовой линии с РАВНЫМ шагом, даты отцентрованы в пилюлях, описания
    чередуются сверху/снизу и не пересекают пилюли, линия тянется от первой
    пилюли до последней по базовой Y."""
    problems = []
    u, rf, zone = gspec["unit"], gspec["reflow"], gspec["zone"]
    baseline = rf["baseline_y_in"]
    pills = sorted((sh for sh in slide.shapes if _prst(sh) == u.get("pill_prst", "roundRect")),
                   key=lambda s: s.left)
    if len(pills) != n:
        problems.append(f"пилюль {len(pills)} ≠ N={n}")
        return problems
    centers = [_in(p.left) + _in(p.width) / 2 for p in pills]
    cys = [_in(p.top) + _in(p.height) / 2 for p in pills]
    if rf.get("edge_to_edge") and n > 1:        # user-fix: крайние пилюли на краях полосы
        if abs(_in(pills[0].left) - zone["left_in"]) > eps:
            problems.append("первая пилюля не на левом краю полосы")
        right = zone["left_in"] + zone["width_in"]
        if abs(_in(pills[-1].left) + _in(pills[-1].width) - right) > eps:
            problems.append("последняя пилюля не на правом краю полосы")
    if max(cys) - min(cys) > eps:
        problems.append(f"пилюли не на одной линии (Δy {max(cys) - min(cys):.2f}\")")
    if abs(cys[0] - baseline) > eps:
        problems.append(f"базовая линия уехала: центр {cys[0]:.2f} ≠ {baseline}")
    steps = [b - a for a, b in zip(centers, centers[1:])]
    if steps and max(steps) - min(steps) > eps:
        problems.append(f"шаг пилюль неравный ({min(steps):.2f}–{max(steps):.2f}\")")
    dates = [sh for sh in slide.shapes if sh.has_text_frame
             and sh.name.startswith(u.get("date_name_prefix", "TextBox"))]
    for i, p in enumerate(pills):
        cx = centers[i]
        d = min(dates, key=lambda s: abs(_in(s.left) + _in(s.width) / 2 - cx), default=None)
        if d is not None and abs(_in(d.left) + _in(d.width) / 2 - cx) > eps:
            problems.append(f"дата этапа {i + 1} не по центру пилюли")
    descs = [sh for sh in slide.shapes if sh.name == u.get("desc_name") and sh.has_text_frame]
    if len(descs) != n:
        problems.append(f"описаний {len(descs)} ≠ N={n}")
    for i, cx in enumerate(centers):
        d = min(descs, key=lambda s: abs(_in(s.left) + _in(s.width) / 2 - cx), default=None)
        if d is None:
            continue
        above = i % 2 == 1
        d_top, d_bot = _in(d.top), _in(d.top) + _in(d.height)
        pill_top = _in(pills[i].top)
        pill_bot = pill_top + _in(pills[i].height)
        if above and d_bot > pill_top + eps:
            problems.append(f"описание этапа {i + 1} должно быть НАД пилюлей")
        if not above and d_top < pill_bot - eps:
            problems.append(f"описание этапа {i + 1} должно быть ПОД пилюлей")
        # якорь текста — к пилюле (сверху=bottom, снизу=top), иначе текст
        # в одинаковых боксах висит на разной высоте
        from pptx.enum.text import MSO_ANCHOR
        want = MSO_ANCHOR.BOTTOM if above else MSO_ANCHOR.TOP
        if d.text_frame.vertical_anchor != want:
            problems.append(f"описание этапа {i + 1}: якорь текста не "
                            f"{'bottom' if above else 'top'} — текст уедет от пилюли")
    line = next((sh for sh in slide.shapes
                 if sh.name.startswith("atom:connector")
                 or (_prst(sh) == "rect" and abs(_in(sh.top) + _in(sh.height) / 2 - baseline) < 0.1
                     and _in(sh.width) > 5)), None)
    if line is None:
        problems.append("базовая линия (коннектор) не найдена")
    elif gspec.get("connector", {}).get("full_bleed"):
        sw = _in(slide.part.package.main_document_part.presentation.slide_width)
        if _in(line.left) > eps or _in(line.width) < sw - 2 * eps:
            problems.append("линия не во всю ширину слайда (full_bleed)")
    else:
        if abs(_in(line.left) - centers[0]) > eps or \
           abs(_in(line.left) + _in(line.width) - centers[-1]) > eps:
            problems.append("линия не тянется от первой пилюли до последней")
    # user-fix: описание шириной с пилюлю, по её ЛЕВОМУ краю, зазор = desc_gap
    for i, p in enumerate(pills):
        d = min(descs, key=lambda s: abs(_in(s.left) + _in(s.width) / 2 - centers[i]),
                default=None)
        if d is None:
            continue
        if abs(_in(d.left) - _in(p.left)) > eps:
            problems.append(f"описание этапа {i + 1} не по левому краю пилюли")
        if abs(_in(d.width) - _in(p.width)) > 0.1:
            problems.append(f"описание этапа {i + 1} не шириной с пилюлю")
        gap_want = rf["desc_gap_in"]
        if i % 2 == 1:                     # сверху: низ бокса на gap выше пилюли
            got = _in(p.top) - (_in(d.top) + _in(d.height))
        else:                              # снизу: верх бокса на gap ниже пилюли
            got = _in(d.top) - (_in(p.top) + _in(p.height))
        if abs(got - gap_want) > eps:
            problems.append(f"описание этапа {i + 1}: зазор {got:.2f}\" ≠ {gap_want}\"")
    return problems
