#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""charts.py — замена ДАННЫХ нативного чарта при нетронутом дизайне.

python-pptx `chart.replace_data` переписывает категории/значения/имена серий
и встроенную xlsx-книгу; тип чарта, цвета серий и оформление осей остаются
дизайнерскими (при числе серий ≤ дизайнерского; лишние серии клонируют стиль
последней — честный флаг в отчёте).

Ось категорий: у линейного чарта Еды (#35/атом) ось ДАТНАЯ (`c:dateAx`,
formatCode «[$-419]mmmm» — русские месяцы) → категории deckspec для него
обязаны быть датами ISO («2026-03-01»); текст на датной оси даст мусор —
громкая ошибка, не молчаливая порча. Обычные COLUMN/BAR (#40/#41) — текстовые
категории как есть.

После замены — readback-верификация тем же python-pptx: категории/серии/значения
читаются назад и сравниваются с заказанными (мера, не глаз).
"""
from __future__ import annotations

import datetime as _dt

from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn


class ChartError(Exception):
    pass


def find_chart_frame(slide, name: str):
    """graphicFrame с чартом по имени шейпа."""
    for sh in slide.shapes:
        if sh.has_chart and sh.name == name:
            return sh
    return None


def _has_date_axis(chart) -> bool:
    return chart._chartSpace.find(f".//{qn('c:dateAx')}") is not None


def _parse_date(v):
    if isinstance(v, (_dt.date, _dt.datetime)):
        return v
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y"):
        try:
            return _dt.datetime.strptime(str(v), fmt).date()
        except ValueError:
            continue
    raise ChartError(
        f"категория «{v}»: у чарта ДАТНАЯ ось — категории должны быть датами "
        f"ISO (2026-03-01 / 2026-03 / 2026), не текстом")


def _norm_series(series) -> list[tuple[str, list[float]]]:
    """{имя: [числа]} ИЛИ [[имя, [числа]], ...] → пары; значения — числа/None."""
    pairs = list(series.items()) if isinstance(series, dict) else \
        [(s[0], s[1]) for s in series]
    out = []
    for name, vals in pairs:
        clean = [None if v is None else float(v) for v in vals]
        out.append((str(name), clean))
    return out


def replace_chart_data(chart, spec: dict) -> dict:
    """Заменить данные чарта: spec = {"categories": [...], "series": {...}}."""
    cats = spec.get("categories") or []
    sers = _norm_series(spec.get("series") or {})
    if not cats or not sers:
        raise ChartError("chart-спека без categories/series")
    for name, vals in sers:
        if len(vals) != len(cats):
            raise ChartError(f"серия «{name}»: {len(vals)} значений при "
                             f"{len(cats)} категориях")

    date_ax = _has_date_axis(chart)
    cat_fmt = None
    if date_ax:
        cats = [_parse_date(c) for c in cats]
        # дизайнерский формат категорий (напр. [$-419]mmmm — русские месяцы):
        # ось sourceLinked=1 берёт его из кэша c:cat, а replace_data затирает
        # кэш ISO-датами → без возврата формата ось покажет «2026-01-01»
        fc = chart._chartSpace.find(
            f".//{qn('c:ser')}/{qn('c:cat')}//{qn('c:formatCode')}")
        cat_fmt = fc.text if fc is not None else None

    n_designer = len(chart.plots[0].series)
    data = CategoryChartData()
    data.categories = cats
    if cat_fmt:
        data.categories.number_format = cat_fmt
    for name, vals in sers:
        data.add_series(name, tuple(vals))
    chart.replace_data(data)
    if str(chart.chart_type).startswith("LINE"):
        _tune_line_labels(chart, sers)

    # readback-верификация: то, что читается назад, == заказанному
    plot = chart.plots[0]
    got = [(s.name, [None if v is None else float(v) for v in s.values])
           for s in plot.series]
    want = [(n, v) for n, v in sers]
    problems = []
    if len(got) != len(want):
        problems.append(f"серий {len(got)} ≠ {len(want)}")
    for (gn, gv), (wn, wv) in zip(got, want):
        if gn != wn:
            problems.append(f"имя серии «{gn}» ≠ «{wn}»")
        if gv != wv:
            problems.append(f"значения «{wn}» разошлись: {gv} ≠ {wv}")
    n_cats = len(list(plot.categories))
    if n_cats != len(cats):
        problems.append(f"категорий {n_cats} ≠ {len(cats)}")

    rep = {"replaced": True, "series": len(sers), "points": len(cats),
           "date_axis": date_ax, "designer_series": n_designer,
           "verify_ok": not problems}
    if problems:
        rep["problems"] = problems
    if len(sers) > n_designer:
        rep["style_cloned"] = (f"серий {len(sers)} > дизайнерских {n_designer} — "
                               f"лишние наследуют стиль последней")
    return rep


def _tune_line_labels(chart, sers) -> None:
    """Метки данных LINE-чарта (user-fix 2026-07-03, «цифры слиплись»):
    (1) пер-точечные <c:dLbl>-переопределения донора ВЫЧИЩАЮТСЯ — их manualLayout
    привязан к СТАРЫМ значениям и с новыми данными метки слипаются/уезжают;
    (2) позиция метки СЕРИЙНО: нижняя серия (наименьшее финальное значение) —
    ПОД линией (b), остальные — НАД (t). Ровно так пользователь развела метки
    руками (Код=t, Тексты=t, Анализ=b)."""
    from lxml import etree
    finals = {name: next((v for v in reversed(vals) if v is not None), 0)
              for name, vals in sers}
    if not finals:
        return
    lowest = min(finals, key=lambda k: finals[k])
    for ser in chart._chartSpace.findall(f".//{qn('c:ser')}"):
        nm = ser.find(f".//{qn('c:tx')}//{qn('c:v')}")
        name = nm.text if nm is not None else ""
        dlbls = ser.find(qn("c:dLbls"))
        if dlbls is None:
            continue
        for dlbl in dlbls.findall(qn("c:dLbl")):
            dlbls.remove(dlbl)
        want = "b" if name == lowest else "t"
        pos = dlbls.find(qn("c:dLblPos"))
        if pos is None:
            pos = etree.SubElement(dlbls, qn("c:dLblPos"))
            # схема CT_DLbls: dLblPos идёт ПЕРЕД show*-флагами
            first_show = next((ch for ch in dlbls
                               if etree.QName(ch).localname.startswith("show")), None)
            if first_show is not None:
                first_show.addprevious(pos)
        pos.set("val", want)


def replace_chart_slots(slide, variant: dict, charts: dict) -> dict:
    """Все chart-слоты варианта: charts = {слот: {categories, series}}.
    Слот без значения — passthrough (дизайнерские данные остаются)."""
    report = {}
    for name, slot in (variant.get("chart_slots") or {}).items():
        spec = charts.get(name)
        if not spec:
            report[name] = {"replaced": False, "passthrough": True}
            continue
        frame = find_chart_frame(slide, slot["match_name"])
        if frame is None:
            report[name] = {"replaced": False,
                            "error": f"чарт «{slot['match_name']}» не найден"}
            continue
        try:
            report[name] = replace_chart_data(frame.chart, spec)
        except ChartError as e:
            report[name] = {"replaced": False, "error": str(e)}
    return report
