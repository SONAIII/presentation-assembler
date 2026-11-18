#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""route.py — роутер «примерь готовый → иначе собери» (архитектура §3).

Чистая логика решения: контент → спека → FILL (есть готовый слайд под N и текст
влезает) либо GENERATE (нет готового / не влезает). Исполнение делегируется
lib.fill (FILL) и lib.generate (GENERATE). Замеры — через lib.fit.

Порядок внутри решения (как в §3): fill как есть → fill со shrink в пределах
min_pt → generate.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, field

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import anchor  # noqa: E402
import fit  # noqa: E402
import generate as gen  # noqa: E402


@dataclass
class Content:
    form: str                       # "metrics" | "title" | "divider" | "toc" | "closing" | "contacts"
    items: list                     # [{"value": "32%", "label": "..."}], len = N (контентный тир)
    title: str | None = None
    subtitle: str | None = None
    slots: dict | None = None       # структурный тир: произвольные именованные слоты
                                    # {"section_label": "Раздел 1", "items": [...], "author": "..."}


@dataclass
class Route:
    action: str                     # "FILL" | "GENERATE"
    spec: dict
    n: int
    reason: str
    ready_slide_index: int | None = None
    fit_report: dict = field(default_factory=dict)


def load_spec(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def fit_check_ready(content: Content, spec: dict, slide_index: int, prs) -> dict:
    """Влезает ли контент в готовый слайд на дизайнерских кеглях (с допустимым
    shrink)? Значение wrap=off выровнено вправо — растёт влево, к подписи не лезет,
    поэтому проверяем, что ширина значения не больше допустимой (до зоны/фото)."""
    cu = spec["card_unit"]
    vfont, vsize = cu["value"]["font"], cu["value"]["size_pt"]
    min_pt = spec["fit"]["value"]["min_pt"]
    max_val_w = cu["value"]["rel_left_in"] + cu["value"]["width_in"] + 1.0
    lfont, lsize = cu["label"]["font"], cu["label"]["size_pt"]
    lwidth = cu["label"]["width_in"]
    soft_lines = spec["fit"]["label"]["max_lines_soft"]

    slots = []
    ok = True
    for it in content.items:
        try:
            vw = fit.text_width_in(str(it["value"]), vfont, vsize)
            size_used = vsize
            shrunk = False
            if vw > max_val_w:
                size_used = fit.shrink_to_fit(str(it["value"]), vfont, vsize, max_val_w,
                                              wrap=False, min_size_pt=min_pt)
                shrunk = True
                vw = fit.text_width_in(str(it["value"]), vfont, size_used)
            value_fits = vw <= max_val_w + 0.02
            lines = fit.line_count(str(it["label"]), lfont, lsize, lwidth)
            label_ok = lines <= soft_lines
            slots.append({"value": it["value"], "value_w_in": round(vw, 2),
                          "size_used": size_used, "shrunk": shrunk,
                          "value_fits": value_fits, "label_lines": lines, "label_ok": label_ok})
            if not value_fits:
                ok = False
        except fit.FontNotFound as e:
            slots.append({"value": it["value"], "error": f"font not found: {e}"})
            ok = False
    return {"ok": ok, "slots": slots,
            "shrink_needed": any(s.get("shrunk") for s in slots)}


def decide(content: Content, spec: dict, prs) -> Route:
    """FILL, если есть готовый слайд ровно под N и контент влезает; иначе GENERATE."""
    n = len(content.items)
    ready = [r for r in spec.get("ready_slides", []) if r["cards"] == n]
    if ready:
        idx = anchor.resolve_entry(ready[0], prs)   # §3Б: по сигнатуре, не по индексу
        fr = fit_check_ready(content, spec, idx, prs)
        if fr["ok"]:
            reason = "готовый %d-карточный слайд подходит%s" % (
                n, " (со shrink значения)" if fr["shrink_needed"] else "")
            return Route("FILL", spec, n, reason, ready_slide_index=idx, fit_report=fr)
        return Route("GENERATE", spec, n, f"готовый {n}-карточный есть, но контент не влез",
                     fit_report=fr)
    return Route("GENERATE", spec, n, f"нет готового {n}-карточного слайда")


# ------------------------------------------ контент-семейство с генеративной веткой


def _norm_block(b) -> list[str]:
    """Блок из deckspec: {name, desc} | [имя, описание, ...] → список абзацев."""
    if isinstance(b, dict):
        paras = [b.get("name") or b.get("title") or ""]
        d = b.get("desc") or b.get("body") or ""
        paras += d if isinstance(d, list) else [d]
    else:
        paras = list(b) if isinstance(b, (list, tuple)) else [str(b)]
    return [str(p) for p in paras if str(p).strip()]


def decide_columns(content: Content, spec: dict, prs, *, variant_name: str | None = None) -> Route:
    """Честная развилка «примерь → иначе собери» (§3; редирект session 6: счёт-совпал
    ≠ подходит). FILL дизайнерского слайда ТОЛЬКО когда контент его ЗАПОЛНЯЕТ:
      * N == ёмкости варианта (счёт блоков на живом слайде, не в спеке);
      * всё влезает на дизайнерском кегле (ratio ≤ fit_tolerance);
      * fill-ratio ≥ fill_ratio_min — короткие тезисы в высоких био-карточках
        оставляют карточки пустыми → GENERATE компактных плиток.
    Калибровка порогов: собственный контент дизайнера меряется 0.79–1.03 (все 4
    варианта columns) → min 0.5, tolerance 1.05. Явный variant_name = осознанный
    override агента → FILL без развилки."""
    if variant_name:
        return decide_structural(content, spec, prs, variant_name=variant_name)
    import fill as F  # локально: route ← fill только для счёта/замера блоков

    raw = (getattr(content, "slots", None) or {}).get("blocks") or content.items
    blocks = [_norm_block(b) for b in raw]
    n = len(blocks)
    fc = spec.get("fit_check", {})
    rmin = fc.get("fill_ratio_min", 0.5)
    rmax = fc.get("fit_tolerance", 1.05)
    ls = fc.get("line_spacing", 1.15)

    lo, hi = spec.get("match", {}).get("count", [1, 999])
    if not (lo <= n <= hi):
        return Route("GENERATE", spec, n,
                     f"нет готового слайда под N={n} (готовые: {lo}–{hi})")

    checked = []
    for v in sorted(spec["variants"], key=lambda x: not x.get("preferred")):
        idx = anchor.resolve_entry(v, prs)
        slot = v["slots"]["blocks"]
        shs = F.find_shapes_by_name(prs.slides[idx], slot["match_name"],
                                    sort=slot.get("sort", "doc"),
                                    skip_regex=slot.get("skip_regex"),
                                    max_w_in=slot.get("max_w_in"))
        if len(shs) != n:
            checked.append(f"{v['name']}: ёмкость {len(shs)} ≠ N={n}")
            continue
        try:
            ratios = [sum(fit.text_height_in(p, slot["font"], slot["size_pt"],
                                             slot["box_w_in"], line_spacing=ls)
                          for p in b) / slot["box_h_in"] for b in blocks]
        except fit.FontNotFound as e:
            checked.append(f"{v['name']}: шрифт не найден ({e})")
            continue
        if max(ratios) > rmax:
            checked.append(f"{v['name']}: не влезает (ratio {max(ratios):.2f} > {rmax})")
            continue
        if min(ratios) < rmin:
            checked.append(f"{v['name']}: fill-ratio {min(ratios):.2f} < {rmin} — "
                           f"карточки останутся пустыми")
            continue
        return Route("FILL", spec, n,
                     f"вариант '{v['name']}' заполнен честно "
                     f"(fill-ratio {min(ratios):.2f}–{max(ratios):.2f})",
                     ready_slide_index=idx,
                     fit_report={"variant": v,
                                 "ratios": [round(r, 2) for r in ratios]})
    return Route("GENERATE", spec, n,
                 "готовые варианты не подходят → компактные плитки "
                 f"({'; '.join(checked)})")


def decide_timeline(content: Content, spec: dict, prs, *, variant_name: str | None = None) -> Route:
    """Роутер таймлайна — по ВМЕСТИМОСТИ (у вариантов фиксированные dateN-слоты):
    N=3 → t3 (preferred), N=4 → t4; прочие N в generate.count → GENERATE прямой
    линии из атома connector (user-правило 2026-07-02: таймлайн 5+ = простая
    прямая); за пределами — честная ошибка (дели на слайды, ребаланс)."""
    if variant_name:
        return decide_structural(content, spec, prs, variant_name=variant_name)
    stages = (getattr(content, "slots", None) or {}).get("stages") or content.items
    n = len(stages)
    if n == 0:
        raise ValueError("timeline ждёт slots.stages = [{date: [метка, год], "
                         "desc: строка}, ...] — этапы не переданы")
    def capacity(v):
        # ёмкость варианта: явный "steps" (Городские — номера зашиты в овалы
        # донора, слот один repeated "descs") ИЛИ число dateN-слотов (Еда)
        return v.get("steps") or sum(1 for k in v["slots"] if k.startswith("date"))
    cand = next((v for v in sorted(spec["variants"], key=lambda x: not x.get("preferred"))
                 if capacity(v) == n), None)
    if cand is not None:
        idx = anchor.resolve_entry(cand, prs)
        return Route("FILL", spec, n,
                     f"готовый {n}-точечный таймлайн '{cand['name']}' (дизайнерская кривая)",
                     ready_slide_index=idx, fit_report={"variant": cand})
    gen = spec.get("generate")
    if gen:
        lo, hi = gen.get("count", [2, 6])
        if lo <= n <= hi:
            return Route("GENERATE", spec, n,
                         f"нет готового таймлайна на N={n} → прямая линия (атом connector)")
    caps = "/".join(str(c) for c in sorted({capacity(v) for v in spec["variants"]} - {0}))
    extra = (f", прямая линия — {gen.get('count', [2, 6])[0]}–{gen.get('count', [2, 6])[1]}"
             if gen else " (клон-генератора у шаблона нет)")
    raise ValueError(
        f"таймлайн N={n} не собрать: готовые — {caps} точки{extra}; "
        f"реворди под готовый N, дели на два таймлайн-слайда (ребаланс, напр. "
        f"{n}→{(n + 1) // 2}+{n // 2}) или custom")


# ----------------------------------------------------- структурный тир (всегда FILL)


def decide_structural(content: Content, fillmap: dict, prs, *, variant_name: str | None = None) -> Route:
    """Структурный тир (обложка/раздел/содержание/закрытие): GENERATE не бывает —
    всегда FILL готового дизайнерского слайда. Выбираем вариант (явный variant_name
    → preferred → первый); вариант кладём в fit_report для исполнителя FILL."""
    variants = fillmap.get("variants", [])
    if not variants:
        raise ValueError(f"fill-map '{fillmap.get('id')}' без variants")
    if variant_name:
        chosen = next((v for v in variants if v["name"] == variant_name), None)
        if chosen is None:
            raise ValueError(f"вариант '{variant_name}' не найден в '{fillmap.get('id')}'")
    else:
        chosen = next((v for v in variants if v.get("preferred")), variants[0])
    idx = anchor.resolve_entry(chosen, prs)          # §3Б: по сигнатуре, не по индексу
    return Route(
        "FILL", fillmap, 1,
        f"структурный {fillmap.get('family')} → вариант '{chosen['name']}' (slide {idx + 1})",
        ready_slide_index=idx,
        fit_report={"variant": chosen},
    )


def route_slide(content: Content, artifact: dict, prs, **kw) -> Route:
    """Верхняя развилка по `kind` (§3 + тиры): спека с fit_check
    (columns/takeaways) → честная развилка decide_columns: FILL, а не заполнил
    донора — GENERATE-вердикт, который build_deck превращает в честную ошибку
    «реворд или custom» (плитки выпилены 0.21.0; раньше диспатч требовал ещё и
    generate-секцию — с её удалением из пака условие сжато до fit_check, иначе
    несовпавший контент молча уходил в структурный FILL с потерей блоков);
    прочие structural/content_fill → всегда FILL по fill-map (§3Б); иначе
    контентный роутер metrics («примерь → иначе собери»)."""
    if artifact.get("router") == "timeline":
        return decide_timeline(content, artifact, prs, **kw)
    if artifact.get("kind") in ("structural", "content_fill"):
        if artifact.get("fit_check"):
            return decide_columns(content, artifact, prs, **kw)
        return decide_structural(content, artifact, prs, **kw)
    return decide(content, artifact, prs)
