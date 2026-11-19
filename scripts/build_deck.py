#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_deck.py — entrypoint скилла: deckspec.json → готовый .pptx + гейт.

Дек = последовательность слайдов из СЕМЕЙСТВ шаблона (references/families-<tdir>.md).
Каждый слайд: FILL готового дизайнерского (роутер по anchor-сигнатуре, слоты по
fill-map) ЛИБО клон-GENERATE из дизайнерских юнитов донора (прямой таймлайн
N=5–6, стат-сетка метрик N≠3 — роутеры включают их сами). Плиточная генерация
из атомов ВЫПИЛЕНА (0.21.0, решение владельца 2026-07-23): family=reasons и
эскалации columns/takeaways в плитки дают честную ошибку с рецептом.

deckspec.json:
{
  "template": "еда",
  "slides": [
    {"family": "cover", "variant": "photo_title_bottom",
     "slots": {"title": "...", "author": "..."},
     "photos": {"photo": "Бьюти_01.jpg"}},        ← имя из пула Я.Диска ИЛИ путь
    {"family": "toc", "slots": {"items": ["...", "..."], "active_index": 0}},
    {"family": "two_values", "slots": {...}},
    {"family": "metrics", "slots": {"heading": "...", "subheading": "...",
     "items": [{"value": "32%", "label": "..."}, ...]}},   ← N=3 FILL, иначе GENERATE
    {"family": "timeline", "slots": {"heading": "...",
     "stages": [{"date": ["Мар", "2023"], "desc": "..."}, ...]}},
                                    ← N=3/4 FILL кривой, N=5–6 GENERATE прямой
    {"family": "linechart", "slots": {"heading": "...", "comment": ["...", "..."]},
     "charts": {"chart": {"categories": ["2026-01", "2026-02", ...],
                          "series": {"Имя серии": [числа, ...]}}}},
                                    ← данные чарта РЕДАКТИРУЮТСЯ (SKILL.md «Чарты»)
    {"family": "custom", "html": "custom_market_map.html", "html_slide": 0},
                                    ← HTML-хэтч (вторая ветка роутера
                                      fill → custom): слайд по дизайн-системе
                                      Еды (assets/eda/deck.css в скилле), конвертируется
                                      в шейпы на хроме лэйаута; references/custom-html.md
    {"family": "closing", "slots": {"title": "Обсудим раскатку?", ...},
     "notes": "Полный текст для устного рассказа."}
                                    ← заметки спикера: строка или список абзацев;
                                      на любом слайде, в канву не попадают
  ]
}

Ключи deckspec (свод; контракты слотов по семействам — references/families-<tdir>.md):
  верхний уровень:
    template          имя шаблона, дефолт «еда» → TDIR_ALIASES → LAYOUT_DIRS
    slides            [обяз.] список записей-слайдов
    regression_smoke  тест-флаг: ослабляет жёсткие гварды A13/A15 до ⚠;
                      в боевых деках БЫТЬ НЕ ДОЛЖЕН (только смоуки TESTING.md)
  запись slides[]:
    family    [обяз.] семейство (custom→HTML-хэтч; иначе fill-map
              layouts/<t>/<family>.json; toc→agenda.json)
    variant   явное имя варианта (override роутера)
    slots     dict контент-слотов (title/heading, subtitle/subheading, items,
              blocks, stages, legend, active_index; ручки accent/bullets/numbered=false)
    charts    dict chart_slot → {categories, series:{имя:[числа]}} (replace_data)
    photos    dict image_slot → имя из пула Я.Диска ИЛИ путь (photo-swap)
    notes     заметки спикера: строка или список абзацев (в канву не попадают)
    html      [custom] путь к HTML-слайду; html_slide — индекс в файле (дефолт 0)
    items     [metrics] список {value,label} — фолбэк slots.items
    heading   заголовок (наравне со slots.heading)

Сборка ДВУХФАЗНАЯ через клоны (2026-07-02, снят v0-запрет повторов):
  фаза 1 — резолв всех записей по ПРИСТИННОМУ шаблону (якоря не видят клонов);
  фаза 2 — на каждую запись безопасный клон дизайнерского слайда (lib/clone.py:
  ремап creationId/a16/modId, rels с сохранением rId) → заливка КЛОНА;
  keep_ordered оставляет клоны, все оригиналы + палитра уходят.
Повторы семейства разрешены (агенда-per-section: toc ×M c active_index).
Ограничители:
  * слайд с chart/OLE-партом клонируется не более 1 раза на дек (crash-check
    known_issues #6: шаринг chart-партов Mac PP не прощает, дубль не лечится);
  * ≥3 слайдов одного семейства подряд — предупреждение A10 (монотонность);
  * данные чартов linechart/charts_compare/charts_duo редактируются через
    "charts" (тип чарта, цвета серий, оси — дизайнерские); мини-чарты
    think-cell (#30) и таблицы — passthrough, заполняется текст вокруг;
  * клон-GENERATE: прямой таймлайн (N 5–6), стат-сетка метрик (N≠3) — роутеры
    включают их сами, причину печатают; контент, не заполнивший донора
    columns/takeaways, — честная ошибка «реворд или custom» (плиток нет, 0.21.0).

Гейт после сборки: crash-check → lint → composition-check (для generate) →
стилевой чекер агенды (внутри fill_agenda_slot, FAIL = дек не сохраняется).
Рендер и глазную проверку агент делает отдельно (references/quality-gates.md).

Быстрая проверка спеки без сборки: `--dry-run` — фаза резолва + гварды
(chart-guard, A10), печать плана роутера; файл не пишется.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from collections import Counter

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
REPO = SKILL_DIR       # папка скилла самодостаточна: lib/layouts/Templates внутри
sys.path.insert(0, os.path.join(REPO, "lib"))

from pptx import Presentation                   # noqa: E402

import anchor                                   # noqa: E402
import charts as CHRT                           # noqa: E402
import clone as CL                              # noqa: E402
import fill as F                                # noqa: E402
import generate as GEN                          # noqa: E402
import html2pptx as H2P                         # noqa: E402
import images as IMG                            # noqa: E402
import lint as LINT                             # noqa: E402
import route                                    # noqa: E402
import template_profile as TP                    # noqa: E402  (профиль кастома по tdir)

sys.path.insert(0, os.path.join(REPO, "scripts"))
import fetch_photo                              # noqa: E402  (пул фото Я.Диска)
import fetch_template                           # noqa: E402  (шаблон с Я.Диска по пину)

RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
LAYOUT_DIRS = {"eda": os.path.join(REPO, "layouts/eda"),
               "gor-services": os.path.join(REPO, "layouts/gor-services"),
               "ridetech-khural": os.path.join(REPO, "layouts/ridetech-khural"),
               "committee": os.path.join(REPO, "layouts/committee")}
TDIR_ALIASES = fetch_template.TDIR_ALIASES   # единый источник — scripts/fetch_template.py (не дублируем)


def template_path(tdir: str) -> str:
    """Абсолютный путь к .pptx шаблона: локальный Templates/ (дев-режим) →
    кэш → скачивание с Я.Диска по пину из Templates/manifest.json."""
    return fetch_template.resolve_template(tdir)


def theme_latin_fonts(prs):
    """(major, minor) latin-семейства темы мастера — для резолва наследуемых/темовых
    шрифтов ранов (+mj-lt/+mn-lt/None) в фактические имена под bold-карту (F.set_theme_fonts).
    Тема недоступна/нечитаема → (None, None): резолв выключен, поведение прежнее."""
    import re
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as _RT
        theme = prs.slide_masters[0].part.part_related_by(_RT.THEME)
        xml = theme.blob.decode("utf-8", "ignore")
        mj = re.search(r'<a:majorFont>.*?<a:latin typeface="([^"]+)"', xml, re.S)
        mn = re.search(r'<a:minorFont>.*?<a:latin typeface="([^"]+)"', xml, re.S)
        return (mj.group(1) if mj else None), (mn.group(1) if mn else None)
    except Exception:
        return None, None
FAMILY_FILES = {"toc": "agenda.json"}          # алиасы семейство → файл fill-map


class DeckError(Exception):
    pass


def known_families(tdir: str) -> list[str]:
    """Меню семейств шаблона: fill-maps из layouts/."""
    skip = {"_canvas", "_inventory", "brand-tokens", "atoms", "custom-profile"}
    alias = {v: k for k, v in FAMILY_FILES.items()}      # agenda.json → toc
    fams = {alias.get(f, f[:-5]) for f in os.listdir(LAYOUT_DIRS[tdir])
            if f.endswith(".json") and f[:-5] not in skip}
    return sorted(fams)


# --- типографика контента (brand-rules §14.2 Городских + правило точек user 2026-07-02)
import re as _re

_SHORT_WORDS = ("в к с о у и а я на по до из за от во со ко об не ни но бы же ли да "
                "для без над под про при обо изо или что как чем тем").split()
_NBSP_RE = _re.compile(r"(?i)(?<![\wЀ-ӿ-])(" + "|".join(_SHORT_WORDS) + r") ")


def _typo(s: str) -> str:
    """(1) NBSP после коротких служебных слов — никаких висячих предлогов;
    (2) самостоятельный абзац не заканчивается точкой (внутренние точки живут)."""
    prev = None
    while prev != s:                     # «и в Москве» — каскад коротких слов
        prev = s
        s = _NBSP_RE.sub(lambda m: m.group(1) + " ", s)
    t = s.rstrip()
    if t.endswith(".") and not t.endswith(("..", "…")):
        s = t[:-1]
    return s


# значения-пути/имена файлов: NBSP после «о»/«в» в имени папки ломает резолв
# (фидбек переноса HTML 2026-08-10 — «пришлось уносить файлы в папку без пробелов»)
_TYPO_SKIP_KEYS = {"html", "photos"}


def typo_walk(obj):
    if isinstance(obj, str):
        return _typo(obj)
    if isinstance(obj, list):
        return [typo_walk(x) for x in obj]
    if isinstance(obj, dict):
        return {k: (v if k in _TYPO_SKIP_KEYS else typo_walk(v))
                for k, v in obj.items()}
    return obj


def keep_ordered(prs, indices: list[int]) -> None:
    """Оставить слайды indices (0-based) в заданном порядке, прочие отвязать
    (drop_rel ДО удаления — иначе пакет бьётся; палитра атомов уходит здесь же)."""
    lst = prs.slides._sldIdLst
    ids = list(lst)
    keep = set(indices)
    for i, sld in enumerate(ids):
        if i not in keep:
            prs.part.drop_rel(sld.get(RID))
        lst.remove(sld)
    for i in indices:
        lst.append(ids[i])


def _norm_stage(s) -> dict:
    """Этап таймлайна: {date: [метка,год]|строка, desc} | [дата, описание] →
    {"date": [строки], "desc": строка}."""
    if isinstance(s, dict):
        d = s.get("date")
        date = [str(x) for x in d] if isinstance(d, (list, tuple)) else [str(d or "")]
        return {"date": date, "desc": str(s.get("desc") or "")}
    d = s[0]
    return {"date": [str(x) for x in d] if isinstance(d, (list, tuple)) else [str(d)],
            "desc": str(s[1]) if len(s) > 1 else ""}


def resolve_photo(value: str, tdir: str) -> str:
    """Значение фото-слота deckspec → локальный путь: готовый путь как есть,
    иначе имя файла из одобренного пула шаблона (скачивается с кешем)."""
    if os.path.sep in value or os.path.exists(value):
        return value
    try:
        return fetch_photo.fetch(tdir, value)
    except Exception as e:
        raise DeckError(f"фото «{value}» не достать из пула: {e}") from e


def set_speaker_notes(prs, idx: int, notes) -> None:
    """Заметки спикера на построенный слайд (deckspec "notes": строка | [абзацы]).
    Клон рождается БЕЗ notesSlide (clone.py, паттерн #1) — python-pptx создаёт
    свежий парт от родного notesMaster шаблона, шаринга партов не возникает.
    Текст кладём СЫРЫМ (мимо _typo): заметки — проза для спикера, не канва."""
    text = ("\n".join(str(x) for x in notes)
            if isinstance(notes, (list, tuple)) else str(notes))
    if text.strip():
        prs.slides[idx].notes_slide.notes_text_frame.text = text


def fill_heading(slide, text, *, name="Title 1"):
    sh = F.find_shape_by_name(slide, name)
    if sh is not None and text:
        F.clear_soft_breaks(sh.text_frame)
        F.fill_lines(sh, [text])
        F.strip_empty_lines(sh.text_frame)


def resolve_entry_plan(prs, entry: dict, tdir: str) -> dict:
    """Фаза 1: ЧИСТЫЙ резолв записи deckspec по пристинному шаблону — какой
    дизайнерский слайд источник и как его заливать. Никаких мутаций prs."""
    family = entry["family"]

    if family == "custom":
        # третья ветка роутера (fill → structure → custom): слайд свёрстан
        # по дизайн-системе Еды скилла (assets/eda/deck.css), обмеряется
        # Chrome'ом и собирается нативными шейпами на хроме (lib/html2pptx.py)
        html = entry.get("html")
        if not html:
            raise DeckError("custom: нужен ключ \"html\" — путь к HTML-слайду "
                            "по дизайн-системе Еды (references/custom-html.md)")
        html_path = os.path.abspath(html)
        if not os.path.exists(html_path):
            raise DeckError(f"custom: HTML не найден: {html_path}")
        try:
            data = H2P.measure_html(html_path, int(entry.get("html_slide", 0)),
                                    SKILL_DIR, tdir)
        except H2P.Html2PptxError as e:
            raise DeckError(f"custom: {e}") from e
        return {"mode": "CUSTOM", "family": family, "entry": entry,
                "src": None, "measure": data,
                "routed": f"HTML-хэтч: {os.path.basename(html_path)}"
                          f"[{entry.get('html_slide', 0)}]"}

    if entry.get("generate") or family == "reasons":
        raise DeckError(
            "плиточная генерация из атомов выпилена (0.21.0, решение владельца "
            "2026-07-23: ставка — дизайнерские слайды + custom). Тезисам выбери "
            "готовое семейство (takeaways/text_grid/mixed/bodytext/columns — "
            "карта в references/deck-archetypes.md §0), реворди под донора или "
            "сверстай custom (references/custom-html.md)")

    fname = FAMILY_FILES.get(family, f"{family}.json")
    spec_path = os.path.join(LAYOUT_DIRS[tdir], fname)
    if not os.path.exists(spec_path):
        raise DeckError(f"нет семейства {family!r}; доступны: "
                        f"{', '.join(known_families(tdir))} "
                        f"(каталог: references/families-{tdir}.md)")
    fillmap = route.load_spec(spec_path)
    form = fillmap.get("match", {}).get("form") or family
    slots = entry.get("slots") or {}
    if "blocks" in slots:                # {name,desc} | [имя, ...] → список абзацев
        slots = dict(slots, blocks=[route._norm_block(b) for b in slots["blocks"]])
    if "stages" in slots:                # таймлайн: [{date:[метка,год], desc}, ...]
        slots = dict(slots, stages=[_norm_stage(s) for s in slots["stages"]])
    if fillmap.get("card_unit") and not fillmap.get("variants"):
        # стат-сетка метрик (спека §9 этап 1, единственная без variants):
        # проверенная машинка demo — decide по N + глиф-фиту значений,
        # FILL готовой 3-карточной сетки / clone-elastic на другие N
        items = slots.get("items") or entry.get("items") or []
        if not items:
            raise DeckError(f"{family}: нужен slots.items = "
                            f"[{{\"value\": ..., \"label\": ...}}, ...]")
        # значение = число (3–4 знака); «6 городов» / «~35 мин» — единица в
        # значении, из-за неё группа ужимается и вылетает (user-фидбек 2026-07-07)
        for it in items:
            val = str(it.get("value", "")) if isinstance(it, dict) else str(it)
            if _re.search(r"\d[\s ]+[А-Яа-яA-Za-zёЁ]", val):
                print(f"  ⚠ {family}: значение {val!r} содержит единицу/слово — "
                      f"оставь в значении число, слово перенеси в подпись "
                      f"(«6 городов» → значение «6», подпись «городов: …»)")
        content = route.Content(form, items,
                                title=slots.get("heading") or slots.get("title"),
                                subtitle=slots.get("subheading") or slots.get("subtitle"),
                                slots=slots)
        r = route.decide(content, fillmap, prs)
        src = (r.ready_slide_index if r.action == "FILL"
               else anchor.resolve_entry(fillmap["ready_slides"][0], prs))
        return {"mode": r.action, "generator": "statgrid", "family": family,
                "entry": entry, "src": src, "spec": fillmap,
                "content": content, "reason": r.reason, "routed": r.reason}

    content = route.Content(form, slots.get("items", []),
                            title=slots.get("title"),
                            subtitle=slots.get("subtitle"), slots=slots)
    try:
        r = route.route_slide(content, fillmap, prs,
                              variant_name=entry.get("variant"))
    except ValueError as e:              # честный отказ роутера (N вне диапазонов)
        raise DeckError(f"{family}: {e}") from e
    if r.action == "GENERATE" and fillmap.get("router") == "timeline":
        base = next(v for v in fillmap["variants"]
                    if v["name"] == fillmap["generate"]["base_variant"])
        return {"mode": "GENERATE", "generator": "timeline", "family": family,
                "entry": entry, "src": anchor.resolve_entry(base, prs),
                "spec": fillmap, "content": route.Content("timeline", slots["stages"]),
                "routed": r.reason}
    if (r.action == "FILL" and fillmap.get("router") == "timeline"
            and "stages" in slots):
        # stages → слоты выбранного варианта (t3: dateN+descN; t4: dateN+descs-repeat);
        # легаси-запись (явный variant + dateN/descN напрямую) идёт мимо маппинга
        stages, vs = slots["stages"], r.fit_report["variant"]["slots"]
        vslots = {key: value for key, value in slots.items() if key in vs}
        for i, st in enumerate(stages, 1):
            if f"date{i}" in vs:
                vslots[f"date{i}"] = st["date"]
            if f"desc{i}" in vs:
                vslots[f"desc{i}"] = st["desc"]
        if "descs" in vs:
            vslots["descs"] = [st["desc"] for st in stages]
        content = route.Content(form, [], title=slots.get("heading"), slots=vslots)
    if r.action != "FILL":
        # плиточной эскалации больше нет (0.21.0, решение владельца 2026-07-23):
        # контент, не заполнивший дизайнерский вариант, — честная ошибка
        raise DeckError(
            f"{family}: {r.reason} — плиточная генерация выпилена: реворди/дели "
            f"контент под дизайнерский вариант (references/families-eda.md) "
            f"или сверстай custom (references/custom-html.md)")
    variant = r.fit_report["variant"]
    if fillmap.get("strict_slots"):
        if (fillmap.get("router") == "timeline" and "stages" in slots
                and len(slots["stages"]) != variant.get("steps")):
            raise DeckError(f"{family}: вариант {variant['name']} вмещает {variant['steps']} этапа; получено {len(slots['stages'])}")
        validate_strict_slots(family, variant, content.slots or {}, entry, prs.slides[r.ready_slide_index])
    # автолегенда: имена серий чарта → слот legend (если не задан явно).
    # Порядок = ранг ПОСЛЕДНЕГО значения (верхняя подпись — серия с наибольшим
    # финальным значением: подписи стоят у концов линий, так их ставил дизайнер)
    charts = entry.get("charts") or {}
    if charts and "legend" in variant["slots"] and not slots.get("legend"):
        series = next(iter(charts.values())).get("series") or {}
        names = sorted(series, key=lambda n: (series[n][-1] is None,
                                              -(series[n][-1] or 0)))
        if names:
            content.slots = dict(slots, legend=names)
    return {"mode": "FILL", "family": family, "entry": entry,
            "src": r.ready_slide_index, "variant": variant,
            "content": content, "reason": r.reason, "routed": r.reason}


def validate_strict_slots(family, variant, slots, entry, slide=None):
    """Opt-in contract for clean local packs: no retained donor placeholders.

    Existing packs retain their existing filling behavior. Reject partial or
    oversized input before cloning, including the fixed capacity of native tables.
    """
    for name, meta in variant["slots"].items():
        if slide is not None:
            wanted = meta.get("match_names") or [meta.get("match_name")]
            actual = [sh.name for sh in slide.shapes]
            if any(not n or actual.count(n) != 1 for n in wanted):
                raise DeckError(f"{family}: слот {name} не найден однозначно в шаблоне; обнови fill-map")
        value = slots.get(name)
        empty = value is None or value == [] or (isinstance(value, str) and not value.strip())
        if empty and not meta.get("optional"):
            raise DeckError(f"{family}: обязателен slots.{name}")
        if empty:
            continue
        if meta.get("repeat"):
            expected = len(meta.get("match_names") or [])
            if not isinstance(value, list) or len(value) != expected:
                raise DeckError(f"{family}: slots.{name} должен содержать ровно {expected} элементов")
            if any(v is None or not str(v).strip() for v in value):
                raise DeckError(f"{family}: slots.{name} содержит пустой элемент")
        if meta.get("kind") == "table":
            if not isinstance(value, list) or not value or any(not isinstance(row, list) or not row for row in value):
                raise DeckError(f"{family}: slots.{name} должен быть непустой прямоугольной таблицей")
            width = len(value[0])
            if any(len(row) != width for row in value):
                raise DeckError(f"{family}: строки таблицы должны иметь одинаковое число колонок")
            if any(cell is None for row in value for cell in row):
                raise DeckError(f"{family}: None в таблице сохраняет заглушку; передай текст или пустую строку")
            if len(value) > meta["max_rows"] or width > meta["max_cols"]:
                raise DeckError(f"{family}: таблица больше вместимости {meta['max_rows']}×{meta['max_cols']}; раздели её на слайды")
            if meta.get('exact_cols') and width != meta['exact_cols']:
                raise DeckError(f"{family}: нужно ровно {meta['exact_cols']} колонки")
    for name in variant.get("chart_slots") or {}:
        data = (entry.get("charts") or {}).get(name) or {}
        if not data.get("categories") or not data.get("series"):
            raise DeckError(f"{family}: обязательны charts.{name}.categories и series")
        if variant['chart_slots'][name].get('combo_series'):
            try:
                CHRT.validate_combo_data(data, variant['chart_slots'][name]['combo_series'])
            except (CHRT.ChartError, TypeError, ValueError) as e:
                raise DeckError(str(e)) from e


def build_slide_on_clone(prs, plan: dict, tdir: str) -> tuple[int, dict]:
    """Фаза 2: безопасный клон источника → заливка КЛОНА. Оригинал остаётся
    пристинным (его вместе с палитрой уберёт keep_ordered).
    CUSTOM — без донора: новый слайд с лэйаута хрома (lib/html2pptx.py)."""
    family, entry = plan["family"], plan["entry"]

    if plan["mode"] == "CUSTOM":
        try:
            ci, rep = H2P.build_custom_slide(prs, plan["measure"], tdir)
        except H2P.Html2PptxError as e:
            raise DeckError(f"{family}: {e}") from e
        for wmsg in rep["warnings"]:
            print(f"  ⚠ {family}: {wmsg}")
        return ci, rep

    ci = CL.clone_slide(prs, plan["src"])
    slide = prs.slides[ci]

    if plan.get("generator") == "statgrid":
        spec, content = plan["spec"], plan["content"]
        overflow = []
        if plan["mode"] == "GENERATE":
            try:
                overflow = GEN.generate_statgrid(prs, spec, content, ci)
            except GEN.GenerateError as e:
                raise DeckError(f"{family}: {e}") from e
        else:
            cards = GEN.locate_cards(slide, spec)
            cu, vfit = spec["card_unit"], spec["fit"]["value"]
            # потолок = внутренняя зона плашки (вылет за плашку — дефект B10;
            # люфт «бокс + 1"» отменён user-фидбеком 2026-07-07)
            max_val_w = F.value_max_width(cu["value"], vfit)
            vals = [str(it["value"]) for it in content.items]
            vshapes = [c.value for c in cards[:len(vals)] if c.value is not None]
            # однородные значения — ЕДИНЫЙ кегль (user 2026-07-02)
            group_pt = F.value_group_pt(vshapes, vals,
                                        {"box_w_in": max_val_w,
                                         "min_pt": vfit["min_pt"],
                                         "font": cu["value"].get("font"),
                                         "size_pt": cu["value"].get("size_pt")})
            for i, (card, item) in enumerate(zip(cards, content.items), 1):
                if card.value is not None:
                    F.set_value_slot(card.value, str(item["value"]),
                                     max_width_in=max_val_w,
                                     min_pt=vfit["min_pt"], force_pt=group_pt)
                    if F.value_overflows(card.value, item["value"], max_val_w):
                        overflow.append(f"value_{i}")
                if card.label is not None:
                    F.fill_lines(card.label, [str(item["label"])])
                    F.strip_empty_lines(card.label.text_frame)
        for pref, val in (("Заголовок", plan["content"].title),
                          ("Подзаголовок", plan["content"].subtitle)):
            sh = F.find_shape_by_text(slide, pref)
            if sh is not None and val:
                F.clear_soft_breaks(sh.text_frame)
                F.fill_lines(sh, [val])
                F.strip_empty_lines(sh.text_frame)
            elif sh is not None and sh.text_frame.text.strip():
                print(f"  ⚠ {family}: слот «{pref.lower()}» не заполнен — на слайде "
                      f"останется заглушка дизайнера "
                      f"{sh.text_frame.text.strip()[:30]!r}; дай текст или убери слот")
        return ci, {"action": plan["mode"],
                    "slots": {"items": len(plan["content"].items)},
                    "overflow_slots": overflow, "src": plan["src"]}

    if plan["mode"] == "GENERATE" and plan.get("generator") == "timeline":
        gspec = plan["spec"]["generate"]
        try:
            GEN.generate_timeline(prs, gspec, plan["content"], ci, tdir=tdir)
        except GEN.GenerateError as e:
            raise DeckError(f"{family}: {e}") from e
        fill_heading(slide, (entry.get("slots") or {}).get("heading"))
        problems = GEN.verify_timeline(slide, gspec, len(plan["content"].items))
        if problems:
            raise DeckError(f"{family}: composition-check таймлайна FAIL: {problems}")
        return ci, {"action": "GENERATE", "rows": [len(plan["content"].items)],
                    "size_pt": gspec["fit"]["desc"]["size_pt"], "flags": [],
                    "src": plan["src"]}

    if plan["mode"] == "GENERATE":
        # недостижимо после 0.21.0 (плиточная генерация выпилена; клон-ветки
        # timeline/statgrid обработаны выше) — защитный стоп от рассинхрона
        raise DeckError(f"{family}: план GENERATE без генератора — плиточная "
                        f"генерация выпилена (0.21.0), это рассинхрон роутера; "
                        f"сообщи владельцу скилла")

    # защита от опечаток deckspec: неизвестный ключ слота = громкая ошибка,
    # а не молчаливая заглушка дизайнера в финальном деке (урок пробы 2026-07-02)
    given = set(entry.get("slots") or {})
    allowed = set(plan["variant"]["slots"]) | {"active_index", "stages", "bullets"}
    for _slot in plan["variant"]["slots"].values():      # композитные слоты (v2.2 Хурала):
        allowed |= set(_slot.get("compose") or ())       # deckspec-ключи живут в compose
    if plan["variant"].get("decor_highlights"):
        allowed |= {"highlights"}        # ручка «без подсветок» (datatable, v0.15.0)
    if plan["variant"].get("accent_cards"):
        allowed |= {"accent"}            # ручка «все карточки серые» (v0.16.0)
    if family == "timeline":
        allowed |= {"accent"}            # O9: номер акцентной вехи (1-based)
    unknown = given - allowed
    if unknown:
        raise DeckError(f"{family}: слотов {sorted(unknown)} нет у варианта "
                        f"'{plan['variant'].get('name')}' (есть: "
                        f"{sorted(plan['variant']['slots'])})")

    rep = F.fill_structural_slide(slide, plan["variant"], plan["content"])
    if tdir == "committee":
        import committee_layout
        rep["_composition"] = committee_layout.apply(slide, family, plan["variant"])

    # O9 (хэндофф №2 2026-08-13): акцент вехи таймлайна — семантика контента.
    # Применяется ТОЛЬКО к вариантам с пилюлями atom:milestone_date (линейка
    # Хурала): там дефолт «без ключа — вся линейка серая» гасит донорскую
    # зелёную третью веху. Варианты других шаблонов без пилюль (кривая Еды)
    # не трогаем; заданный на них accent — честная ошибка.
    if family == "timeline":
        acc = (entry.get("slots") or {}).get("accent")
        try:
            F.apply_timeline_accent(slide, acc)
        except ValueError as e:
            if acc is not None:
                raise DeckError(f"{family}: {e}") from e
        else:
            if acc is not None:
                print(f"  {family}: акцентная веха №{acc} — зелёная пилюля, "
                      f"остальные серые (O9)")

    trep = next((v for v in rep.values()
                 if isinstance(v, dict) and v.get("donor_dims")), None)
    if trep and (trep.get("trimmed_rows") or trep.get("trimmed_cols")):
        print(f"  {family}: таблица ужата до {trep['dims'][0]}×{trep['dims'][1]} "
              f"(донор {trep['donor_dims'][0]}×{trep['donor_dims'][1]}) — лишние "
              f"строки/колонки удалены физически, высоты донорские")
    if trep and trep.get("stretched"):
        # высотный stretch логируем явно: он единственный, кто трогает КЕГЛЬ —
        # без этой строки рост кегля приписывали stretch_width (хэндофф Ultima O3)
        font_note = ("кегль донорский (stretch_font: false)"
                     if trep.get("stretch_font") is False
                     else "кегль растёт с потолком ×4/3")
        print(f"  {family}: строки растянуты ×{trep['stretched']} до донорской "
              f"высоты таблицы (stretch, общий режим); {font_note}")
    if trep and trep.get("stretched_width"):
        print(f"  {family}: колонки растянуты ×{trep['stretched_width']} до "
              f"донорской ширины таблицы (stretch_width, общий режим; кегль "
              f"от ширины не растёт)")
    if trep and trep.get("aligned"):
        al = trep["aligned"]
        if al.get("r"):
            print(f"  {family}: числовые колонки {al['r']} выровнены вправо, "
                  f"текстовые {al.get('l') or '—'} влево (правило вёрстки)")
    if rep.get("_decor_removed"):
        print(f"  {family}: декоративные подсветки дизайнера сняты "
              f"({len(rep['_decor_removed'])}) — прибиты к ячейкам полной сетки")
    eslots = entry.get("slots") or {}
    if eslots.get("bullets") is False:
        F.strip_bullets_slide(slide)
        print(f"  {family}: текстовые буллеты сняты (slots.bullets=false)")
    if eslots.get("accent") is False and plan["variant"].get("accent_cards"):
        gone = F.decolor_accent_cards(slide, plan["variant"]["accent_cards"])
        if gone:
            print(f"  {family}: жёлтый акцент карточки снят ({', '.join(gone)}) — "
                  f"карточки равноправные (slots.accent=false)")
    for sname, srep in rep.items():
        if isinstance(srep, dict) and srep.get("text_shrink"):
            print(f"  {family}.{sname}: ужат единым кеглем до "
                  f"{srep['text_shrink']:.0f}pt (fit_policy text_shrink, "
                  f"floor min_pt спеки)")
    if rep.get("_geometry_broken"):
        raise DeckError(f"{family}: {rep['_geometry_broken']}")
    if rep.get("_slot_conflicts"):
        raise DeckError(f"{family}: {'; '.join(rep['_slot_conflicts'])}")
    filled_any = any(isinstance(v, dict) and (v.get("filled") or v.get("cleared")
                                              or v.get("cells_filled"))
                     for v in rep.values())
    if not filled_any and (plan["variant"]["slots"] or not entry.get("photos")):
        # фото-only вариант (текст-слотов в спеке нет, фото переданы) легитимен:
        # заполнение проверит photo-swap ниже; иначе — прежний гейт от опечаток
        raise DeckError(f"{family}: ни один слот не заполнился — проверь ключи "
                        f"slots (доступны: {sorted(plan['variant']['slots'])})")
    # утечка заглушки дизайнера: объявленный текст-слот варианта без значения в
    # deckspec и без авточистки — демо-текст шаблона уедет в финальный дек
    for sname, s in plan["variant"]["slots"].items():
        if sname in given or s.get("repeat") or s.get("kind") == "table":
            continue
        v = rep.get(sname)
        if isinstance(v, dict) and (v.get("filled") or v.get("cleared")
                                    or v.get("cells_filled")):
            continue
        sh = F.find_shape_by_name(slide, s.get("match_name") or "")
        txt = (sh.text_frame.text.strip()
               if sh is not None and sh.has_text_frame else "")
        if txt:
            print(f"  ⚠ {family}: слот «{sname}» не заполнен — в деке останется "
                  f"заглушка дизайнера {txt[:30]!r}; дай текст (или слот должен "
                  f"чиститься сам — скажи владельцу скилла)")
    photos = entry.get("photos")
    if photos:
        local = {k: resolve_photo(v, tdir) for k, v in photos.items()}
        prep = IMG.swap_image_slots(slide, plan["variant"], local)
        bad = {k: v for k, v in prep.items()
               if v.get("error") or (v.get("swapped") and not v.get("aspect_ok"))}
        if bad:
            raise DeckError(f"{family}: photo-swap FAIL: {bad}")
        unknown = set(photos) - set(plan["variant"].get("image_slots") or {})
        if unknown:
            raise DeckError(f"{family}: фото-слотов {sorted(unknown)} нет у варианта "
                            f"(есть: {sorted(plan['variant'].get('image_slots') or {})})")
        for k, v in prep.items():
            if v.get("low_res"):
                print(f"  ⚠ {family}.{k}: «{v['image']}» {v['eff_dpi']} dpi на этом "
                      f"боксе — будет мыло; возьмите фото крупнее/горизонтальное")
        rep["_photos"] = prep
    charts = entry.get("charts")
    if charts:
        cslots = plan["variant"].get("chart_slots") or {}
        unknown = set(charts) - set(cslots)
        if unknown:
            raise DeckError(f"{family}: chart-слотов {sorted(unknown)} нет у "
                            f"варианта (есть: {sorted(cslots)})")
        crep = CHRT.replace_chart_slots(slide, plan["variant"], charts)
        bad = {k: v for k, v in crep.items()
               if v.get("error") or (v.get("replaced") and not v.get("verify_ok"))}
        if bad:
            raise DeckError(f"{family}: chart replace_data FAIL: {bad}")
        for k, v in crep.items():
            if v.get("style_cloned"):
                print(f"  ⚠ {family}.{k}: {v['style_cloned']}")
        rep["_charts"] = crep
    styles = {k: v["style_problems"] for k, v in rep.items()
              if isinstance(v, dict) and v.get("style_problems")}
    if styles:                        # перестановка стилей агенды сломалась
        raise DeckError(f"{family}: стилевой чекер FAIL: {styles}")
    overflows = {k: v for k, v in rep.items()
                 if isinstance(v, dict) and (v.get("overflow") or v.get("too_many"))}
    return ci, {"action": "FILL", "variant": plan["reason"], "slots": rep,
                "overflow_slots": list(overflows), "src": plan["src"]}


def check_template_drift(tdir: str) -> None:
    """Дизайнер активно правит шаблон: сверяем живой шаблон с бейзлайн-инвентарём
    (layouts/<t>/_inventory.json) и громко предупреждаем о дрифте. Ранний сигнал:
    при реальной поломке якоря упадут сами, но уже непонятной ошибкой резолва."""
    base = os.path.join(LAYOUT_DIRS[tdir], "_inventory.json")
    if not os.path.exists(base):
        return
    try:
        import inventory_template as INV
        inv = INV.build_inventory(template_path(tdir))
        with open(base, encoding="utf-8") as f:
            old = json.load(f)
    except Exception as e:                    # дрифт-чек не должен валить сборку
        print(f"  ⚠ drift-чек не отработал: {e}")
        return

    def key_of(sl):
        uq = sl.get("unique_strings")
        # min по (len, str): порядок в unique_strings нестабилен между процессами
        # (сортировка только по len + set-итерация) — ключ берём детерминированно
        return min(uq, key=lambda s: (len(s), s)) if uq else sl["all_text"][:40]

    old_pos = {key_of(sl): sl["index1"] for sl in old["slides"]}
    new_pos = {key_of(sl): sl["index1"] for sl in inv["slides"]}
    moved = [(k, old_pos[k], new_pos[k])
             for k in new_pos if k in old_pos and old_pos[k] != new_pos[k]]
    added = [k for k in new_pos if k not in old_pos]
    removed = [k for k in old_pos if k not in new_pos]
    if not (old.get("slide_count") != inv["slide_count"] or moved or added or removed):
        return
    print(f"  ⚠ ШАБЛОН ИЗМЕНИЛСЯ с бейзлайна: {old.get('slide_count')} → "
          f"{inv['slide_count']} слайдов, {len(moved)} перемещений, "
          f"{len(added)} новых / {len(removed)} убранных строк-сигнатур")
    for k, o, n in sorted(moved, key=lambda x: x[2])[:5]:
        print(f"      MOVED {o}→{n} «{k[:45]}»")
    for k in added[:5]:
        print(f"      ADDED →{new_pos[k]} «{k[:45]}»")
    for k in removed[:5]:
        print(f"      REMOVED (was {old_pos[k]}) «{k[:45]}»")
    print("      Обычно это правки контента дизайнером — сборка продолжится; если "
          "упадёт резолв якорей, шаблон разошёлся со спеками: сообщите владельцу "
          "скилла (диагностика: scripts/inventory_template.py --diff)")


def gate(out_path: str) -> dict:
    """crash + lint (composition уже прошёл при сборке). Crash-чекер вендорен в
    scripts/ (папка скилла самодостаточна, работает и в облаке); источник правды
    новых паттернов порчи — скилл pptx-crash-check, при обновлении копировать."""
    checker = os.path.join(REPO, "scripts", "check_pptx.py")
    if not os.path.exists(checker):
        checker = os.path.expanduser(
            "~/.claude/skills/pptx-crash-check/scripts/check_pptx.py")
    crash = subprocess.run(
        [sys.executable, checker, out_path],
        capture_output=True, text=True)
    lint = subprocess.run([sys.executable, os.path.join(REPO, "lib/lint.py"),
                           out_path], capture_output=True, text=True)
    return {"crash_ok": crash.returncode == 0,
            "lint_ok": lint.returncode == 0,
            "lint_tail": lint_high_lines(lint.stdout)}


def lint_high_lines(lint_stdout: str) -> list[str]:
    """ВСЕ [HIGH]-строки отчёта линта (фидбек 2026-08-10: гейт показывал 3
    хвостовые строки — HIGH уезжали из виду, находки чинились по одной за
    пересборку). Ни одной HIGH-строки → 3 хвостовые (контекст сбоя линта)."""
    lines = lint_stdout.strip().splitlines()
    highs = [l.strip() for l in lines if l.strip().startswith("[HIGH]")]
    return highs or lines[-3:]


def custom_text_columns_reason(measure: dict) -> str | None:
    """A13 (fill-first): тело custom-слайда = 2–4 однородные текстовые карточки
    в один ряд и ничего кроме них → это готовое семейство columns/takeaways,
    а не custom. Возвращает причину или None (слайд легитимно кастомный).
    Легитимный кастом мимо эвристики проходит свободно: любая картинка, SVG,
    таблица, крупный не-текстовый атом или текст вне карточек — не флаг."""
    if measure.get("images") or measure.get("svgs") or measure.get("tables") \
            or measure.get("charts"):
        return None
    boxes = measure.get("boxes") or []
    cards = [b for b in boxes
             if b.get("bg") and b["bbox"]["w"] >= 260 and b["bbox"]["h"] >= 220]
    if not 2 <= len(cards) <= 4:
        return None
    ys = [c["bbox"]["y"] for c in cards]
    ws = [c["bbox"]["w"] for c in cards]
    if max(ys) - min(ys) > 8 or max(ws) - min(ws) > 8:
        return None                     # не один ряд одинаковых карточек
    decor = [b for b in boxes if b not in cards]
    if any(max(b["bbox"]["w"], b["bbox"]["h"]) > 40 for b in decor):
        return None                     # крупный не-текстовый атом — композиция сложнее
    texts = [t for t in measure.get("texts") or []
             if not t["cls"].startswith(("h1", "pagenum", "sub"))
             and not H2P.is_title_text(t)]

    def in_card(t):
        cx = t["bbox"]["x"] + t["bbox"]["w"] / 2
        cy = t["bbox"]["y"] + t["bbox"]["h"] / 2
        return any(c["bbox"]["x"] <= cx <= c["bbox"]["x"] + c["bbox"]["w"]
                   and c["bbox"]["y"] <= cy <= c["bbox"]["y"] + c["bbox"]["h"]
                   for c in cards)
    if not texts or not all(in_card(t) for t in texts):
        return None                     # текст вне карточек — композиция сложнее
    return (f"тело слайда — {len(cards)} однородные текстовые карточки "
            f"в один ряд и ничего кроме них")


def _custom_body_atoms(measure: dict, tdir: str) -> list[tuple]:
    """Атомы ТЕЛА custom-слайда по обмеру: (вид, класс, bbox, санкция-фото).
    h1/pagenum — не тело (их даёт хром лэйаута). Санкция — фото-панель
    дизайнерского канона (от верхнего инсета до нижней кромки). Общий сборщик
    гвардов вертикали: A15 (за полями) и A18 (всплыл над нижним полем)."""
    H = measure["slide"]["h"]
    pp_top_max, pp_bot_min = TP.a15_photo_panel(tdir)   # санкция фото-панели — из профиля
    atoms = []
    for b in measure.get("boxes") or []:
        atoms.append(("плашка", b.get("cls", ""), b["bbox"], False))
    for t in measure.get("texts") or []:
        if t["cls"].startswith(("h1", "pagenum")) or H2P.is_title_text(t):
            continue                 # хром лэйаута (title/pagenum) — не тело
        atoms.append(("текст", t.get("cls", ""), t["bbox"], False))
    for tb in measure.get("tables") or []:
        atoms.append(("таблица", tb.get("cls", ""), tb["bbox"], False))
    for ln in measure.get("lines") or []:
        atoms.append(("линия", ln.get("cls", ""), ln["bbox"], False))
    for sv in measure.get("svgs") or []:
        atoms.append(("svg", sv.get("cls", ""), sv["bbox"], False))
    for chn in measure.get("charts") or []:
        atoms.append(("чарт", chn.get("cls", ""), chn["bbox"], False))
    for im in measure.get("images") or []:
        fr = im["frame"]
        sanctioned = fr["y"] <= pp_top_max and fr["y"] + fr["h"] >= H - pp_bot_min
        atoms.append(("фото", im.get("cls", ""), fr, sanctioned))
    return atoms


def custom_anchor_violations(measure: dict, anchors: list[dict],
                             tdir: str = "eda") -> list[str]:
    """A15 (фидбек-репорт 2026-07-20 «custom заезжает на номер страницы»;
    user-правило: якоря лэйаута — заголовок/футер/номер — неподвижны, вёрстка
    строится внутри них, наложение лечится размерами карточек/текста, не
    перекрытием). Возвращает нарушения: атом тела ЗА КАНВОЙ либо В ПОЛОСЕ
    якоря низа (anchors — rect'ы в px канвы: футер-плейсхолдер idx=10 и
    статическое поле номера fld slidenum выбранного лэйаута; пустой список —
    в лэйауте полосы нет). h1/pagenum — не тело, их даёт хром лэйаута.
    Единственное исключение полосы — фото-панель дизайнерского канона
    (inset-29 до нижней кромки): у дизайнеров номер лежит на фото
    (custom-html.md «не дефект»); контентным атомам исключений нет."""
    W, H = measure["slide"]["w"], measure["slide"]["h"]
    eps = 2
    atoms = _custom_body_atoms(measure, tdir)
    m = TP.margins(tdir)                 # контентные поля канвы (user-фидбек
    out = []                             # 2026-08-07: «распределять по полям»)
    for kind, cls, bb, sanctioned in atoms:
        x2, y2 = bb["x"] + bb["w"], bb["y"] + bb["h"]
        head = cls.split()[0] if cls else ""
        tag = f"{kind}{' .' + head if head else ''} (y {bb['y']:.0f}→{y2:.0f})"
        if bb["x"] < -eps or bb["y"] < -eps or x2 > W + eps or y2 > H + eps:
            out.append(f"{tag} за канвой {W:.0f}×{H:.0f}")
            continue
        if m and not sanctioned:         # фото-панель канона до кромки — санкция та же
            edges = (("левое", bb["x"] < m["left"] - eps),
                     ("правое", x2 > W - m["right"] + eps),
                     ("верхнее", bb["y"] < m["top"] - eps),
                     ("нижнее", y2 > H - m["bottom"] + eps))
            hit = [n for n, bad in edges if bad]
            if hit:
                out.append(f"{tag} нарушает {'/'.join(hit)} поле канвы "
                           f"(поля {m['left']}/{m['right']}/{m['top']}/{m['bottom']}px)")
                continue
        for a in anchors:
            if not sanctioned \
                    and x2 > a["x"] + eps and bb["x"] < a["x"] + a["w"] - eps \
                    and y2 > a["y"] + eps and bb["y"] < a["y"] + a["h"] - eps:
                out.append(f"{tag} в полосе номера страницы "
                           f"y {a['y']:.0f}→{a['y'] + a['h']:.0f}")
                break
    return out


# A18: порог «всплывшего» кастома — зазор между низом чернил и нижним полем
# как доля КОНТЕНТНОЙ высоты канвы (H − верхнее − нижнее поле). 8% — заказ
# хэндоффа; на Хурале (920px контента) это 73.6px, что флагует ровно три
# слайда дека Ultima, которые дизайнер назвала «поднятыми» (низ 74.5–81.8%
# высоты), и пропускает матрицу (87%), к которой претензий по низу не было.
_A18_BOTTOM_GAP_FRAC = 0.08


def custom_bottom_gap_violation(measure: dict, tdir: str = "eda") -> str | None:
    """A18 (хэндофф Ultima 2026-08-11 O1, слова дизайнера: «плашки подняты
    сильно от нижней границы» — три кастом-слайда из четырёх). Дизайнерские
    доноры доводят чернила до нижнего поля и даже ниже (94–100% высоты), а
    кастом, кончающийся заметно выше поля, визуально «всплывает» — и до сих
    пор это не ловилось: A15 односторонний (только выход ЗА поле). Меряем
    зазор между нижней кромкой чернил тела и нижним полем канвы; больше 8%
    контентной высоты — нарушение. Вместе A15+A18 держат низ контента в
    коридоре у нижнего поля. Пустой слайд не судим; санкционированная
    фото-панель канона доходит до кромки и закрывает вопрос сама."""
    m = TP.margins(tdir)
    if not m:
        return None
    H = measure["slide"]["h"]
    atoms = _custom_body_atoms(measure, tdir)
    if not atoms:
        return None
    ink_bottom = max(bb["y"] + bb["h"] for _, _, bb, _ in atoms)
    floor = H - m["bottom"]                      # нижнее поле (у Хурала y=960)
    content_h = H - m["top"] - m["bottom"]
    gap = floor - ink_bottom
    if gap <= _A18_BOTTOM_GAP_FRAC * content_h:
        return None
    return (f"контент кончается на y {ink_bottom:.0f} — на {gap:.0f}px выше "
            f"нижнего поля (y {floor:.0f}): слайд «всплывает», дизайнерские "
            f"доноры доводят чернила до поля. Прижми блок контента к нижнему "
            f"полю (низ ≈ y {floor:.0f}, не выше y {floor - 60:.0f}), а "
            f"свободный воздух отдай МЕЖДУ заголовком и блоком, не под блок "
            f"(правило вертикали — custom-html.md «Вёрстка в полях и плашках»)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deckspec")
    ap.add_argument("-o", "--out", default="deck.pptx")
    ap.add_argument("--dry-run", action="store_true",
                    help="только фаза резолва: валидация deckspec + план "
                         "роутера, файл не собирается")
    a = ap.parse_args()
    with open(a.deckspec, encoding="utf-8") as f:
        ds = json.load(f)
    # html кастом-слайдов лежит РЯДОМ с deckspec — резолвим от него, не от cwd
    spec_dir = os.path.dirname(os.path.abspath(a.deckspec))
    for e in ds["slides"]:
        if e.get("family") == "custom" and e.get("html") \
                and not os.path.isabs(e["html"]):
            e["html"] = os.path.join(spec_dir, e["html"])
    tdir = ds.get("template", "еда").lower()
    tdir = TDIR_ALIASES.get(tdir, tdir)
    if tdir not in LAYOUT_DIRS:
        raise SystemExit(f"неизвестный шаблон {tdir!r}; доступны: {list(LAYOUT_DIRS)}")
    TP.set_active(tdir)              # ambient tdir дека: fill-хот-путь (bold-карта) читает его
    tpl_path = template_path(tdir)   # дев-файл → кэш → Я.Диск (до дрифт-чека)
    check_template_drift(tdir)

    prs = Presentation(tpl_path)
    F.set_theme_fonts(*theme_latin_fonts(prs))   # наследуемые/темовые шрифты ранов → фактические семейства (bold-карта)

    # фаза 1: резолв всех записей по пристинному шаблону (якоря не видят клонов)
    # notes снимаем ДО typo_walk: заметкам спикера NBSP и срез точки не нужны
    raw_notes = [e.get("notes") for e in ds["slides"]]
    plans = [resolve_entry_plan(prs, typo_walk(e), tdir) for e in ds["slides"]]

    # chart-guard (known_issues #6): chart/OLE-парт нельзя делить между клонами
    # (custom-слайды без донора: src is None)
    for src, cnt in Counter(p["src"] for p in plans if p["src"] is not None).items():
        if cnt > 1 and CL.slide_needs_exclusive_parts(prs.slides[src]):
            fam = next(p["family"] for p in plans if p["src"] == src)
            raise DeckError(
                f"дизайнерский слайд #{src + 1} ({fam}) содержит chart/OLE и "
                f"запрошен {cnt} раза — Mac PP не прощает шаринг chart-партов "
                f"(pptx-crash-check #6: даже приватные копии партов не спасли, "
                f"нужен верифицированный в PowerPoint deep-clone — спайк "
                f"chart-cloning 2026-07, архив владельца). Нативные чарт-доноры "
                f"пока НЕПОВТОРЯЕМЫ: оставьте одно вхождение, а второй график "
                f"покажите таблицей (datatable) или другим семейством — chart-"
                f"selection.md §2. custom-шейпы для второго графика = НЕ "
                f"редактируемый чарт, так не делаем")

    # A14 (v0.16.0, фидбек C медиакита «слайды 10 и 12 с одинаковым фото»):
    # повтор донора с фото без photo-swap = одна и та же картинка дважды
    photo_use = Counter()
    for i, p in enumerate(plans):
        if p["src"] is not None and p["mode"] == "FILL" \
                and (p.get("variant") or {}).get("photo") \
                and not ds["slides"][i].get("photos"):
            photo_use[p["src"]] += 1
    for src, cnt in photo_use.items():
        if cnt > 1:
            print(f"  ⚠ A14: донор #{src + 1} с фото повторён {cnt} раза без замены "
                  f"фото — в деке будет одна и та же картинка; задай \"photos\" "
                  f"хотя бы на повторах (слоты — image_slots в families-{tdir}.md)")

    # impact-честность (фидбек раунд 2): мини-чарты think-cell #30 — ДЕКОР с
    # цифрами дизайнера, движок их не заполняет (add-in-метаданные хрупкие).
    # На деке пользователя это ЧУЖИЕ числа — предупреждаем явно.
    for i, p in enumerate(plans, 1):
        if p["family"] == "impact":
            print(f"  ⚠ impact (слайд #{i}): мини-чарты think-cell — ДУММИ дизайнера, "
                  f"цифры НЕ ваши и не заполняются (хрупкий add-in). Замените слайд "
                  f"или согласуйте, что диаграмма — декор без ваших данных")

    # A10: ≥3 слайдов одного семейства подряд — монотонность (флаг, не запрет)
    fams = [p["family"] for p in plans]
    for i in range(2, len(fams)):
        if fams[i] == fams[i - 1] == fams[i - 2]:
            print(f"  ⚠ A10: «{fams[i]}» ≥3 подряд (слайды {i - 1}–{i + 1}) — "
                  f"монотонно; сначала исчерпайте варианты семейства")
            break

    # A11: разнообразие семейств (user-фидбек 2026-07-07 «скилл любит одну
    # структуру»; усилен 0.21.0 — жалоба тестеров: повторяемость МЕЖДУ деками,
    # постоянный пользователь получает одни и те же слайды) — флаг, не запрет
    if len(plans) >= 6 and len(set(fams)) <= 4:
        print(f"  ⚠ A11: {len(plans)} слайдов из {len(set(fams))} семейств — "
              f"однообразно; шаблон богаче (правило разнообразия: SKILL.md "
              f"шаг 2, карта «контент → семейства» в deck-archetypes.md)")
    _structural = {"cover", "toc", "divider", "closing", "contacts"}
    _frequent = {"two_values", "columns", "timeline"}  # содержательная часть частой пятёрки
    content_fams = [f for f in fams if f not in _structural]
    if len(content_fams) >= 4:
        freq = sum(1 for f in content_fams if f in _frequent)
        if freq * 2 > len(content_fams):
            print(f"  ⚠ A11: {freq} из {len(content_fams)} содержательных слайдов — "
                  f"частая пятёрка (two_values/columns/timeline); квота — не больше "
                  f"половины (SKILL.md шаг 2): покажи часть контента другими "
                  f"семействами (карта в deck-archetypes.md §0)")

    # A12 (user-правило 2026-07-13 «однотипные текстовые плашки заказчиков
    # не устраивают»; плиточная часть снята 0.21.0 — плиток в движке больше
    # нет): повтор текстово-карточного FILL-семейства — флаг
    for fam, cnt in Counter(p["family"] for p in plans
                            if p["family"] in ("columns", "takeaways")
                            and p["mode"] == "FILL"
                            # фото-варианты (three_photos и т.п.) — не текстовые плашки
                            and not (p.get("variant") or {}).get("image_slots")).items():
        if cnt > 1:
            print(f"  ⚠ A12: «{fam}» {cnt} раза — текстово-карточная структура "
                  f"повторяется; однотипные плашки с текстом делают дек "
                  f"монотонным, покажи списки разными семействами")

    # Гварды custom (A13/A15/A16/A17) — ВСЕ нарушения по ВСЕМ слайдам за ОДИН
    # проход (фидбек переноса HTML 2026-08-10: «каждая пересборка находила по
    # несколько новых пересечений» — раньше валились на первом слайде)
    guard_msgs: list[tuple[str, str]] = []       # (код, сообщение)
    for i, p in enumerate(plans, 1):
        if p["mode"] != "CUSTOM":
            continue

        # A13 (fill-first, user-правило 2026-07-15 «простой трёхколоночный
        # слайд уехал в custom»): custom из однородных текстовых карточек —
        # это готовое семейство шаблона, HTML-хэтч для него запрещён
        why = custom_text_columns_reason(p["measure"])
        if why:
            guard_msgs.append(("A13",
                f"слайд #{i} [{p['routed']}]: {why} — это дизайнерское семейство "
                f"колонок, custom для текстовых плашек запрещён (fill-first). "
                f"Перепиши запись deckspec: {{\"family\": \"columns\", \"slots\": "
                f"{{\"heading\": ..., \"blocks\": [{{\"name\": ..., \"desc\": ...}}, "
                f"...]}}}} — desc принимает список абзацев; альтернативы "
                f"(takeaways, плитки) — references/families-{tdir}.md"))

        # A15 (фидбек-репорт 2026-07-20: карточки custom накрыли номер
        # страницы, гейт молчал): тело не выходит за канву/поля и не заходит
        # в полосу футера выбранного лэйаута. Якоря неподвижны — лечится
        # размерами карточек/текста (низ 120px, custom-html.md)
        layout = H2P.pick_layout(prs, p["measure"]["slide"]["bg"], tdir)
        anchors = []                     # якоря низа: футер + поле номера
        for sh in layout.shapes:
            is_foot = sh.is_placeholder \
                and sh.placeholder_format.idx == 10
            is_num = not sh.is_placeholder \
                and 'type="slidenum"' in sh._element.xml
            if (is_foot or is_num) and sh.left is not None:
                anchors.append({"x": sh.left / H2P.PX, "y": sh.top / H2P.PX,
                                "w": sh.width / H2P.PX, "h": sh.height / H2P.PX})
        viols = custom_anchor_violations(p["measure"], anchors, tdir)
        if viols:
            guard_msgs.append(("A15",
                f"слайд #{i} [{p['routed']}]: тело заезжает на якоря лэйаута — "
                + "; ".join(viols[:4])
                + (f" (+ещё {len(viols) - 4})" if len(viols) > 4 else "")
                + ". Якоря (заголовок/футер/номер) неподвижны, наложение не "
                  "маскируется — ужми карточки/текст, чтобы контент кончался "
                  "выше полосы номера (контракт канвы: низ 120px, custom-html.md)"))

        # A18 (хэндофф Ultima 2026-08-11, O1): зеркало A15 — контент не
        # ВЫХОДИТ за нижнее поле, но и не ВСПЛЫВАЕТ над ним
        gap_v = custom_bottom_gap_violation(p["measure"], tdir)
        if gap_v:
            guard_msgs.append(("A18", f"слайд #{i} [{p['routed']}]: {gap_v}"))

        # A16/A17 (фидбек переноса HTML 2026-08-10): заголовок мимо TITLE-
        # плейсхолдера / график картинкой (SVG с числовыми подписями)
        for v in H2P.custom_content_violations(p["measure"], tdir):
            guard_msgs.append((v.split(":")[0],
                               f"слайд #{i} [{p['routed']}]: {v}"))

    if guard_msgs:
        if ds.get("regression_smoke"):
            # смоукам кастом-ветки гварды ослабляются до флага (TESTING.md)
            for code, msg in guard_msgs:
                print(f"  ⚠ {code} (ослаблен: regression_smoke): {msg}")
        else:
            raise DeckError(
                f"гварды custom: {len(guard_msgs)} нарушений — полный список "
                f"одним проходом, чини все разом и пересобирай:\n  - "
                + "\n  - ".join(m for _, m in guard_msgs))

    if a.dry_run:
        for i, plan in enumerate(plans, 1):
            why = f"  [{plan['routed']}]" if plan.get("routed") else ""
            donor = (f"донор — слайд #{plan['src'] + 1}" if plan["src"] is not None
                     else "хром лэйаута, без донора")
            print(f"  {i}. {plan['family']:12} → {plan['mode']} ({donor}){why}")
        print(f"\ndry-run: спека валидна, план на {len(plans)} слайдов; "
              f"сборка не выполнялась (уберите --dry-run)")
        return

    # фаза 2: клон → заливка клона
    order = []
    for i, plan in enumerate(plans, 1):
        idx, rep = build_slide_on_clone(prs, plan, tdir)
        order.append(idx)
        if raw_notes[i - 1]:
            set_speaker_notes(prs, idx, raw_notes[i - 1])
        extra = ""
        if rep.get("overflow_slots"):
            extra = f"  ⚠ overflow: {rep['overflow_slots']} — REWORD контент"
        why = f"  [{plan['routed']}]" if plan.get("routed") else ""
        marks = "  +заметки" if raw_notes[i - 1] else ""
        donor = (f"клон слайда #{rep['src'] + 1}" if rep.get("src") is not None
                 else f"хром лэйаута {rep.get('layout', '')}".rstrip())
        print(f"  {i}. {plan['family']:12} → {rep['action']} "
              f"({donor}){why}{marks}{extra}")

    keep_ordered(prs, order)
    widows = LINT.defuse_widows(prs)          # висячие слова заголовков → nbsp (B13)
    if widows:
        print(f"  ✓ widow: склеено {widows} висячих слов заголовков (nbsp)")
    prs.save(a.out)
    g = gate(a.out)
    print(f"\nсохранено: {a.out} ({len(order)} слайдов)")
    print(f"гейт: crash {'✓' if g['crash_ok'] else '✗ FAIL'} | "
          f"lint {'✓' if g['lint_ok'] else '✗ HIGH (все находки):'}")
    if not g["lint_ok"]:
        for l in g["lint_tail"]:
            print(f"    {l}")
    if not (g["crash_ok"] and g["lint_ok"]):
        # красный гейт → артефакт не должен уйти пользователю под рабочим именем
        # (user-фидбек 2026-08-07: дек с красным гейтом отдали дизайнеру)
        rej = (a.out[:-5] if a.out.endswith(".pptx") else a.out) + ".REJECTED.pptx"
        os.replace(a.out, rej)
        print(f"гейт КРАСНЫЙ — дек НЕ готов и НЕ показывается (SKILL.md §5); "
              f"файл перемещён: {rej}\nчини по находкам выше (reword/размеры/"
              f"custom-правки) и пересобери — рабочего .pptx не существует, "
              f"пока гейт не зелёный.")
        raise SystemExit(1)
    print("дальше по quality-gates.md §4: рендер из /private/tmp (движок — "
          "scripts/render.py) + самопроверка B2–B13, потом показывать.")


if __name__ == "__main__":
    main()
