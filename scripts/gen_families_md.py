#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_families_md.py — генерирует references/families-<tdir>.md из layouts/<tdir>/*.json.

Каталог семейств — ЕДИНСТВЕННОЕ место с точными счётчиками (сколько семейств,
вариантов, слайдов в шаблоне): проза SKILL.md/TESTING.md ссылается сюда и чисел
не дублирует. Пер-шаблонный: `--template <tdir>` (дефолт eda). Запускать после
правки layouts/ или бампа шаблона:

    python3 scripts/gen_families_md.py                      # families-eda.md (дефолт)
    python3 scripts/gen_families_md.py --template gorod     # families-gorod.md
    python3 scripts/gen_families_md.py --check              # exit 1, если ЛЮБОЙ файл разошёлся

Откуда что берётся:
  * механика (тир, варианты, номера слайдов #N, метки GEN/chart/фото) — из спек
    layouts/eda/*.json, всегда актуальна;
  * контрактные описания роутеров (metrics/timeline/takeaways/чарты/reasons) —
    из OVERRIDES ниже: _doc спеки описывает дизайнерский слайд, а override —
    контракт deckspec. Меняется контракт → обнови override здесь же;
  * прочие описания — первые предложения _doc спеки (целиком, без обрезки
    посреди слова).

Вывод детерминированный (без дат) — диффы в VCS осмысленные.
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
LAYOUTS_ROOT = os.path.join(SKILL_DIR, "layouts")
DEFAULT_TDIR = "eda"
NOT_FAMILIES = {"_canvas", "_inventory", "brand-tokens", "atoms", "custom-profile"}
DESC_LIMIT = 240   # мягкий потолок описания; режем ТОЛЬКО по границе предложения

# Человекочитаемое имя шаблона для заголовка каталога; фолбэк — сам tdir.
# Новый шаблон → строка здесь (онбординг шаблонов делает команда скилла).
TDIR_TITLE = {"eda": "Еда", "gor-services": "Городские сервисы",
              "ridetech-khural": "Хурал Райдтеха", "committee": "Комитет"}


def layout_dir(tdir: str) -> str:
    return os.path.join(LAYOUTS_ROOT, tdir)


def out_path(tdir: str) -> str:
    return os.path.join(SKILL_DIR, "references", f"families-{tdir}.md")


def known_tdirs() -> list[str]:
    """Все шаблоны = подпапки layouts/ (де-факто LAYOUT_DIRS движка). Новый пак
    кладёт layouts/<tdir>/ → --check подхватывает его families-<tdir>.md сам."""
    return sorted(n for n in os.listdir(LAYOUTS_ROOT)
                  if os.path.isdir(os.path.join(LAYOUTS_ROOT, n)))


# Контракты deckspec поверх _doc (описание дизайнерского слайда → как его звать
# из deckspec). Это ДОМ контрактов слотов: SKILL.md «Что умеет» ссылается сюда
# по 1–2 строки на семейство и числа/детали в прозе не дублирует. Меняется
# контракт → правь override здесь, потом `gen_families_md.py`. Словарь по tdir
# (структура OVERRIDES = {"<tdir>": {family: contract}}); имена семейств — общий
# словарь паков (решение владельца №1: костяк один).
OVERRIDES = {"eda": {
    "datatable": "Одна таблица (#36, донор 9×4): `slots.table` = rows[r][c] "
                 "(строка или список строк). Сетка УЖИМАЕТСЯ под контент "
                 "(2026-07-20): лишние строки/колонки удаляются физически, "
                 "болванки чистятся, затем усечённая сетка РАСТЯГИВАЕТСЯ на "
                 "донорскую высоту с ростом кегля до потолка (stretch, общий "
                 "режим 2026-07-29 — компактная таблица не липнет к заголовку); "
                 "подсветки дизайнера живут только над полной сеткой "
                 "(`slots.highlights: false` — снять всегда); выравнивание "
                 "колонок по типу контента: текст → влево, числа → вправо "
                 "(заголовок — как колонка), числа со словами («18 тыс») "
                 "считаются текстом; строки подавай осмысленным порядком "
                 "(группировка/убывание), не как пришло",
    "toc": "Содержание/раздел-с-оглавлением Еды: метка «Раздел N» + многострочный "
           "список пунктов (один пункт = один абзац). `slots.active_index` (0-based) "
           "выделяет пункт текущего раздела дизайнерским стилем — агенда-per-section "
           "длинных деков, слайд на раздел",
    "metrics": "Стат-сетка: контракт `slots.items` = [{value, label}]. Роутер по N "
               "и глиф-фиту: N=3 и влезает → FILL дизайнерской 3-карточной сетки "
               "(первая карточка жёлтая); иначе → GENERATE clone-elastic на бренде. "
               "Однородные значения получают единый кегль автоматически",
    "timeline": "Таймлайн: `slots.stages` = [{date: [метка, год], desc}]. N=3/4 → "
                "FILL дизайнерской кривой; N=5–6 → GENERATE прямой линии (атом "
                "connector, описания через одно сверху/снизу); N≥7 → честная ошибка "
                "«дели на два слайда». Узкий бокс описания: слова длиннее ~12 букв "
                "не влезают → reword. Дата-боксы авторасширены под 4-значный год",
    "takeaways": "Слайд «основные выводы»: заголовок + слот `intro` (строка на жёлтом "
                 "баннере) + `blocks` = 3 развёрнутых вывода {name, desc}. Контент "
                 "обязан заполнить колонки (fill-ratio), иначе честная ошибка: "
                 "реворди под донора или custom (плитки выпилены 0.21.0)",
    "linechart": "Слайд «график с пояснением»: заголовок + подзаголовок + подпись оси "
                 "+ LINE-чарт с ДАТНОЙ осью — `charts.chart` = {categories: ISO-даты "
                 "«2026-03», series}; легенда авто по рангу финального значения; "
                 "комментарий справа. Заголовок может лечь в 2 строки — тогда "
                 "подзаголовок ОБЯЗАН быть пуст (машинный чек `multiline_requires_empty`)",
    "charts_compare": "Два COLUMN-чарта на белых плашках для сравнения: подписи над "
                      "чартами, дельта-бейджи (+50%/−29%), комменты [лид, тело]; "
                      "`charts.chart_left/chart_right`",
    "charts_duo": "BAR (категории-строки) + COLUMN (до 3 серий) на плашках; ручная "
                  "легенда `legend_1..3` = имена серий правого чарта; "
                  "`charts.chart_left/chart_right`; BAR-категории подавай по убыванию",
    "columns": "Семейство «колонки-персонажи» Еды: заголовок + N блоков "
               "`blocks` = [{name, desc}] + фото. Запись БЕЗ variant — роутер "
               "примеряет дизайнерские варианты (N совпал, fill-ratio ≥ 0.5) → "
               "FILL; не заполнил — честная ошибка: реворд или custom (плитки "
               "выпилены 0.21.0)",
}}

# Семейства без собственного fill-map, собираемые движком, — в каталоге обязаны
# быть видны. Пусто с 0.21.0: плиточное `reasons` выпилено (решение владельца
# 2026-07-23, ставка — дизайнерские слайды + custom). Словарь по tdir.
SYNTHETIC_ROWS = {"eda": []}


def first_sentences(text: str, limit: int = DESC_LIMIT) -> str:
    """Первые предложения _doc целиком; никогда не режем посреди слова."""
    text = " ".join((text or "").split())
    if len(text) <= limit:
        return text.rstrip(". ")
    out = ""
    for part in text.split(". "):
        cand = (out + ". " if out else "") + part
        if out and len(cand) > limit:
            break
        out = cand
    if len(out) > limit:                      # первое же предложение длиннее лимита
        cut = out.rfind(" ", 0, limit)
        out = out[: cut if cut > 0 else limit].rstrip(",;: ") + "…"
    return out.rstrip(". ")


def md_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def load_families(tdir: str) -> list[dict]:
    fams = []
    for p in sorted(glob.glob(os.path.join(layout_dir(tdir), "*.json"))):
        base = os.path.basename(p)[:-5]
        if base in NOT_FAMILIES:
            continue
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        fams.append(d)
    return sorted(fams, key=lambda d: d["family"])


def family_row(d: dict, tdir: str) -> tuple[str, str, str, str, str]:
    fam = d["family"]
    tier = d.get("kind") or ("content" if d.get("card_unit") else "?")
    if d.get("variants"):
        variants = ", ".join(f"`{v['name']}` (#{v['_hint'] + 1})"
                             for v in d["variants"])
    elif d.get("card_unit") and d.get("ready_slides"):
        variants = f"3-карточная сетка (#{d['ready_slides'][0]['_hint'] + 1})"
    else:
        variants = ""
    gen = []
    if d.get("generate") or d.get("card_unit"):
        gen.append("✓")
    if any(v.get("chart_slots") for v in d.get("variants", [])):
        gen.append("✓chart")
    desc = OVERRIDES.get(tdir, {}).get(fam) or first_sentences(d.get("_doc", ""))
    return fam, tier, variants, " ".join(gen), desc


def image_slot_rows(fams: list[dict]) -> list[tuple[str, str]]:
    """Слот-карта photo-swap ИЗ спек (поле image_slots) — не пишется руками.
    Возвращает [(«семейство · вариант», «`слот` — что это · …»)]."""
    rows = []
    for d in fams:
        fam = d["family"]
        for v in d.get("variants", []):
            islots = v.get("image_slots")
            if not islots:
                continue
            parts = []
            for name, meta in islots.items():
                doc = (meta or {}).get("_doc", "") if isinstance(meta, dict) else ""
                doc = " ".join(doc.split())
                # подпись позиции: без пояснительного хвоста после «—» и без «(…)»
                doc = doc.split(" — ")[0].split(" (")[0].strip(" .")
                parts.append(f"`{name}`" + (f" {doc}" if doc else ""))
            rows.append((f"`{fam}` · {v['name']}", " · ".join(parts)))
    return rows


def render(tdir: str) -> str:
    fams = load_families(tdir)
    rows = [family_row(d, tdir) for d in fams]
    n_variants = sum(len(d.get("variants", [])) or 1 for d in fams)
    inv_path = os.path.join(layout_dir(tdir), "_inventory.json")
    slide_count = "?"
    if os.path.exists(inv_path):
        with open(inv_path, encoding="utf-8") as f:
            slide_count = json.load(f).get("slide_count", "?")
    title = TDIR_TITLE.get(tdir, tdir)
    # дефолтный шаблон печатается голой командой (файл eda байт-в-байт как прежде)
    gen_cmd = ("python3 scripts/gen_families_md.py"
               + ("" if tdir == DEFAULT_TDIR else f" --template {tdir}"))

    lines = [
        f"# Семейства шаблона {title} — меню готовых слайдов",
        "",
        f"**{len(fams)} семейств, {n_variants} дизайнерских вариантов** "
        f"(слайдов в шаблоне-доноре: {slide_count}, включая скрытую палитру "
        f"атомов).",
        "",
        f"Файл генерируется: `{gen_cmd}` (правки руками "
        "перетрутся; контрактные описания — в OVERRIDES генератора). Номера "
        "слайдов (#N, 1-based) — подсказка, резолв всегда по сигнатуре (anchor), "
        "не по индексу. `structural` = всегда FILL точь-в-точь; `content_fill` = "
        "FILL с фит-политиками по слотам; GEN ✓ = у семейства есть генеративная "
        "ветка; ✓chart = данные чарта редактируются через `charts` в deckspec.",
        "",
        "| Семейство | Тир | Варианты (слайд) | GEN | Что это |",
        "|---|---|---|---|---|",
    ]
    all_rows = [tuple(map(md_cell, r)) for r in rows]
    for srow in SYNTHETIC_ROWS.get(tdir, []):
        all_rows.append(tuple(map(md_cell, srow)))
    all_rows.sort(key=lambda r: r[0])
    for fam, tier, variants, gen, desc in all_rows:
        lines.append(f"| `{fam}` | {tier} | {variants} | {gen} | {desc} |")

    lines += [
        "",
        "## Фото-слоты (`photos` в deckspec)",
        "",
        "Карта сменных фото по вариантам — из поля `image_slots` спек (не руками). "
        "Ключ = имя слота в `photos`; значение = имя из пула шаблона или локальный "
        "путь. Слот без значения = passthrough (донорское фото остаётся). "
        "Геометрия и center-crop под аспект — движка; меняется только картинка.",
        "",
        "| Семейство · вариант | Слоты |",
        "|---|---|",
    ]
    for name, slots in image_slot_rows(fams):
        lines.append(f"| {md_cell(name)} | {md_cell(slots)} |")

    lines += [
        "",
        "Контракты слотов и ограничения — в таблицах выше (колонка «Что это» + "
        "карта фото-слотов); демо-контент каждого семейства — "
        "`examples/smoke_A_text.json` / `smoke_B_data.json` / `smoke_C_photo_gen.json` "
        "(вместе покрывают все семейства). Сборочный процесс, роутинг и STOP-точки — "
        "SKILL.md.",
        "",
    ]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=DEFAULT_TDIR,
                    help=f"tdir для записи (дефолт {DEFAULT_TDIR})")
    ap.add_argument("--check", action="store_true",
                    help="не писать, а сверить files ВСЕХ шаблонов со спеками")
    a = ap.parse_args()

    if a.check:
        bad = False
        for tdir in known_tdirs():
            text = render(tdir)
            op = out_path(tdir)
            rel = os.path.relpath(op, SKILL_DIR)
            cmd = ("python3 scripts/gen_families_md.py"
                   + ("" if tdir == DEFAULT_TDIR else f" --template {tdir}"))
            if not os.path.exists(op):
                print(f"{rel} не сгенерирован — {cmd}")
                bad = True
                continue
            with open(op, encoding="utf-8") as f:
                current = f.read()
            if current != text:
                print(f"{rel} разошёлся со спеками layouts/{tdir}/ — {cmd}")
                bad = True
            else:
                print(f"{rel} соответствует спекам ✓")
        sys.exit(1 if bad else 0)

    tdir = a.template
    if not os.path.isdir(layout_dir(tdir)):
        sys.exit(f"нет layouts/{tdir}/ — неизвестный шаблон {tdir!r} "
                 f"(есть: {', '.join(known_tdirs())})")
    text = render(tdir)
    op = out_path(tdir)
    with open(op, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"записан {os.path.relpath(op, SKILL_DIR)} "
          f"({len(text.splitlines())} строк)")


if __name__ == "__main__":
    main()
