#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""palette_lint.py — проверка палитры атомов шаблона + генерация карты атомов.

Палитра = СКРЫТЫЕ слайды шаблона с фигурами `atom:*` (слой адаптера шаблона).
Контракт палитры:
  * имя верхнего уровня `atom:*` уникально на всю палитру (одно имя = одна фигура);
  * атом клонируется ЦЕЛИКОМ (группа = один атом; внутренние имена свободные);
  * текстовые атомы несут образцовый текст (стиль живёт в ранах);
  * REQUIRED-набор ролей закрыт. С 0.21.0 (плитки выпилены, решение владельца
    2026-07-23) рантайм клонирует только `connector` (прямой клон-таймлайн) —
    REQUIRED сжат до него; прочие роли — информационная справка дизайнеру
    и задел под будущие клон-структуры, их отсутствие приёмку не блокирует.

Выдаёт:
  * lint-отчёт (exit 1 при ошибках контракта / недостающих ролях);
  * layouts/<t>/atoms.json — карта РОЛЬ → фигура (+params). Единственное место,
    где семантика привязана к именам палитры: переименования в PowerPoint не
    нужны, алиасы живут здесь (пример: chart_legend_label → atom:chart_line-label-*);
  * layouts/<t>/atoms.md — человекочитаемая справка (что докрутить дизайнеру).

  python3 scripts/palette_lint.py                # Еда: линт + atoms.json + atoms.md
  python3 scripts/palette_lint.py --check-only   # только линт, файлы не писать
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

DEFAULT_TDIR = "eda"            # шаблон по умолчанию; pptx/atoms-пути резолвятся по tdir
EMU = 914400

# Роль → имя фигуры на палитре (алиасы; prefix=True — повторяемая роль по префиксу).
# params — знание, нужное движку при клонировании (не выводимо из самой фигуры).
ROLES = {
    "card":               {"shape": "atom:card"},
    "badge":              {"shape": "atom:badge",
                           "params": {"number_scheme_first": "tx1",
                                      "number_scheme_rest": "accent3",
                                      "_doc": "бейджи #19 XML-идентичны, различается ТОЛЬКО "
                                              "цвет цифры: первая колонка tx1, остальные accent3"}},
    "block":              {"shape": "atom:block"},
    "block_bullets":      {"shape": "atom:block-bullets"},
    "value":              {"shape": "atom:value"},
    "label":              {"shape": "atom:label"},
    "pill":               {"shape": "atom:pill"},
    "stage_desc":         {"shape": "atom:stage_desc"},
    "connector":          {"shape": "atom:connector", "alt": ["atom:line"],
                           "_doc": "alt: Городские v2.0 держат палитру НА эталонных "
                                   "слайдах (не витриной) — линия таймлайна atom:line "
                                   "и есть канонный коннектор клон-таймлайна"},
    "arrow":              {"shape": "atom:arrow"},
    "divider":            {"shape": "atom:divider"},
    "icon_check":         {"shape": "atom:icon-check"},
    "icon_cross":         {"shape": "atom:icon-cross"},
    "chart_line":         {"shape": "atom:chart_line"},
    "chart_column":       {"shape": "atom:chart_column"},
    # chart_plate: атом atom:chart_shape СНЯТ пользователем из палитры (шаблон v11,
    # 2026-07-02) — плашки под чарты запечены в дизайнерских #40/#41; понадобится
    # генеративной чарт-композиции → белая заливка atom:card
    "chart_title":        {"shape": "atom:chart_title"},
    "chart_legend_label": {"shape": "atom:chart_line-label-", "prefix": True},
}
REQUIRED = ["connector"]   # 0.21.0: единственная роль, клонируемая рантаймом (таймлайн)
OPTIONAL_MISSING_OK = {"palette:marker"}


def kind_of(sh) -> str:
    st = sh.shape_type
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st == MSO_SHAPE_TYPE.CHART:
        return "chart"
    if st == MSO_SHAPE_TYPE.LINE:
        return "line"
    if sh.has_text_frame and sh.text_frame.text.strip():
        return "text"
    return "shape"


def first_run_style(sh):
    """(font, size_pt) первого рана с текстом — репорт, не резолв наследования."""
    if not sh.has_text_frame:
        return None, None
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                sz = r.font.size.pt if r.font.size else None
                return r.font.name, sz
    return None, None


def collect_palette(prs):
    """[(slide_idx0, hidden, [top-level shapes])] для слайдов с atom:*."""
    out = []
    for i, s in enumerate(prs.slides):
        tops = list(s.shapes)
        if any(sh.name.startswith(("atom:", "palette:")) for sh in tops):
            out.append((i, s._element.get("show") == "0", tops))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=DEFAULT_TDIR,
                    help=f"tdir шаблона (дефолт {DEFAULT_TDIR}); pptx резолвится по пину")
    ap.add_argument("--out", default=None, help="дефолт layouts/<tdir>/atoms.json")
    ap.add_argument("--md", default=None, help="дефолт layouts/<tdir>/atoms.md")
    ap.add_argument("--check-only", action="store_true")
    args = ap.parse_args()

    sys.path.insert(0, HERE)
    import fetch_template            # дистрибуция без .pptx: кэш → Я.Диск по пину
    from gen_families_md import TDIR_TITLE   # человекочитаемое имя шаблона (один реестр)
    tdir = fetch_template.TDIR_ALIASES.get(args.template.lower(), args.template.lower())
    tpl_path = fetch_template.resolve_template(tdir)
    out_path = args.out or os.path.join(ROOT, "layouts", tdir, "atoms.json")
    md_path = args.md or os.path.join(ROOT, "layouts", tdir, "atoms.md")
    title = TDIR_TITLE.get(tdir, tdir)

    prs = Presentation(tpl_path)
    palette = collect_palette(prs)
    errors, warns = [], []

    if not palette:
        print("FAIL: в шаблоне нет слайдов с фигурами atom:*")
        return 1

    atoms = {}   # top-level name -> record
    for idx0, hidden, tops in palette:
        # Две модели палитры: слайд-ВИТРИНА (все верхнеуровневые шейпы atom:*/
        # palette:* — Еда #44/#45) обязана быть скрыта, иначе утечёт в показ;
        # разметка IN-PLACE (atom:* точечно на эталонных ДОНОРАХ — Городские
        # v2.0 #29/32/33/34) легальна на видимых слайдах: донор и есть эталон,
        # его клон уносит атом в дек штатно.
        vitrine = all(sh.name.startswith(("atom:", "palette:")) for sh in tops)
        if not hidden and vitrine:
            errors.append(f"слайд-витрина {idx0 + 1} с атомами НЕ скрыт (утечёт в показ)")
        for sh in tops:
            if not sh.name.startswith("atom:"):
                continue
            if sh.name in atoms:
                warns.append(f"имя {sh.name!r} размечено на нескольких слайдах "
                             f"({atoms[sh.name]['slide1']} и {idx0 + 1}) — эталоном "
                             f"беру первое вхождение")
                continue
            font, size = first_run_style(sh)
            rec = {"slide1": idx0 + 1, "kind": kind_of(sh),
                   "w_in": round(sh.width / EMU, 3), "h_in": round(sh.height / EMU, 3)}
            if font or size:
                rec["font"], rec["size_pt"] = font, size
            if rec["kind"] == "group":
                rec["children"] = [{"name": c.name, "kind": kind_of(c)} for c in sh.shapes]
                if not any(c.has_text_frame and c.text_frame.text.strip() for c in sh.shapes) \
                        and "chart" not in sh.name and "icon" not in sh.name:
                    warns.append(f"группа {sh.name!r} без образцового текста внутри")
            elif rec["kind"] == "text" and not sh.text_frame.text.strip():
                warns.append(f"текстовый атом {sh.name!r} пуст — стиль ранов потеряется")
            atoms[sh.name] = rec

    # полнота ролей
    role_map = {}
    for role, cfg in ROLES.items():
        if cfg.get("prefix"):
            names = sorted(n for n in atoms if n.startswith(cfg["shape"]))
            if names:
                role_map[role] = {"shapes": names, **{k: v for k, v in cfg.items()
                                                      if k not in ("shape", "prefix")}}
            elif role in REQUIRED:
                errors.append(f"роль {role!r}: нет фигур с префиксом {cfg['shape']!r}")
            continue
        resolved = next((n for n in [cfg["shape"], *cfg.get("alt", [])] if n in atoms), None)
        if resolved is not None:
            role_map[role] = {**{k: v for k, v in cfg.items()
                                 if k not in ("prefix", "alt", "_doc")},
                              "shape": resolved}
        elif role in REQUIRED:
            errors.append(f"роль {role!r}: на палитре нет фигуры {cfg['shape']!r}"
                          + (f" (ни альтернатив {cfg['alt']})" if cfg.get("alt") else ""))
    unmapped = [n for n in atoms
                if not any(n == c.get("shape") or n in c.get("alt", [])
                           or (c.get("prefix") and n.startswith(c["shape"]))
                           for c in ROLES.values())]
    for n in unmapped:
        warns.append(f"атом {n!r} есть на палитре, но не привязан ни к одной роли (добавь в ROLES)")

    print(f"палитра: слайды {[i + 1 for i, _, _ in palette]}, атомов {len(atoms)}, "
          f"ролей закрыто {len(role_map)}/{len(ROLES)}")
    for e in errors:
        print(f"  [ERROR] {e}")
    for w in warns:
        print(f"  [WARN]  {w}")

    if not args.check_only:
        data = {"template_file": os.path.relpath(tpl_path, ROOT),
                "palette_slides_1based": [i + 1 for i, _, _ in palette],
                "roles": role_map, "atoms": atoms}
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(f"# Палитра атомов — {title}\n\n"
                    "Автогенерируется `scripts/palette_lint.py` из скрытых слайдов шаблона. "
                    "Роли назначаются в ROLES там же; в PowerPoint переименовывать не нужно.\n\n"
                    "| Роль | Фигура | Тип | Размер | Слайд |\n|---|---|---|---|---|\n")
            for role, cfg in role_map.items():
                names = cfg.get("shapes", [cfg.get("shape")])
                for n in names:
                    a = atoms[n]
                    f.write(f"| {role} | `{n}` | {a['kind']} | "
                            f"{a['w_in']}×{a['h_in']}\" | {a['slide1']} |\n")
            if warns:
                f.write("\n## Докрутить\n\n" + "".join(f"- {w}\n" for w in warns))
        print(f"написано: {os.path.relpath(out_path, ROOT)}, {os.path.relpath(md_path, ROOT)}")

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
