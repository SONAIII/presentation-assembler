#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""atoms.py — клоны дизайнерских атомов с палитры шаблона (скрытые слайды,
см. palette_lint.py).

Атом = byte-exact дизайнерская фигура, клонируемая на целевой слайд.
Единственный рантайм-потребитель после 0.21.0 — роль `connector` прямого
клон-таймлайна (lib/generate.py::_add_connector): плиточная генерация из
атомов ВЫПИЛЕНА (решение владельца 2026-07-23 — ставка на дизайнерские
слайды + custom). Прочие роли палитры — справка дизайнеру и задел под
будущие клон-структуры (карта ролей — palette_lint.py ROLES).

Правила клона: группу — целиком (никогда не resize группу, только move:
внутренние координаты живут в child-space); плоский шейп можно и двигать,
и растягивать.
"""
from __future__ import annotations

import json
import os

import generate as _gen   # clone-хелперы (deepcopy + renumber id + strip a16)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class AtomError(Exception):
    pass


def load_atoms(path=None) -> dict:
    path = path or os.path.join(ROOT, "layouts/eda/atoms.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _find_on_slide(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    return None


def find_atom(prs, atoms: dict, role: str):
    """Живой Shape атома на палитровом слайде (для клона или чтения стиля)."""
    cfg = atoms["roles"].get(role)
    if cfg is None:
        raise AtomError(f"роль {role!r} не в карте атомов (palette_lint ROLES)")
    name = cfg.get("shape") or cfg["shapes"][0]
    rec = atoms["atoms"][name]
    slide = prs.slides[rec["slide1"] - 1]
    sh = _find_on_slide(slide, name)
    if sh is None:
        raise AtomError(f"атом {name!r} не найден на палитровом слайде {rec['slide1']} "
                        f"(шаблон обновился? перегони palette_lint)")
    return sh


def clone_atom(prs, target_slide, atoms: dict, role: str, alloc):
    """Byte-exact клон атома на целевой слайд (crash-safe, id перенумерованы)."""
    src = find_atom(prs, atoms, role)
    return _gen.clone_shape_el(target_slide, src, alloc)
