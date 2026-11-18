#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""images.py — photo-swap: замена дизайнерского фото на своё при нетронутой вёрстке.

Меняем ТОЛЬКО картинку внутри Picture-шейпа: новый image-парт + repoint
`a:blip r:embed`. Геометрия шейпа, roundRect-скругления, порядок отрисовки —
дизайнерские, не трогаем. Дизайнерский `a:srcRect` (кроп) заточен под ИХ
картинку → пересчитываем под аспект новой: center-crop cover-fit (видимая
область заполняет шейп без искажения). Машинный чек: аспект видимой области
== аспект шейпа (eps 2%) — в отчёт, не на глаз.

Осиротевшую rel старого фото чистим ТОЛЬКО если в XML слайда не осталось
ссылок (пара пикч может делить парт); неиспользуемая rel легальна, но
бесплатно раздувает файл.
"""
from __future__ import annotations

import os

from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def find_picture(slide, name: str):
    """Picture-шейп по имени (find_shape_by_name из fill — только текст)."""
    for sh in slide.shapes:
        if sh.name == name and sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return sh
    return None


def cover_fit_crop(img_w: int, img_h: int, shape_w: int, shape_h: int) -> dict:
    """Доли кропа {l,r,t,b} center-crop'а: видимая область = аспект шейпа."""
    img_a, shape_a = img_w / img_h, shape_w / shape_h
    if img_a > shape_a:                      # картинка шире — режем бока
        cut = 1.0 - shape_a / img_a
        return {"l": cut / 2, "r": cut / 2, "t": 0.0, "b": 0.0}
    cut = 1.0 - img_a / shape_a              # выше — режем верх/низ
    return {"l": 0.0, "r": 0.0, "t": cut / 2, "b": cut / 2}


def _set_srcrect(pic_el, crop: dict) -> None:
    blip_fill = pic_el.blipFill
    old = blip_fill.find(qn("a:srcRect"))
    if old is not None:
        blip_fill.remove(old)
    if not any(v > 1e-4 for v in crop.values()):
        return
    src = etree.SubElement(blip_fill, qn("a:srcRect"))
    blip = blip_fill.find(qn("a:blip"))
    blip.addnext(src)                        # схема: blip, затем srcRect
    for k, v in crop.items():
        if v > 1e-4:
            src.set(k, str(int(round(v * 100000))))


def swap_picture(slide, pic, image_path: str) -> dict:
    """Заменить картинку шейпа pic на image_path. Возвращает отчёт с чеком."""
    if not os.path.exists(image_path):
        return {"swapped": False, "error": f"файл не найден: {image_path}"}
    with Image.open(image_path) as im:
        iw, ih = im.size

    blip = pic._element.blipFill.find(qn("a:blip"))
    old_rid = blip.get(qn("r:embed"))
    image_part, rid = slide.part.get_or_add_image_part(image_path)
    blip.set(qn("r:embed"), rid)

    crop = cover_fit_crop(iw, ih, pic.width, pic.height)
    _set_srcrect(pic._element, crop)

    # осиротевшая rel: старый rId больше не встречается в XML слайда → дропаем
    if old_rid and old_rid != rid:
        xml = etree.tostring(slide.part._element).decode()
        if f'"{old_rid}"' not in xml:
            slide.part.drop_rel(old_rid)

    # машинный чек: аспект видимой области == аспект шейпа; DPI — честный флаг
    vis_a = (iw * (1 - crop["l"] - crop["r"])) / (ih * (1 - crop["t"] - crop["b"]))
    shape_a = pic.width / pic.height
    ok = abs(vis_a - shape_a) / shape_a < 0.02
    eff_dpi = iw * (1 - crop["l"] - crop["r"]) / (pic.width / 914400)
    rep = {"swapped": True, "image": os.path.basename(image_path),
           "px": [iw, ih], "crop": {k: round(v, 3) for k, v in crop.items()},
           "aspect_ok": ok, "eff_dpi": int(eff_dpi)}
    if eff_dpi < 72:                     # калибровка: дизайнерский full-bleed = 90,
        rep["low_res"] = True            # пул 79–90; ниже 72 — заметное мыло
    return rep


def swap_image_slots(slide, variant: dict, photos: dict) -> dict:
    """Все фото-слоты варианта: photos = {slot: имя-файла-или-путь (уже локальный)}.
    Слот без значения — passthrough (дизайнерское фото остаётся)."""
    report = {}
    for name, slot in (variant.get("image_slots") or {}).items():
        val = photos.get(name)
        if not val:
            report[name] = {"swapped": False, "passthrough": True}
            continue
        pic = find_picture(slide, slot["match_name"])
        if pic is None:
            report[name] = {"swapped": False,
                            "error": f"Picture «{slot['match_name']}» не найден"}
            continue
        report[name] = swap_picture(slide, pic, val)
    return report
