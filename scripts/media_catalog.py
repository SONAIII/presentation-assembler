#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""media_catalog.py — опись изображений: из чужого .pptx или из папки пользователя.

Зачем (user-фидбек 2026-07-07, медиакит): пользователю не хватает картинок из
одобренного пула — свои скрины/фото/лого он приносит ЛИБО папкой, ЛИБО внутри
черновой презентации. Скрипт делает оба источника пригодными для photo-swap:

  * .pptx  → извлекает все картинки из ppt/media В ПРИВЯЗКЕ К СЛАЙДАМ
             (какая картинка где стояла и каким размером) в папку
             `<имя>_media/` + manifest.json; имена файлов = s<слайд>_<парт>;
  * папка  → опись на месте (файлы не трогаются): пиксели, аспект, вес.

Манифест/таблица — сырьё для выбора: аспект кандидата сверяется со слотом
(каталог слотов: SKILL.md «Photo-swap»), финальный отбор — глазами (Read),
кроп и dpi-предупреждения делает движок при сборке.

Использование:
  python3 scripts/media_catalog.py draft.pptx [-o outdir]   # извлечь + манифест
  python3 scripts/media_catalog.py ~/Pictures/approved/     # опись папки
  ... --json                                                # JSON в stdout

Флаги в описи: tiny (< 200px по большей стороне — иконка, не фото),
vector (emf/wmf — превью OLE-объектов, как фото непригодны).
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

try:
    from PIL import Image
except ImportError:                                   # PIL обязателен движку;
    Image = None                                      # опись выживет и без него

EMU = 914400
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"}
VECTOR_EXT = {".emf", ".wmf", ".svg"}
R_EMBED = qn("r:embed")


def _px(blob: bytes):
    if Image is None:
        return None
    try:
        with Image.open(io.BytesIO(blob)) as im:
            return list(im.size)
    except Exception:
        return None


def _flags(px, ext):
    flags = []
    if ext in VECTOR_EXT:
        flags.append("vector")
    elif px and max(px) < 200:
        flags.append("tiny")
    return flags


def _walk_pictures(shapes):
    """Picture-шейпы рекурсивно (группы: координаты детей — в child-space,
    но нам нужен только размер, он в EMU абсолютный после масштаба группы —
    для описи достаточно размера самого шейпа)."""
    for sh in shapes:
        if sh.shape_type == 6:                        # GROUP
            yield from _walk_pictures(sh.shapes)
        elif sh.shape_type == 13:                     # PICTURE
            yield sh


def catalog_pptx(path: str, outdir: str | None):
    prs = Presentation(path)
    stem = os.path.splitext(os.path.basename(path))[0]
    outdir = outdir or os.path.join(os.path.dirname(os.path.abspath(path)),
                                    f"{stem}_media")
    entries = {}                                      # partname → entry
    for si, slide in enumerate(prs.slides, 1):
        # display-размеры Picture-шейпов этого слайда по rId
        pic_sizes = {}
        for pic in _walk_pictures(slide.shapes):
            blip = pic._element.find(f".//{qn('a:blip')}")
            if blip is not None and blip.get(R_EMBED):
                pic_sizes.setdefault(blip.get(R_EMBED), []).append(
                    [round(pic.width / EMU, 2), round(pic.height / EMU, 2)])
        # ВСЕ blip-ссылки слайда (pic + заливки картинкой)
        for blip in slide.part._element.findall(f".//{qn('a:blip')}"):
            rid = blip.get(R_EMBED)
            if not rid or rid not in slide.part.rels:
                continue
            part = slide.part.rels[rid].target_part
            ext = os.path.splitext(str(part.partname))[1].lower()
            e = entries.get(part.partname)
            if e is None:
                px = _px(part.blob)
                e = entries[part.partname] = {
                    "part": os.path.basename(str(part.partname)),
                    "ext": ext.lstrip("."), "px": px,
                    "aspect": round(px[0] / px[1], 3) if px and px[1] else None,
                    "kb": round(len(part.blob) / 1024),
                    "flags": _flags(px, ext), "used_on": [],
                    "_blob": part.blob,
                }
            use = {"slide": si,
                   "kind": "picture" if rid in pic_sizes else "fill"}
            if rid in pic_sizes:
                use["display_in"] = max(pic_sizes[rid],
                                        key=lambda wh: wh[0] * wh[1])
            if not any(u["slide"] == si and u["kind"] == use["kind"]
                       for u in e["used_on"]):
                e["used_on"].append(use)

    items = sorted(entries.values(),
                   key=lambda e: (e["used_on"][0]["slide"], e["part"]))
    os.makedirs(outdir, exist_ok=True)
    for e in items:
        e["file"] = f"s{e['used_on'][0]['slide']:02d}_{e['part']}"
        with open(os.path.join(outdir, e["file"]), "wb") as f:
            f.write(e.pop("_blob"))
    manifest = {"source": os.path.abspath(path), "outdir": os.path.abspath(outdir),
                "images": items}
    with open(os.path.join(outdir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    return manifest


def catalog_dir(path: str):
    items = []
    for name in sorted(os.listdir(path)):
        ext = os.path.splitext(name)[1].lower()
        if ext not in IMG_EXT | VECTOR_EXT:
            continue
        full = os.path.join(path, name)
        px = None
        if ext in IMG_EXT:
            try:
                with open(full, "rb") as f:
                    px = _px(f.read())
            except Exception:
                px = None
        items.append({"file": name, "ext": ext.lstrip("."), "px": px,
                      "aspect": round(px[0] / px[1], 3) if px and px[1] else None,
                      "kb": round(os.path.getsize(full) / 1024),
                      "flags": _flags(px, ext) + ([] if px or ext in VECTOR_EXT
                                                  else ["unreadable"])})
    return {"source": os.path.abspath(path), "images": items}


def render_table(manifest: dict) -> str:
    out = [f"ОПИСЬ ИЗОБРАЖЕНИЙ: {manifest['source']}", "-" * 72]
    if not manifest["images"]:
        out.append("изображений не найдено")
        return "\n".join(out)
    for e in manifest["images"]:
        px = f"{e['px'][0]}×{e['px'][1]}" if e.get("px") else "?"
        asp = f" аспект {e['aspect']}" if e.get("aspect") else ""
        fl = f"  [{','.join(e['flags'])}]" if e.get("flags") else ""
        uses = ""
        if e.get("used_on"):
            parts = []
            for u in e["used_on"]:
                d = (f" {u['display_in'][0]}×{u['display_in'][1]}\""
                     if u.get("display_in") else "")
                parts.append(f"с.{u['slide']} {u['kind']}{d}")
            uses = "  ← " + "; ".join(parts)
        out.append(f"  {e['file']:34.34} {px:>10}{asp} {e['kb']:>6}КБ{fl}{uses}")
    if manifest.get("outdir"):
        out.append(f"\nизвлечено в: {manifest['outdir']} (+ manifest.json)")
    out.append("дальше: кандидатов смотреть глазами, аспект сверять со слотом "
               "(SKILL.md «Photo-swap»); кроп и dpi-чек сделает сборка")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", help=".pptx или папка с изображениями")
    ap.add_argument("-o", "--out", default=None,
                    help="куда извлекать медиа из .pptx (default: <имя>_media рядом)")
    ap.add_argument("--json", action="store_true", help="JSON в stdout")
    a = ap.parse_args()
    if os.path.isdir(a.source):
        manifest = catalog_dir(a.source)
    elif a.source.lower().endswith(".pptx"):
        manifest = catalog_pptx(a.source, a.out)
    else:
        raise SystemExit(f"не .pptx и не папка: {a.source}")
    print(json.dumps(manifest, ensure_ascii=False, indent=1) if a.json
          else render_table(manifest))


if __name__ == "__main__":
    main()
