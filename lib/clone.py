#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""clone.py — безопасный клон слайда внутри презентации.

Снимает v0-запрет «один дизайнерский слайд на дек»: дек собирается из КЛОНОВ
пристинных дизайнерских слайдов, оригиналы дропаются в конце (keep_ordered).

Наивный дубль ломает файл (Mac PowerPoint «Repair») — каталог паттернов в
~/.claude/skills/pptx-crash-check/known_issues.md:
  #1 orphan r:id — rels клона копируем С СОХРАНЕНИЕМ rId (в XML слайда зашиты
     r:embed/r:id; пере-нумерация даст orphan refs); notesSlide не копируем;
  #3 дубль слайдового <p14:creationId> → свежий val (max+1 по всему деку);
  #4 дубль шейпового <a16:creationId> GUID кросс-слайд → свежие детерминированные
     uuid5, КОНСИСТЕНТНО внутри слайда (mc:Choice/mc:Fallback делят GUID намеренно);
  #5 дубль <p14:modId> → свежие val;
  #6 шаринг chart/OLE-партов двумя слайдами НЕ ПОБЕЖДЁН (дубль партов не помог) →
     чарт-слайд клонируем не больше ОДНОГО раза на дек: оригинал дропается,
     единственный клон держит парт единолично, шаринга не возникает.
"""
from __future__ import annotations

import copy
import uuid

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.package import _Relationship
from pptx.parts.slide import SlidePart

A16_CREATION = "{http://schemas.microsoft.com/office/drawing/2014/main}creationId"
P14_CREATION = "{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId"
P14_MODID = "{http://schemas.microsoft.com/office/powerpoint/2010/main}modId"

# rels этих типов указывают на парты, которые Mac PP требует приватными на слайд
_EXCLUSIVE_RELTYPES = ("chart", "oleObject", "package")


def slide_needs_exclusive_parts(slide) -> bool:
    """True, если слайд держит chart/OLE-парт (паттерн #6: делить нельзя)."""
    return any(any(k in rel.reltype for k in _EXCLUSIVE_RELTYPES)
               for rel in slide.part.rels.values())


def _fresh_val(prs) -> int:
    """max+1 по всем слайдам дека — детерминированный уникальный val для p14."""
    mx = 0
    for s in prs.slides:
        for el in s.part._element.iter():
            if el.tag in (P14_CREATION, P14_MODID):
                try:
                    mx = max(mx, int(el.get("val") or 0))
                except ValueError:
                    pass
    return mx + 1


def _remap_ids(prs, slide_el, partname: str) -> None:
    """Паттерны #3/#4/#5: свежие идентификаторы клону, консистентно внутри слайда."""
    next_val = _fresh_val(prs)
    guid_map: dict[str, str] = {}
    for el in slide_el.iter():
        if el.tag in (P14_CREATION, P14_MODID):
            el.set("val", str(next_val))
            next_val += 1
        elif el.tag == A16_CREATION:
            old = el.get("id")
            if old:
                if old not in guid_map:
                    fresh = uuid.uuid5(uuid.NAMESPACE_OID, f"{old}@{partname}")
                    guid_map[old] = "{%s}" % str(fresh).upper()
                el.set("id", guid_map[old])


def clone_slide(prs, index: int) -> int:
    """Клонировать слайд index (0-based) в конец презентации. Возвращает индекс клона.

    Клон = deepcopy XML-парта + копия rels (те же target-парты, те же rId,
    без notesSlide) + ремап идентификаторов + новый sldId. Layout/theme/фото
    остаются ОБЩИМИ партами — это легально и штатно для PowerPoint; приватности
    требуют только chart/OLE (см. slide_needs_exclusive_parts — гейт у вызывающего).
    """
    src_part = prs.slides[index].part
    pkg = src_part.package
    partname = pkg.next_partname("/ppt/slides/slide%d.xml")

    new_el = copy.deepcopy(src_part._element)
    _remap_ids(prs, new_el, str(partname))
    new_part = SlidePart(partname, CT.PML_SLIDE, pkg, new_el)

    # rels копируем с СОХРАНЕНИЕМ rId (в XML клона зашиты r:embed/r:id — паттерн #1)
    for rel in src_part.rels.values():
        if rel.reltype == RT.NOTES_SLIDE:
            continue                      # заметки не тащим: их backlink — на оригинал
        target = rel.target_ref if rel.is_external else rel.target_part
        mode = RTM.EXTERNAL if rel.is_external else RTM.INTERNAL
        new_part.rels._rels[rel.rId] = _Relationship(
            new_part.rels._base_uri, rel.rId, rel.reltype,
            target_mode=mode, target=target)

    rid = prs.part.relate_to(new_part, RT.SLIDE)
    prs.slides._sldIdLst.add_sldId(rid)
    return len(prs.slides) - 1
