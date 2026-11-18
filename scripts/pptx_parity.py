#!/usr/bin/env python3
"""Fail-closed content parity gate for PPTX restyle workflows."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def walk_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)


def inventory(path: Path) -> dict:
    prs = Presentation(path)
    texts, notes, images = [], [], []
    for slide in prs.slides:
        for shape in walk_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                value = norm(shape.text_frame.text)
                if value:
                    texts.append(value)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        value = norm(cell.text)
                        if value:
                            texts.append(value)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    images.append(hashlib.sha256(shape.image.blob).hexdigest())
                except ValueError:
                    images.append(f"non-embedded:{shape.name}:{shape.width}x{shape.height}")
        try:
            value = norm(slide.notes_slide.notes_text_frame.text)
            if value:
                notes.append(value)
        except (AttributeError, KeyError):
            pass
    numbers = re.findall(r"(?<!\w)[+−-]?\d+(?:[.,]\d+)?%?", " ".join(texts))
    return {"slides": len(prs.slides), "texts": Counter(texts), "numbers": Counter(numbers),
            "notes": Counter(notes), "images": Counter(images)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare source and restyled PPTX content")
    parser.add_argument("source", type=Path)
    parser.add_argument("result", type=Path)
    parser.add_argument("--allow-extra", action="store_true", help="allow extra result content")
    args = parser.parse_args()
    source, result = inventory(args.source), inventory(args.result)
    missing = {}
    for key in ("texts", "numbers", "notes", "images"):
        delta = source[key] - result[key]
        if delta:
            missing[key] = dict(delta)
    extras = {}
    if not args.allow_extra:
        for key in ("texts", "numbers", "notes", "images"):
            delta = result[key] - source[key]
            if delta:
                extras[key] = dict(delta)
    report = {"source_slides": source["slides"], "result_slides": result["slides"],
              "missing": missing, "extras": extras}
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if missing or extras:
        print("BLOCK: PPTX content parity failed", file=sys.stderr)
        return 1
    print("PASS: text, numbers, images and speaker notes preserved")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
