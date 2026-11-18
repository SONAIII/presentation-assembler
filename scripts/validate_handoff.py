#!/usr/bin/env python3
"""Validate an approved content handoff without third-party dependencies."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from evidence_utils import validate_traceability


FORBIDDEN_KEYS = {"family", "variant", "deckspec", "pptx", "donor_slide"}
MODES = {"single_slide", "deck", "restyle", "repair"}
DELIVERY = {"document", "presentation", "hybrid"}
VISUAL_INTENTS = {
    "time_trend", "comparison", "kpis", "process", "statement",
    "screenshots", "table", "timeline", "people",
}


def validate(data: dict) -> list[str]:
    errors: list[str] = []
    leaked_anywhere = sorted(FORBIDDEN_KEYS.intersection(_all_keys(data)))
    if leaked_anywhere:
        errors.append("design keys are forbidden in handoff: " + ", ".join(leaked_anywhere))
    required = {
        "schema_version", "status", "mode", "brief",
        "main_conclusion", "slides", "evidence",
    }
    missing = sorted(required - data.keys())
    if missing:
        errors.append("missing top-level keys: " + ", ".join(missing))
        return errors

    if data["schema_version"] != 1:
        errors.append("schema_version must be 1")
    if data["status"] != "approved":
        errors.append("status must be approved before handoff")
    if data["mode"] not in MODES:
        errors.append("mode must be single_slide, deck, restyle or repair")

    brief = data.get("brief")
    if not isinstance(brief, dict):
        errors.append("brief must be an object")
    else:
        for key in ("goal", "audience", "delivery", "template"):
            if not brief.get(key):
                errors.append(f"brief.{key} must not be empty")
        if brief.get("delivery") not in DELIVERY:
            errors.append("brief.delivery has an unsupported value")

    if not str(data.get("main_conclusion", "")).strip():
        errors.append("main_conclusion must not be empty")

    slides = data.get("slides")
    evidence = data.get("evidence")
    if not isinstance(slides, list) or not slides:
        errors.append("slides must be a non-empty array")
        slides = []
    if not isinstance(evidence, list):
        errors.append("evidence must be an array")
        evidence = []

    if data.get("mode") == "single_slide" and len(slides) != 1:
        errors.append("single_slide handoff must contain exactly one slide")

    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            errors.append(f"slide {index} must be an object")
            continue
        leaked = sorted(FORBIDDEN_KEYS.intersection(_all_keys(slide)))
        if leaked:
            errors.append(f"slide {index}: design keys are forbidden: {', '.join(leaked)}")
        for key in ("id", "heading", "content", "visual_intent", "evidence_ids"):
            if key not in slide:
                errors.append(f"slide {index}: missing {key}")
        if slide.get("visual_intent") not in VISUAL_INTENTS:
            errors.append(f"slide {index}: unsupported visual_intent")

    errors.extend(validate_traceability(slides, evidence))
    return errors


def _all_keys(value):
    keys = set()
    if isinstance(value, dict):
        for key, item in value.items():
            keys.add(key)
            keys.update(_all_keys(item))
    elif isinstance(value, list):
        for item in value:
            keys.update(_all_keys(item))
    return keys


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("handoff", type=Path)
    args = parser.parse_args()
    try:
        data = json.loads(args.handoff.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"BLOCK: cannot read handoff: {exc}", file=sys.stderr)
        return 2
    errors = validate(data)
    if errors:
        print("BLOCK: invalid approved-content handoff", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    print("PASS: approved-content handoff is valid")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
