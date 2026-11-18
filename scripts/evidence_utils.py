#!/usr/bin/env python3
"""Small deterministic helpers for presentation evidence.

This module knows nothing about slide layouts. It normalizes source records and
checks traceability before approved content is handed to presentation-assembler.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import date, datetime, timezone
from typing import Any, Iterable


@dataclass(frozen=True)
class Evidence:
    id: str
    kind: str
    provider: str
    locator: str
    retrieved_at: str
    period: dict[str, str] | None = None
    metric: str | None = None
    unit: str | None = None
    value: Any = None
    calculation: str | None = None
    caveats: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["caveats"] = list(self.caveats)
        return {key: value for key, value in data.items() if value is not None}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def period(from_date: str, to_date: str, label: str) -> dict[str, str]:
    start = date.fromisoformat(from_date)
    end = date.fromisoformat(to_date)
    if start > end:
        raise ValueError("period.from must not be after period.to")
    if not label.strip():
        raise ValueError("period.label must not be empty")
    return {"from": start.isoformat(), "to": end.isoformat(), "label": label.strip()}


def numeric_sum(rows: Iterable[dict[str, Any]], value_key: str) -> float:
    total = 0.0
    found = False
    for row in rows:
        raw = row.get(value_key)
        if raw is None:
            continue
        text = str(raw).replace(" ", "").replace(",", ".").replace("%", "")
        try:
            total += float(text)
            found = True
        except ValueError:
            continue
    if not found:
        raise ValueError(f"no numeric values in column {value_key!r}")
    return total


def percent_delta(current: float, previous: float) -> float | None:
    if previous == 0:
        return None
    return round((current - previous) / previous * 100, 1)


def validate_traceability(slides: list[dict[str, Any]],
                          evidence: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    evidence_ids = [item.get("id") for item in evidence]
    known = {item for item in evidence_ids if isinstance(item, str) and item}
    if len(known) != len(evidence_ids):
        errors.append("evidence ids must be present and unique")
    by_id = {item.get("id"): item for item in evidence if isinstance(item, dict)}
    allowed_kinds = {"file", "url", "connected_service", "user_input", "web", "other"}
    for item_id, item in by_id.items():
        for key in ("kind", "provider", "locator", "retrieved_at"):
            if not item.get(key):
                errors.append(f"evidence {item_id}: missing {key}")
        if item.get("kind") not in allowed_kinds:
            errors.append(f"evidence {item_id}: unsupported kind")
        try:
            stamp = datetime.fromisoformat(str(item.get("retrieved_at", "")).replace("Z", "+00:00"))
            if stamp.tzinfo is None:
                errors.append(f"evidence {item_id}: retrieved_at must include timezone")
        except ValueError:
            errors.append(f"evidence {item_id}: retrieved_at must be ISO-8601")

    for index, slide in enumerate(slides, start=1):
        refs = slide.get("evidence_ids")
        if not isinstance(refs, list):
            errors.append(f"slide {index}: evidence_ids must be an array")
            continue
        missing = [ref for ref in refs if ref not in known]
        if missing:
            errors.append(f"slide {index}: unknown evidence ids: {', '.join(missing)}")
        numerical = _contains_number(slide.get("heading")) or _contains_number(slide.get("content"))
        if numerical:
            if not refs:
                errors.append(f"slide {index}: numerical claim has no evidence")
            for ref in refs:
                item = by_id.get(ref, {})
                for key in ("period", "unit", "value", "calculation"):
                    if key not in item or item.get(key) in (None, ""):
                        errors.append(f"slide {index}: numerical evidence {ref} missing {key}")
                period_value = item.get("period")
                if isinstance(period_value, dict):
                    missing_period = [key for key in ("from", "to", "label") if not period_value.get(key)]
                    if missing_period:
                        errors.append(f"slide {index}: numerical evidence {ref} period missing {', '.join(missing_period)}")
                    else:
                        try:
                            if date.fromisoformat(period_value["from"]) > date.fromisoformat(period_value["to"]):
                                errors.append(f"slide {index}: numerical evidence {ref} period is reversed")
                        except (TypeError, ValueError):
                            errors.append(f"slide {index}: numerical evidence {ref} period dates must be ISO dates")
                elif period_value is not None:
                    errors.append(f"slide {index}: numerical evidence {ref} period must be an object")
                if not isinstance(item.get("unit"), str):
                    errors.append(f"slide {index}: numerical evidence {ref} unit must be a string")
                if not isinstance(item.get("calculation"), str):
                    errors.append(f"slide {index}: numerical evidence {ref} calculation must be a string")
                caveats = item.get("caveats")
                if not isinstance(caveats, list) or not all(isinstance(value, str) for value in caveats):
                    errors.append(f"slide {index}: numerical evidence {ref} caveats must be an array of strings")
    return errors


def _contains_number(value: Any) -> bool:
    if isinstance(value, str):
        return any(char.isdigit() for char in value)
    if isinstance(value, dict):
        return any(_contains_number(item) for item in value.values())
    if isinstance(value, list):
        return any(_contains_number(item) for item in value)
    return False
