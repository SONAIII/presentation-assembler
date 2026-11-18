"""Period parsing: human text + ISO ranges → (from_date, to_date)."""

from datetime import date, timedelta
import re


def period_to_dates(period: str) -> tuple[str, str]:
    """
    Convert human-readable period string to ISO date range.

    Examples:
        "прошлая неделя" → ("2026-08-11", "2026-08-17")  # last 7 days from today
        "10–16 августа 2026" → ("2026-08-10", "2026-08-16")
        "август 2026" → ("2026-08-01", "2026-08-31")
        "2026-08-10..2026-08-16" → ("2026-08-10", "2026-08-16")

    Returns: (from_date: YYYY-MM-DD, to_date: YYYY-MM-DD)
    """
    period = period.strip()
    today = date.today()

    # ISO range: "2026-08-10..2026-08-16"
    if ".." in period:
        parts = period.split("..")
        if len(parts) == 2:
            return (parts[0].strip(), parts[1].strip())

    # "прошлая неделя" / "last week"
    if "прошлая неделя" in period.lower() or "last week" in period.lower():
        to_d = today - timedelta(days=1)
        from_d = to_d - timedelta(days=6)
        return (from_d.isoformat(), to_d.isoformat())

    # "N дней" / "N days"
    match = re.search(r"(\d+)\s*(?:дн|day)", period, re.IGNORECASE)
    if match:
        days = int(match.group(1))
        to_d = today - timedelta(days=1)
        from_d = to_d - timedelta(days=days - 1)
        return (from_d.isoformat(), to_d.isoformat())

    # "10–16 августа 2026" / "Aug 10–16, 2026" / "августа 2026"
    # Кириллица месяцев
    _MONTHS_RU = {
        "янв": 1, "февр": 2, "март": 3, "апр": 4, "май": 5, "июнь": 6,
        "июль": 7, "авг": 8, "сент": 9, "окт": 10, "ноя": 11, "дек": 12,
        "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
        "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
    }

    # Try "10–16 августа 2026"
    match = re.search(r"(\d{1,2})\s*[–\-]\s*(\d{1,2})\s+(\w+)\s+(\d{4})", period)
    if match:
        day_from = int(match.group(1))
        day_to = int(match.group(2))
        month_str = match.group(3).lower()
        year = int(match.group(4))

        month = None
        for k, v in _MONTHS_RU.items():
            if month_str.startswith(k):
                month = v
                break

        if month:
            from_d = date(year, month, day_from)
            to_d = date(year, month, day_to)
            return (from_d.isoformat(), to_d.isoformat())

    # Try "августа 2026" (full month)
    match = re.search(r"(\w+)\s+(\d{4})", period)
    if match:
        month_str = match.group(1).lower()
        year = int(match.group(2))

        month = None
        for k, v in _MONTHS_RU.items():
            if month_str.startswith(k):
                month = v
                break

        if month:
            # Full month: from 1st to last day
            if month == 12:
                to_d = date(year + 1, 1, 1) - timedelta(days=1)
            else:
                to_d = date(year, month + 1, 1) - timedelta(days=1)
            from_d = date(year, month, 1)
            return (from_d.isoformat(), to_d.isoformat())

    # Fallback: assume ISO YYYY-MM-DD
    try:
        d = date.fromisoformat(period)
        return (d.isoformat(), d.isoformat())
    except ValueError:
        pass

    raise ValueError(f"Cannot parse period: {period}")


def _fmt_period_label(from_date: str, to_date: str) -> str:
    """Format date range for human display, e.g. '11–17 авг 2026'."""
    _MONTHS_RU = ["", "янв", "фев", "мар", "апр", "май", "июн",
                  "июл", "авг", "сен", "окт", "ноя", "дек"]
    try:
        d0 = date.fromisoformat(from_date)
        d1 = date.fromisoformat(to_date)
        if d0.month == d1.month:
            return f"{d0.day}–{d1.day} {_MONTHS_RU[d0.month]} {d0.year}"
        return f"{d0.day} {_MONTHS_RU[d0.month]} – {d1.day} {_MONTHS_RU[d1.month]} {d1.year}"
    except Exception:
        return f"{from_date} – {to_date}"
#!/usr/bin/env python3
