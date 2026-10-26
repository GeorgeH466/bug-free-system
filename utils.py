"""Small, dependency-free helpers shared by the Service Tracker modules."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable

ISO_DATE = "%Y-%m-%d"
DATE_INPUT_FORMATS = (
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d/%m/%y",
    "%d %b %Y",
    "%d %B %Y",
    "%Y/%m/%d",
)
TIME_INPUT_FORMATS = ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p")


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def parse_date(value: Any) -> date | None:
    """Parse a database value, Excel cell, or user-entered date.

    ISO dates are preferred. Ambiguous numeric dates are interpreted day-first,
    which is the usual convention for the target UK audience.
    """
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        # Excel serial date (openpyxl normally converts these already).
        serial = float(value)
        if 1 <= serial <= 2958465:
            return date(1899, 12, 30) + timedelta(days=int(serial))
        raise ValueError("Date is outside the supported calendar range.")

    text = str(value).strip()
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for fmt in DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError("Use a valid calendar date (YYYY-MM-DD).")


def normalize_date(value: Any) -> str | None:
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else None


def parse_time(value: Any) -> str | None:
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0).isoformat()
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{int(value.hour):02d}:{int(value.minute):02d}:{int(getattr(value, 'second', 0)):02d}"
    text = str(value).strip()
    for fmt in TIME_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt).time().replace(microsecond=0).isoformat()
        except ValueError:
            continue
    raise ValueError("Use a valid time (HH:MM).")


def parse_number(value: Any, *, default: float | None = None) -> float | None:
    """Parse a number, tolerating currency symbols, commas and accounting blanks."""
    if is_blank(value):
        return default
    if isinstance(value, bool):
        raise ValueError("Expected a number.")
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = re.sub(r"[\s,£$€]", "", text)
    if not text:
        return default
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", text):
        raise ValueError("Enter a valid number.")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("Enter a valid number.") from exc
    return float(-number if negative else number)


def week_bounds(reference: date | str | None = None) -> tuple[date, date]:
    """Return Monday and Sunday of the week containing ``reference``."""
    day = parse_date(reference) if reference else date.today()
    assert day is not None
    start = day - timedelta(days=day.weekday())
    return start, start + timedelta(days=6)


def format_currency(value: Any) -> str:
    number = parse_number(value, default=0.0) or 0.0
    return f"£{number:,.2f}"


def format_date(value: Any, *, fallback: str = "—") -> str:
    try:
        parsed = parse_date(value)
    except (TypeError, ValueError):
        parsed = None
    return parsed.strftime("%d %b %Y") if parsed else fallback


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value).strip())
    return cleaned.strip("._") or "file"


def chunked(items: Iterable[Any], size: int) -> Iterable[list[Any]]:
    batch: list[Any] = []
    for item in items:
        batch.append(item)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch
