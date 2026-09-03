"""Turning whatever people typed into a spreadsheet cell into dates and times.

Real planning sheets are messy: dates are written "03/09/2026", "3. sep 2026"
or left as a real Excel date; times appear as "9", "9.00", "09:00-12:00" or as
an Excel time fraction. Everything here is best-effort and raises
:class:`ValueError` with a readable message when a cell cannot be understood.
"""

from __future__ import annotations

import datetime as dt
import re

# Excel stores dates as days since 1899-12-30 (the 1900 leap-year bug included).
EXCEL_EPOCH = dt.date(1899, 12, 30)

MONTHS = {
    # English
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12,
    "december": 12,
    # Danish (where it differs)
    "maj": 5, "okt": 10, "oktober": 10, "januar": 1, "februar": 2, "marts": 3,
    "juni": 6, "juli": 7, "december": 12,
}

# "9-12", "09:00 - 12:00", "9.00–12.00" (note the en dash), "9 til 12"
_RANGE_SEP = re.compile(r"\s*(?:-|–|—|to|til|until|indtil)\s*", re.IGNORECASE)

_TIME_RE = re.compile(
    r"^(?P<h>\d{1,2})\s*(?:[:.]\s*(?P<m>\d{2}))?\s*(?P<ampm>am|pm)?$",
    re.IGNORECASE,
)

_DURATION_RE = re.compile(
    r"^(?:(?P<h>\d+(?:[.,]\d+)?)\s*(?:h|hr|hrs|hour|hours|t|time|timer)"
    r"(?:\s*(?P<hm>\d{1,2})\s*(?:m|min|mins|minute|minutes|minutter)?)?"
    r"|(?P<m>\d+(?:[.,]\d+)?)\s*(?:m|min|mins|minute|minutes|minutter))$",
    re.IGNORECASE,
)


def _clean(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_date(value) -> dt.date:
    """Parse a cell into a :class:`datetime.date`."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return EXCEL_EPOCH + dt.timedelta(days=int(value))

    text = _clean(value)
    if not text:
        raise ValueError("empty date")

    # Drop a weekday prefix: "Mon 3 Sep", "mandag d. 3. september"
    text = re.sub(
        r"^(mon|tue|wed|thu|fri|sat|sun|man|tir|ons|tors|fre|lør|søn)[a-zæøå]*\.?,?\s+",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\bd\.\s*", "", text, flags=re.IGNORECASE)

    # ISO first: 2026-09-03 (optionally with a time part we ignore here)
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$", text)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))

    # Day-first numeric: 03/09/2026, 3-9-26, 3.9.2026
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$", text)
    if m:
        day, month, year = int(m[1]), int(m[2]), int(m[3])
        if year < 100:
            year += 2000
        if month > 12 and day <= 12:  # tolerate a US-style sheet
            day, month = month, day
        return dt.date(year, month, day)

    # Named month, either order: "3 Sep 2026" / "Sep 3, 2026" / "3. september"
    m = re.match(r"^(\d{1,2})\.?\s+([A-Za-zÆØÅæøå]+)\.?\s*(\d{4})?$", text)
    if m and m[2].lower() in MONTHS:
        year = int(m[3]) if m[3] else dt.date.today().year
        return dt.date(year, MONTHS[m[2].lower()], int(m[1]))

    m = re.match(r"^([A-Za-zÆØÅæøå]+)\.?\s+(\d{1,2})\.?,?\s*(\d{4})?$", text)
    if m and m[1].lower() in MONTHS:
        year = int(m[3]) if m[3] else dt.date.today().year
        return dt.date(year, MONTHS[m[1].lower()], int(m[2]))

    raise ValueError(f"could not read {text!r} as a date")


def parse_time(value) -> dt.time:
    """Parse a cell into a :class:`datetime.time`."""
    if isinstance(value, dt.datetime):
        return value.time()
    if isinstance(value, dt.time):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= value < 1:  # Excel time fraction of a day
            total = round(value * 24 * 60)
            return dt.time(total // 60 % 24, total % 60)
        if 0 <= value <= 24:  # a bare hour, "9"
            hour = int(value)
            minute = round((value - hour) * 60)
            return dt.time(hour % 24, minute)
        raise ValueError(f"could not read {value!r} as a time")

    text = _clean(value)
    if not text:
        raise ValueError("empty time")
    if re.fullmatch(r"\d{3,4}", text):  # military time, 0900 / 900
        return dt.time(int(text[:-2]) % 24, int(text[-2:]))

    m = _TIME_RE.match(text)
    if not m:
        raise ValueError(f"could not read {text!r} as a time")
    hour = int(m["h"])
    minute = int(m["m"] or 0)
    ampm = (m["ampm"] or "").lower()
    if ampm == "pm" and hour < 12:
        hour += 12
    elif ampm == "am" and hour == 12:
        hour = 0
    if hour == 24 and minute == 0:
        return dt.time(0, 0)
    if not (0 <= hour < 24 and 0 <= minute < 60):
        raise ValueError(f"{text!r} is not a valid clock time")
    return dt.time(hour, minute)


def split_time_range(value) -> tuple[str, str | None]:
    """Split "09:00-12:00" into its two halves; a single time keeps ``None``."""
    text = _clean(value)
    if not text:
        return "", None
    parts = _RANGE_SEP.split(text)
    if len(parts) == 2 and parts[0] and parts[1]:
        return parts[0], parts[1]
    return text, None


def parse_duration(value) -> dt.timedelta:
    """Parse "90", "1.5h", "1h30", "2 timer" into a :class:`~datetime.timedelta`."""
    if isinstance(value, dt.timedelta):
        return value
    if isinstance(value, dt.time):
        return dt.timedelta(hours=value.hour, minutes=value.minute)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # A bare number is minutes; an Excel time fraction is a part of a day.
        if 0 < value < 1:
            return dt.timedelta(days=value)
        return dt.timedelta(minutes=float(value))

    text = _clean(value).replace(",", ".")
    if not text:
        raise ValueError("empty duration")
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return dt.timedelta(minutes=float(text))

    m = re.fullmatch(r"(\d{1,2}):(\d{2})", text)
    if m:
        return dt.timedelta(hours=int(m[1]), minutes=int(m[2]))

    m = _DURATION_RE.match(text)
    if m:
        if m["h"]:
            return dt.timedelta(
                hours=float(m["h"]), minutes=int(m["hm"] or 0)
            )
        return dt.timedelta(minutes=float(m["m"]))
    raise ValueError(f"could not read {text!r} as a duration")


TRUTHY = {"yes", "y", "true", "1", "x", "ja", "sand", "heldag", "all day", "all-day"}


def parse_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return _clean(value).lower() in TRUTHY
