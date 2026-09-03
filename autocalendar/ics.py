"""Writing RFC 5545 iCalendar text.

Hand-rolled rather than pulled from a library: the format we need is small,
and a zero-dependency writer keeps the tool installable on a locked-down DTU
machine. The fiddly parts of the spec that clients actually enforce are the
ones handled here - CRLF line endings, escaping, and folding long lines at 75
octets.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Iterable

from . import __version__
from .model import Event
from .timezones import DEFAULT_TZID, vtimezone

PRODID = f"-//DTU Entrepreneurship//AutoCalendar {__version__}//EN"


def escape(text: str) -> str:
    """Escape a value per RFC 5545 section 3.3.11."""
    return (
        str(text)
        .replace("\\", "\\\\")
        .replace(";", r"\;")
        .replace(",", r"\,")
        .replace("\r\n", r"\n")
        .replace("\n", r"\n")
        .replace("\r", r"\n")
    )


def fold(line: str) -> list[str]:
    """Fold a content line to 75 octets, continuing with a leading space.

    The limit counts octets, not characters, so we fold on the UTF-8 encoding
    and are careful never to split a multi-byte character.
    """
    raw = line.encode("utf-8")
    if len(raw) <= 75:
        return [line]

    out: list[str] = []
    chunk = bytearray()
    limit = 75
    for char in line:
        encoded = char.encode("utf-8")
        if len(chunk) + len(encoded) > limit:
            out.append(chunk.decode("utf-8"))
            chunk = bytearray()
            limit = 74  # the continuation line spends one octet on its space
        chunk += encoded
    if chunk:
        out.append(chunk.decode("utf-8"))
    return [out[0]] + [" " + part for part in out[1:]]


def _stamp(moment: dt.datetime) -> str:
    return moment.strftime("%Y%m%dT%H%M%S")


def _event_lines(
    event: Event,
    tzid: str,
    now: dt.datetime,
    alarm_minutes: int | None,
    uid_domain: str,
) -> list[str]:
    lines = [
        "BEGIN:VEVENT",
        f"UID:{event.uid(uid_domain)}",
        f"DTSTAMP:{_stamp(now)}Z",
        f"SUMMARY:{escape(event.title)}",
    ]

    if event.all_day:
        # DTEND is exclusive for all-day events: a one-day event ends the next day.
        lines.append(f"DTSTART;VALUE=DATE:{event.start:%Y%m%d}")
        lines.append(f"DTEND;VALUE=DATE:{event.end + dt.timedelta(days=1):%Y%m%d}")
    else:
        lines.append(f"DTSTART;TZID={tzid}:{_stamp(event.start)}")
        lines.append(f"DTEND;TZID={tzid}:{_stamp(event.end)}")

    if event.location:
        lines.append(f"LOCATION:{escape(event.location)}")
    if event.description:
        lines.append(f"DESCRIPTION:{escape(event.description)}")
    if event.categories:
        lines.append("CATEGORIES:" + ",".join(escape(c) for c in event.categories))
    if event.url:
        lines.append(f"URL:{escape(event.url)}")
    lines.append("SEQUENCE:0")
    lines.append("STATUS:CONFIRMED")
    lines.append("TRANSP:OPAQUE")

    if alarm_minutes is not None:
        lines += [
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            f"DESCRIPTION:{escape(event.title)}",
            f"TRIGGER:-PT{int(alarm_minutes)}M",
            "END:VALARM",
        ]

    lines.append("END:VEVENT")
    return lines


def build_calendar(
    events: Iterable[Event],
    *,
    calendar_name: str = "DTU Entrepreneurship",
    tzid: str = DEFAULT_TZID,
    alarm_minutes: int | None = None,
    uid_domain: str = "autocalendar.dtu.dk",
    now: dt.datetime | None = None,
) -> str:
    """Render events as a complete iCalendar document (CRLF line endings)."""
    events = list(events)
    now = now or dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:{PRODID}",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{escape(calendar_name)}",
        f"X-WR-TIMEZONE:{tzid}",
    ]
    if any(not e.all_day for e in events):
        lines += vtimezone(tzid).splitlines()
    for event in events:
        lines += _event_lines(event, tzid, now, alarm_minutes, uid_domain)
    lines.append("END:VCALENDAR")

    folded: list[str] = []
    for line in lines:
        folded.extend(fold(line))
    return "\r\n".join(folded) + "\r\n"


def write_calendar(path, events: Iterable[Event], **kwargs) -> str:
    """Write the calendar to ``path`` and return the text that was written."""
    text = build_calendar(events, **kwargs)
    with open(path, "w", encoding="utf-8", newline="") as handle:
        handle.write(text)
    return text
