"""VTIMEZONE blocks, so wall-clock times survive the trip into a calendar.

We write local times with a TZID rather than converting to UTC, which keeps
the .ics readable and keeps daylight-saving handling in the calendar client.
That means the file has to carry a VTIMEZONE definition for the zone it uses.
The European zones below all share the CET/CEST rules, so one template covers
DTU and every partner university we are likely to hand this to.
"""

from __future__ import annotations

_EU_CENTRAL = """BEGIN:VTIMEZONE
TZID:{tzid}
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""

_EU_EASTERN = """BEGIN:VTIMEZONE
TZID:{tzid}
BEGIN:DAYLIGHT
TZOFFSETFROM:+0200
TZOFFSETTO:+0300
TZNAME:EEST
DTSTART:19700329T030000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0300
TZOFFSETTO:+0200
TZNAME:EET
DTSTART:19701025T040000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""

_EU_WESTERN = """BEGIN:VTIMEZONE
TZID:{tzid}
BEGIN:DAYLIGHT
TZOFFSETFROM:+0000
TZOFFSETTO:+0100
TZNAME:BST
DTSTART:19700329T010000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0100
TZOFFSETTO:+0000
TZNAME:GMT
DTSTART:19701025T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""

_UTC = """BEGIN:VTIMEZONE
TZID:UTC
BEGIN:STANDARD
TZOFFSETFROM:+0000
TZOFFSETTO:+0000
TZNAME:UTC
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE"""

_CENTRAL_ZONES = [
    "Europe/Copenhagen", "Europe/Amsterdam", "Europe/Berlin", "Europe/Brussels",
    "Europe/Budapest", "Europe/Madrid", "Europe/Oslo", "Europe/Paris",
    "Europe/Prague", "Europe/Rome", "Europe/Stockholm", "Europe/Vienna",
    "Europe/Warsaw", "Europe/Zurich",
]
_EASTERN_ZONES = ["Europe/Helsinki", "Europe/Riga", "Europe/Tallinn", "Europe/Vilnius"]
_WESTERN_ZONES = ["Europe/London", "Europe/Dublin", "Europe/Lisbon"]

_TEMPLATES = {tz: _EU_CENTRAL for tz in _CENTRAL_ZONES}
_TEMPLATES.update({tz: _EU_EASTERN for tz in _EASTERN_ZONES})
_TEMPLATES.update({tz: _EU_WESTERN for tz in _WESTERN_ZONES})
_TEMPLATES["UTC"] = _UTC

DEFAULT_TZID = "Europe/Copenhagen"


def supported_timezones() -> list[str]:
    return sorted(_TEMPLATES)


def vtimezone(tzid: str) -> str:
    """The VTIMEZONE block for ``tzid``, as CRLF-free lines."""
    try:
        return _TEMPLATES[tzid].format(tzid=tzid)
    except KeyError:
        raise ValueError(
            f"timezone {tzid!r} is not built in. Supported: "
            + ", ".join(supported_timezones())
        ) from None
