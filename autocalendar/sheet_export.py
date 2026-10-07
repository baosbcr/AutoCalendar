"""Editing events as plain fields, and writing them back as a sheet.

The web page lets a programme manager deselect and edit events before the
.ics is made. The sheet stays the source for the later invitation run, so the
same edits must be downloadable as a sheet again - in Outlook's own export
layout, which the reader takes back unchanged.
"""

from __future__ import annotations

import dataclasses
import datetime as dt
from collections.abc import Iterable

from .model import Event
from .parsing import parse_date, parse_time

# The fields the page edits, as text: dates "YYYY-MM-DD", times "HH:MM".
# An event with neither a start nor an end time is all-day, and its end date
# is its last day (inclusive).
EDITABLE = ("title", "start_date", "start_time", "end_date", "end_time", "location")

# Outlook's calendar export columns, less the housekeeping ones the reader ignores.
OUTLOOK_COLUMNS = [
    "Subject", "Start Date", "Start Time", "End Date", "End Time", "All day event",
    "Required Attendees", "Optional Attendees", "Categories", "Location", "Description",
]
# Written only when some event uses them.
OPTIONAL_COLUMNS = [("Link", "url"), ("Calendar", "calendar"), ("Tags", "tag"), ("AutoCalendar ID", "event_id")]


def to_fields(event: Event) -> dict[str, str]:
    if event.all_day:
        return {
            "title": event.title,
            "start_date": event.start.isoformat(),
            "start_time": "",
            "end_date": event.end.isoformat(),
            "end_time": "",
            "location": event.location,
            "tag": event.tag,
        }
    return {
        "title": event.title,
        "start_date": event.start.date().isoformat(),
        "start_time": f"{event.start:%H:%M}",
        "end_date": event.end.date().isoformat(),
        "end_time": f"{event.end:%H:%M}",
        "location": event.location,
        "tag": event.tag,
    }


def apply_fields(event: Event, fields: dict[str, str]) -> Event:
    """A copy of ``event`` with the edited fields; ValueError says what is wrong."""
    title = fields.get("title", "").strip()
    if not title:
        raise ValueError("the subject is empty")
    start_day = parse_date(fields.get("start_date", ""))
    end_text = fields.get("end_date", "").strip()
    end_day = parse_date(end_text) if end_text else start_day
    start_time = fields.get("start_time", "").strip()
    end_time = fields.get("end_time", "").strip()

    if not start_time and not end_time:
        if end_day < start_day:
            raise ValueError("the end date is before the start date")
        start, end = start_day, end_day
    else:
        if not start_time or not end_time:
            raise ValueError("give both a start and an end time, or neither for all day")
        start = dt.datetime.combine(start_day, parse_time(start_time))
        end = dt.datetime.combine(end_day, parse_time(end_time))
        if end <= start:
            raise ValueError("the end is not after the start")

    return dataclasses.replace(
        event, title=title, start=start, end=end, location=fields.get("location", "").strip()
    )


def write_outlook_sheet(events: Iterable[Event], target) -> None:
    """Write events as an .xlsx in Outlook's export layout (``target``: path or file).

    All-day events are written the way Outlook exports them: End Date is the day
    after the last day, End Time 00:00. The reader reads that back as the same span.
    """
    from openpyxl import Workbook

    events = list(events)
    optional = [(header, attr) for header, attr in OPTIONAL_COLUMNS
                if any(getattr(e, attr) for e in events)]
    book = Workbook()
    sheet = book.active
    sheet.title = "Calendar"
    sheet.append(OUTLOOK_COLUMNS + [header for header, _ in optional])
    for e in events:
        if e.all_day:
            first, start_t = e.start, dt.time(0, 0)
            last, end_t = e.end + dt.timedelta(days=1), dt.time(0, 0)
        else:
            first, start_t = e.start.date(), e.start.time()
            last, end_t = e.end.date(), e.end.time()
        sheet.append(
            [e.title, first, start_t, last, end_t, e.all_day,
             "; ".join(e.required), "; ".join(e.optional), ";".join(e.categories),
             e.location, e.description]
            + [getattr(e, attr) for _, attr in optional]
        )
    for row in sheet.iter_rows(min_row=2):
        row[1].number_format = row[3].number_format = "dd-mm-yyyy"
        row[2].number_format = row[4].number_format = "hh:mm"
    for column, width in zip("ABCDEFGHIJK", (70, 12, 10, 12, 10, 12, 30, 30, 14, 50, 40)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = sheet.dimensions
    book.save(target)
