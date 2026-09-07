"""The one thing a calendar entry is, once the spreadsheet has been read."""

from __future__ import annotations

import datetime as dt
import hashlib
from dataclasses import dataclass, field


@dataclass
class Event:
    """A single calendar entry.

    All-day events carry ``start``/``end`` as :class:`datetime.date`; timed
    events carry :class:`datetime.datetime` in *local* wall-clock time of the
    calendar's timezone (no tzinfo attached - the TZID in the .ics file says
    which zone the wall clock belongs to).
    """

    title: str
    start: dt.datetime | dt.date
    end: dt.datetime | dt.date
    location: str = ""
    description: str = ""
    categories: list[str] = field(default_factory=list)
    url: str = ""
    required: list[str] = field(default_factory=list)
    optional: list[str] = field(default_factory=list)
    calendar: str = ""
    source_row: int = 0

    @property
    def has_attendees(self) -> bool:
        return bool(self.required or self.optional)

    @property
    def all_day(self) -> bool:
        return not isinstance(self.start, dt.datetime)

    def uid(self, domain: str = "autocalendar.dtu.dk") -> str:
        """A stable UID.

        Derived from the content, so re-exporting an unchanged spreadsheet
        produces the same UIDs and calendars update the existing entries
        instead of duplicating them.
        """
        seed = "|".join(
            [
                self.title,
                self.start.isoformat(),
                self.end.isoformat(),
                self.location,
                self.description,
            ]
        )
        digest = hashlib.sha1(seed.encode("utf-8")).hexdigest()[:32]
        return f"{digest}@{domain}"

    def __str__(self) -> str:  # pragma: no cover - convenience for --list
        if self.all_day:
            when = f"{self.start:%a %d %b %Y} (all day)"
        else:
            when = f"{self.start:%a %d %b %Y  %H:%M}-{self.end:%H:%M}"
        where = f"  @ {self.location}" if self.location else ""
        return f"{when}  {self.title}{where}"
