"""AutoCalendar - turn a spreadsheet of sessions into an .ics calendar file.

Built for DTU Entrepreneurship: take the Excel list of dates, times, rooms,
subjects and content that a programme is planned in, and produce a single
iCalendar file that can be imported ("installed") into Outlook, Google
Calendar or Apple Calendar.
"""

__version__ = "0.1.0"

from .model import Event
from .reader import read_events, ReadResult, RowProblem
from .ics import build_calendar, write_calendar

__all__ = [
    "Event",
    "read_events",
    "ReadResult",
    "RowProblem",
    "build_calendar",
    "write_calendar",
    "__version__",
]
