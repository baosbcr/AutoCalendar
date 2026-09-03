"""Command line: spreadsheet in, .ics out.

Run with no arguments and it opens a file picker, so it can be handed to
someone who would rather not meet a terminal.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

from . import __version__
from .ics import write_calendar
from .parsing import parse_duration
from .reader import read_events
from .timezones import DEFAULT_TZID, supported_timezones

FIELD_LABELS = {
    "title": "Subject",
    "date": "Date",
    "end_date": "End date",
    "start": "Start time",
    "end": "End time",
    "duration": "Duration",
    "location": "Location",
    "description": "Description",
    "category": "Category",
    "url": "Link",
    "all_day": "All day",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="autocalendar",
        description=(
            "Convert a spreadsheet of sessions (dates, times, rooms, subjects, "
            "content) into an .ics calendar file you can import into Outlook, "
            "Google Calendar or Apple Calendar."
        ),
        epilog="Run without arguments to pick the file in a dialog instead.",
    )
    parser.add_argument(
        "input", nargs="?", help="the .xlsx or .csv planning sheet to convert"
    )
    parser.add_argument(
        "-o", "--output", help="where to write the .ics (default: next to the input)"
    )
    parser.add_argument(
        "-s", "--sheet", help="worksheet name or 0-based index (default: the first one)"
    )
    parser.add_argument(
        "-n", "--name", help="calendar name shown by the calendar app"
    )
    parser.add_argument(
        "--tz",
        default=DEFAULT_TZID,
        metavar="ZONE",
        help=f"timezone of the times in the sheet (default: {DEFAULT_TZID})",
    )
    parser.add_argument(
        "--duration",
        default="60",
        metavar="LENGTH",
        help='how long a session lasts when the sheet does not say, e.g. "90" or "1h30" (default: 60)',
    )
    parser.add_argument(
        "--alarm",
        type=int,
        metavar="MINUTES",
        help="add a reminder this many minutes before each session",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        dest="list_only",
        help="show what was read and write nothing (a dry run)",
    )
    parser.add_argument(
        "--no-extra-columns",
        action="store_true",
        help="do not append unrecognised columns to the event description",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="exit with an error if any row could not be converted",
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="only report problems")
    parser.add_argument(
        "--timezones", action="store_true", help="list the supported timezones and exit"
    )
    parser.add_argument("--version", action="version", version=f"autocalendar {__version__}")
    return parser


def pick_file() -> str | None:
    """Ask for the input file with a dialog (stdlib tkinter)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:  # pragma: no cover - tkinter missing on some builds
        return None
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select the planning spreadsheet",
        filetypes=[("Spreadsheets", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")],
    )
    root.destroy()
    return path or None


def _as_date(moment):
    """All-day events carry dates, timed ones datetimes; compare them as dates."""
    return moment.date() if isinstance(moment, dt.datetime) else moment


def _sheet_argument(value: str | None) -> str | int | None:
    if value is None:
        return None
    return int(value) if value.isdigit() else value


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.timezones:
        print("\n".join(supported_timezones()))
        return 0

    source = args.input or pick_file()
    if not source:
        print("No input file given. Try: autocalendar schedule.xlsx", file=sys.stderr)
        return 2

    path = Path(source)
    if not path.exists():
        print(f"error: {path} does not exist", file=sys.stderr)
        return 2

    try:
        default_duration = parse_duration(args.duration)
    except ValueError as exc:
        print(f"error: --duration {exc}", file=sys.stderr)
        return 2

    try:
        result = read_events(
            path,
            sheet=_sheet_argument(args.sheet),
            default_duration=default_duration,
            keep_extra_columns=not args.no_extra_columns,
        )
    except (ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if not args.quiet:
        print(f"Read {path.name}" + (f" [{result.sheet}]" if result.sheet else ""))
        if result.mapping:
            print(f"  Header row {result.header_row}. Columns used:")
            for field, header in result.mapping.items():
                print(f"    {FIELD_LABELS.get(field, field):<12} <- {header}")
        if result.extra_columns:
            where = (
                "ignored"
                if args.no_extra_columns
                else "added to the event description"
            )
            print(f"  Other columns ({where}): {', '.join(result.extra_columns)}")
        print(f"  {len(result.events)} event(s) found")

    for problem in result.problems:
        stream = sys.stderr if problem.fatal else sys.stdout
        if problem.fatal or not args.quiet:
            print(f"  ! {problem}", file=stream)

    if not result.events:
        print("error: no events could be read from the file", file=sys.stderr)
        return 1

    if args.list_only:
        print()
        for event in result.events:
            print(f"  {event}")
        return 0

    output = Path(args.output) if args.output else path.with_suffix(".ics")
    calendar_name = args.name or path.stem.replace("_", " ").strip()

    try:
        write_calendar(
            output,
            result.events,
            calendar_name=calendar_name,
            tzid=args.tz,
            alarm_minutes=args.alarm,
        )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if not args.quiet:
        first = min(_as_date(e.start) for e in result.events)
        last = max(_as_date(e.end) for e in result.events)
        print(f"  {first:%d %b %Y} - {last:%d %b %Y}, timezone {args.tz}")
        print(f"\nWrote {output}")
        print("Import it: Outlook > File > Open & Export > Import, or")
        print("           Google Calendar > Settings > Import & export.")

    if args.strict and result.skipped:
        return 1
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
