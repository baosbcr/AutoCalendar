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
    "required": "Attendees",
    "optional": "Optional",
    "calendar": "Calendar",
    "event_id": "AutoCalendar ID",
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
    outlook = parser.add_argument_group(
        "Outlook",
        "Create the events in Outlook itself and send real invitations. "
        "A .ics file cannot do this - Outlook stores its attendee list for "
        "reference and sends nothing.",
    )
    outlook.add_argument(
        "--outlook",
        action="store_true",
        help="create the events in Outlook instead of writing a .ics",
    )
    outlook.add_argument(
        "--send",
        action="store_true",
        help="actually send the invitations (without this, meetings are saved as drafts)",
    )
    outlook.add_argument(
        "--test-mode",
        action="store_true",
        help="shift everything into a tagged 2099 window that --purge-tests can remove",
    )
    outlook.add_argument(
        "--keep-dates",
        action="store_true",
        help="with --test-mode: keep the real dates, still tagged so --purge-tests removes them",
    )
    outlook.add_argument(
        "--mailbox", metavar="ADDRESS", help="which mailbox to use (default: the first)"
    )
    outlook.add_argument(
        "--calendar", metavar="NAME", help="which calendar folder (default: the mailbox default)"
    )
    outlook.add_argument(
        "--limit", type=int, metavar="N", help="only process the first N events"
    )
    outlook.add_argument(
        "--list-calendars",
        action="store_true",
        help="show the calendars available in the mailbox and exit",
    )
    outlook.add_argument(
        "--sync",
        action="store_true",
        help="compare the sheet against the calendar and report what changed",
    )
    outlook.add_argument(
        "--apply",
        action="store_true",
        help="carry out the plan from --sync (add --send to notify attendees)",
    )
    outlook.add_argument(
        "--purge-tests",
        action="store_true",
        help="cancel and delete every test item from every mailbox and folder",
    )
    outlook.add_argument(
        "--yes", action="store_true", help="skip the confirmation prompt for --send / --purge-tests"
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


def _confirm(question: str, assume_yes: bool) -> bool:
    """Anything that sends mail or deletes asks first, unless told not to."""
    if assume_yes:
        return True
    try:
        return input(f"{question} [type YES to continue] ").strip() == "YES"
    except (EOFError, KeyboardInterrupt):
        print()
        return False


def _outlook_only(args) -> int:
    """--list-calendars and --purge-tests need no spreadsheet."""
    from .outlook import OutlookError, list_calendars, purge_tests

    try:
        if args.list_calendars:
            for name in list_calendars(args.mailbox):
                print(f"  {name}")
            return 0

        found = purge_tests(apply=False)["found"]
        if not found:
            print("No test items found. Nothing to purge.")
            return 0
        print(f"Found {found} test item(s) across all mailboxes and folders.")
        if not _confirm("Cancel and delete them?", args.yes):
            print("Nothing changed.")
            return 0
        counts = purge_tests(
            apply=True,
            on_progress=lambda folder, item: print(
                f"  - {getattr(item, 'Subject', '?')[:60]}"
            ),
        )
        print(
            f"\nCancelled {counts['cancelled']}, deleted {counts['deleted']} "
            f"in {counts['passes']} pass(es), {counts['remaining']} remaining."
        )
        if counts["remaining"]:
            print("Some items survived every pass - run --purge-tests again.")
            return 1
        return 0
    except OutlookError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


def _sync_with_outlook(args, path, result) -> int:
    """Second and later runs: what changed, and who would hear about it."""
    from .outlook import OutlookError, connect, find_calendar, find_store
    from .sheet_ids import SheetWriteError, ensure_ids
    from .sync import apply_plan, build_plan, describe

    try:
        added = ensure_ids(
            path, result.events, sheet=result.sheet, header_row=result.header_row
        )
    except SheetWriteError as exc:
        print(f"error: {exc}", file=sys.stderr)
        print("Nothing was changed, in the sheet or in Outlook.", file=sys.stderr)
        return 2
    if added:
        print(f"Gave {added} row(s) a permanent id and saved {path.name}.")
        print("That column is how a later run recognises these events. Leave it alone.")

    events = result.events
    if args.test_mode:
        from .outlook import shift_to_test_window

        events = shift_to_test_window(events, keep_dates=args.keep_dates)

    try:
        _, namespace = connect()
        store = find_store(namespace, args.mailbox)
        folder = find_calendar(store, args.calendar)
        plan = build_plan(events, folder)
    except OutlookError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    print()
    print(describe(plan))

    if plan.is_empty:
        print("\nNothing to do.")
        return 0
    if not args.apply:
        print("\nNothing applied. Re-run with --apply to carry this out,")
        print("and add --send if the people involved should be told.")
        return 0

    if plan.people_contacted and args.send:
        if not _confirm(
            f"This will email {plan.people_contacted} recipient(s).", args.yes
        ):
            print("Nothing changed.")
            return 0

    tally = apply_plan(plan, folder, send=args.send)
    print()
    print(
        f"  created {tally['created']}, updated {tally['updated']}, "
        f"removed {tally['cancelled']}, failed {tally['failed']}"
    )
    print(f"  mail sent to {tally['notified']} recipient(s)")
    if not args.send and plan.people_contacted:
        print("  (--send was not given, so the calendar changed but nobody was told)")
    return 1 if tally["failed"] else 0


def _push_to_outlook(args, path, result) -> int:
    from .outlook import OutlookError, push
    from .sheet_ids import SheetWriteError, ensure_ids

    events = result.events
    # Give every row an identity now, so a later --sync can recognise these
    # events after they have been edited.
    try:
        added = ensure_ids(
            path, events, sheet=result.sheet, header_row=result.header_row
        )
        if added:
            print(f"Gave {added} row(s) a permanent id and saved {path.name}.")
    except SheetWriteError as exc:
        print(f"warning: could not write ids into the sheet ({exc}).")
        print("Re-running later will not be able to match these events.")

    # count only what will actually be processed, not the whole sheet
    selected = list(events)[: args.limit] if args.limit else list(events)
    with_people = sum(1 for e in selected if e.has_attendees)
    total = len(selected)

    print()
    if args.test_mode and args.keep_dates:
        print("TEST MODE - real dates kept, everything tagged for --purge-tests.")
    elif args.test_mode:
        print("TEST MODE - everything is shifted into 2099 and tagged for --purge-tests.")
    if args.send:
        print(f"About to create {total} event(s) and SEND invitations for {with_people}.")
        if not _confirm("This emails real people.", args.yes):
            print("Nothing changed.")
            return 0
    else:
        print(f"Creating {total} event(s) as drafts. No invitations will be sent.")
        print("Add --send once the result looks right.")
    print()

    try:
        outcomes = push(
            events,
            mailbox=args.mailbox,
            calendar=args.calendar,
            send=args.send,
            test_mode=args.test_mode,
            keep_dates=args.keep_dates,
            limit=args.limit,
            reminder_minutes=args.alarm,
            on_progress=lambda outcome: print(outcome),
        )
    except OutlookError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    tally: dict[str, int] = {}
    for outcome in outcomes:
        tally[outcome.action] = tally.get(outcome.action, 0) + 1
    print("\n  " + ", ".join(f"{n} {action}" for action, n in sorted(tally.items())))

    unresolved = {name for o in outcomes for name in o.unresolved}
    if unresolved:
        print(f"\n  Could not match {len(unresolved)} name(s) in the address book:")
        for name in sorted(unresolved):
            print(f"    ? {name}")
        print("  Use a full email address for anyone outside DTU.")

    if args.test_mode:
        print("\n  Remove all of this again with:  autocalendar --purge-tests")
    return 1 if tally.get("failed") else 0


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.timezones:
        print("\n".join(supported_timezones()))
        return 0

    if args.keep_dates and not args.test_mode:
        print("error: --keep-dates only makes sense with --test-mode", file=sys.stderr)
        return 2

    if args.list_calendars or args.purge_tests:
        return _outlook_only(args)

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

    if args.sync:
        return _sync_with_outlook(args, path, result)
    if args.outlook:
        return _push_to_outlook(args, path, result)

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
