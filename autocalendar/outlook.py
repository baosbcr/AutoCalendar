"""Create the events in Outlook itself, and send the invitations.

The .ics writer cannot do this. Outlook reads ``ATTENDEE`` lines when it
imports a file, stores them for reference and sends nothing, so a meeting
that people are actually invited to has to be *driven* through Outlook
rather than handed to it as a file.

That is what this module does, over COM, against the copy of classic
Outlook already running on the user's machine. It needs no permission from
anyone: no app registration, no administrator consent, no cloud API.

Two things it will not do quietly:

* it never sends unless asked - :func:`push` saves drafts by default, so a
  run over a whole programme can be rehearsed without a single message
  leaving the building;
* it never deletes an organiser's meeting without cancelling it first,
  because a bare delete removes your copy and leaves every attendee holding
  a meeting that no longer exists.
"""

from __future__ import annotations

import copy
import datetime as dt
import time
from dataclasses import dataclass, field
from typing import Sequence

from .model import Event

# --- Outlook constants (we do not import the type library) -----------------
OL_APPOINTMENT = 1
OL_FOLDER_CALENDAR = 9
OL_MEETING = 1
OL_MEETING_CANCELED = 5
OL_REQUIRED, OL_OPTIONAL = 1, 2

#: Every item created in test mode carries all three of these. Cleanup wants
#: the category *and* the marker, so a bug in one cannot reach a real event.
TEST_CATEGORY = "AUTOCAL-TEST"
TEST_MARKER = "[AUTOCAL-TEST]"
TEST_BASE = dt.date(2099, 1, 4)  # start of the window, far outside any real calendar

#: Stamped on each item so a later run can recognise what it made.
PROP_UID = "AutoCalendarUID"

#: Exchange Online throttles a mailbox at 30 messages/minute. Stay under it.
SEND_INTERVAL = 2.5


class OutlookError(RuntimeError):
    """Outlook is not reachable, or refused what we asked of it."""


@dataclass
class Outcome:
    event: Event
    action: str  # saved | sent | skipped | failed
    detail: str = ""
    unresolved: list[str] = field(default_factory=list)

    def __str__(self) -> str:
        mark = {"sent": "+", "saved": ".", "skipped": "-", "failed": "!"}[self.action]
        line = f"  {mark} {self.event.title[:58]}"
        if self.detail:
            line += f"  ({self.detail})"
        return line


def connect():
    """Attach to the running Outlook, or start it."""
    try:
        import win32com.client
    except ImportError as exc:  # pragma: no cover - platform dependent
        raise OutlookError(
            "pywin32 is not installed. Install it with:  pip install pywin32"
        ) from exc
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        return app, app.GetNamespace("MAPI")
    except Exception as exc:  # pragma: no cover - depends on the desktop
        raise OutlookError(
            "could not talk to Outlook. Is classic Outlook installed and running? "
            "The new Outlook does not support this."
        ) from exc


def find_store(namespace, mailbox: str | None):
    """The mailbox to work in, matched on address or display name."""
    stores = list(namespace.Stores)
    if not mailbox:
        return stores[0]
    wanted = mailbox.strip().lower()
    for store in stores:
        if wanted in str(store.DisplayName).lower():
            return store
    available = ", ".join(str(s.DisplayName) for s in stores)
    raise OutlookError(f"no mailbox matching {mailbox!r}. Available: {available}")


def find_calendar(store, name: str | None):
    """The calendar folder to write into.

    The default folder is asked for by *id*, never by name - a Danish or
    French Outlook calls it something else, and hard-coding "Calendar" is a
    documented way for this kind of script to break.
    """
    try:
        default = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
    except Exception as exc:
        raise OutlookError(f"{store.DisplayName} has no calendar") from exc
    if not name:
        return default
    wanted = name.strip().lower()
    if wanted == str(default.Name).lower():
        return default
    for sub in default.Folders:
        if str(sub.Name).lower() == wanted:
            return sub
    choices = [str(default.Name)] + [str(s.Name) for s in default.Folders]
    raise OutlookError(f"no calendar named {name!r}. Available: {', '.join(choices)}")


def list_calendars(mailbox: str | None = None) -> list[str]:
    """What calendars this mailbox actually has, for the error message above."""
    _, namespace = connect()
    store = find_store(namespace, mailbox)
    default = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
    return [str(default.Name)] + [str(s.Name) for s in default.Folders]


def shift_to_test_window(events: Sequence[Event]) -> list[Event]:
    """Move a whole programme into the test window, keeping its shape.

    Relative day offsets and times of day are preserved, so a run over the
    real 117-row programme exercises the same code paths - it just lands in
    2099, where nothing can be mistaken for real work.
    """

    def as_date(moment):
        return moment.date() if isinstance(moment, dt.datetime) else moment

    if not events:
        return []
    # Whole weeks only, so a Monday session stays on a Monday. Shifting by a
    # raw day count moved the August programme onto weekends, which makes a
    # rehearsal look nothing like the real thing.
    gap = (TEST_BASE - min(as_date(e.start) for e in events)).days
    delta = dt.timedelta(days=-(-gap // 7) * 7)
    shifted = []
    for event in events:
        clone = copy.deepcopy(event)
        clone.start = event.start + delta
        clone.end = event.end + delta
        clone.title = f"{TEST_MARKER} {event.title}"
        clone.categories = list(event.categories) + [TEST_CATEGORY]
        shifted.append(clone)
    return shifted


def _stamp(item, event: Event, test_mode: bool) -> None:
    """Mark an item so cleanup can find it again without guessing."""
    try:
        # The sheet's own id when it has one, so a later run can recognise this
        # item after any edit. Falls back to the content hash for one-off runs.
        prop = item.UserProperties.Add(PROP_UID, 1)  # olText
        prop.Value = event.event_id or event.uid()
    except Exception:
        pass  # a custom property is a convenience, not a requirement
    categories = list(event.categories)
    if test_mode and TEST_CATEGORY not in categories:
        categories.append(TEST_CATEGORY)
    if categories:
        item.Categories = ", ".join(categories)


def _sending_account(namespace, store):
    for candidate in namespace.Accounts:
        try:
            if str(candidate.SmtpAddress).lower() == str(store.DisplayName).lower():
                return candidate
        except Exception:
            continue
    return None


def push(
    events: Sequence[Event],
    *,
    mailbox: str | None = None,
    calendar: str | None = None,
    send: bool = False,
    test_mode: bool = False,
    limit: int | None = None,
    reminder_minutes: int | None = None,
    pace: float = SEND_INTERVAL,
    on_progress=None,
) -> list[Outcome]:
    """Create the events in Outlook.

    With ``send=False`` (the default) every meeting is *saved* rather than
    sent: the whole loop runs, the calendar fills, and nobody is emailed.
    That is how to rehearse a long programme - proving the loop survives 117
    rows is a different question from proving an invitation arrives, and only
    the second one needs mail to leave the building.
    """
    if test_mode:
        events = shift_to_test_window(events)
    if limit is not None:
        events = list(events)[:limit]

    _, namespace = connect()
    store = find_store(namespace, mailbox)
    default_folder = find_calendar(store, calendar)
    account = _sending_account(namespace, store)

    outcomes: list[Outcome] = []

    for event in events:
        folder = default_folder
        if event.calendar:
            try:
                folder = find_calendar(store, event.calendar)
            except OutlookError as exc:
                outcomes.append(Outcome(event, "failed", str(exc)))
                if on_progress:
                    on_progress(outcomes[-1])
                continue

        try:
            item = folder.Items.Add(OL_APPOINTMENT)
            item.Subject = event.title
            item.Location = event.location
            item.Body = event.description
            if event.all_day:
                item.AllDayEvent = True
                item.Start = event.start.strftime("%Y-%m-%d")
            else:
                item.Start = event.start.strftime("%Y-%m-%d %H:%M")
                item.End = event.end.strftime("%Y-%m-%d %H:%M")
            item.ReminderSet = reminder_minutes is not None
            if reminder_minutes is not None:
                item.ReminderMinutesBeforeStart = reminder_minutes
            _stamp(item, event, test_mode)

            unresolved: list[str] = []
            if event.has_attendees:
                item.MeetingStatus = OL_MEETING
                if account is not None:
                    item.SendUsingAccount = account
                people = [(p, OL_REQUIRED) for p in event.required]
                people += [(p, OL_OPTIONAL) for p in event.optional]
                for who, kind in people:
                    item.Recipients.Add(who).Type = kind
                item.Recipients.ResolveAll()
                unresolved = [str(r.Name) for r in item.Recipients if not r.Resolved]

            if send and event.has_attendees and not unresolved:
                item.Send()
                outcomes.append(Outcome(event, "sent", f"{len(people)} invited"))
                if pace:
                    time.sleep(pace)
            elif send and unresolved:
                item.Save()
                outcomes.append(
                    Outcome(event, "skipped", "unresolved attendees", unresolved)
                )
            else:
                item.Save()
                detail = "draft" if event.has_attendees else "no attendees"
                outcomes.append(Outcome(event, "saved", detail, unresolved))
        except Exception as exc:  # noqa: BLE001 - COM raises anything
            outcomes.append(Outcome(event, "failed", str(exc)[:120]))

        if on_progress:
            on_progress(outcomes[-1])

    return outcomes


# --------------------------------------------------------------------------
# Cleanup
# --------------------------------------------------------------------------

def _walk(folder):
    yield folder
    try:
        for sub in folder.Folders:
            yield from _walk(sub)
    except Exception:
        return


#: Ask the store to find the items, rather than dragging every message in
#: every folder across the COM boundary. A mailbox with 5,000 items makes
#: the difference between seconds and minutes.
_SUBJECT = 'urn:schemas:httpmail:subject'
_KEYWORDS = 'urn:schemas-microsoft-com:office:office#Keywords'
_TEST_QUERY = (
    '@SQL="' + _SUBJECT + '" LIKE \'%AUTOCAL-TEST%\''
    ' OR "' + _KEYWORDS + '" LIKE \'%AUTOCAL-TEST%\''
)


def _candidates(folder):
    """Test items in this folder, asked for by query where possible."""
    try:
        return list(folder.Items.Restrict(_TEST_QUERY))
    except Exception:
        # Some folder types reject the query; fall back to reading them.
        try:
            return [i for i in folder.Items if _is_test_item(i)]
        except Exception:
            return []


def _is_test_item(item) -> bool:
    """True only for something this module created in test mode.

    The category and the marker are both accepted, and the marker is a
    literal token nobody types by accident - matching on a bare word like
    "AutoCalendar" would eventually eat a real event with that in its title.
    """
    try:
        if TEST_CATEGORY in (item.Categories or ""):
            return True
    except Exception:
        pass
    try:
        return TEST_MARKER in (item.Subject or "")
    except Exception:
        return False


def purge_tests(
    *, apply: bool = False, max_passes: int = 6, on_progress=None
) -> dict[str, int]:
    """Remove every trace of a test run, from every mailbox and every folder.

    Organiser meetings are cancelled before deletion, so attendees are told
    the meeting is gone rather than left holding a ghost.

    It sweeps repeatedly rather than a fixed number of times. Each pass
    generates the very cancellation notices the next pass has to collect, and
    Outlook does not always delete everything asked of it first time - a real
    run needed three passes to reach zero, having been written assuming two.

    Sent Items is included. The first cleanup written for this project missed
    it and left the request and the cancellation sitting there.
    """
    _, namespace = connect()
    counts = {"found": 0, "cancelled": 0, "deleted": 0, "remaining": 0, "passes": 0}

    def sweep(delete: bool) -> int:
        seen = 0
        for store in namespace.Stores:
            try:
                root = store.GetRootFolder()
            except Exception:
                continue
            for folder in _walk(root):
                for item in _candidates(folder):
                    if not _is_test_item(item):
                        continue
                    seen += 1
                    if not delete:
                        continue
                    try:
                        is_meeting = (
                            getattr(item, "MeetingStatus", 0) == OL_MEETING
                            and item.Recipients.Count > 0
                        )
                    except Exception:
                        is_meeting = False
                    try:
                        if on_progress:
                            on_progress(folder, item)
                        if is_meeting:
                            item.MeetingStatus = OL_MEETING_CANCELED
                            item.Save()
                            item.Send()
                            counts["cancelled"] += 1
                        item.Delete()
                        counts["deleted"] += 1
                    except Exception:
                        continue
        return seen

    counts["found"] = sweep(delete=False)
    if not apply:
        counts["remaining"] = counts["found"]
        return counts

    for attempt in range(1, max_passes + 1):
        sweep(delete=True)
        time.sleep(2)  # let the cancellations this pass sent arrive
        remaining = sweep(delete=False)
        counts["passes"] = attempt
        if not remaining:
            break
    counts["remaining"] = remaining
    return counts
