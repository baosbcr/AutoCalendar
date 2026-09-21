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
    # Exact name first: "Public Folders - you@dtu.dk" also *contains* your address,
    # and it is often listed before the mailbox itself.
    for store in stores:
        if str(store.DisplayName).strip().lower() == wanted:
            return store
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


def shift_to_test_window(events: Sequence[Event], keep_dates: bool = False) -> list[Event]:
    """Move a whole programme into the test window, keeping its shape.

    Relative day offsets and times of day are preserved, so a run over the
    real 117-row programme exercises the same code paths - it just lands in
    2099, where nothing can be mistaken for real work.

    With ``keep_dates`` the events stay on their real dates but are still
    tagged, so --purge-tests can remove them. For rehearsing with the real
    dates when every attendee is in on the test.
    """

    def as_date(moment):
        return moment.date() if isinstance(moment, dt.datetime) else moment

    if not events:
        return []
    # Whole weeks only, so a Monday session stays on a Monday. Shifting by a
    # raw day count moved the August programme onto weekends, which makes a
    # rehearsal look nothing like the real thing.
    gap = (TEST_BASE - min(as_date(e.start) for e in events)).days
    delta = dt.timedelta(0) if keep_dates else dt.timedelta(days=-(-gap // 7) * 7)
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
    keep_dates: bool = False,
    limit: int | None = None,
    reminder_minutes: int | None = None,
    response_requested: bool = True,
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
        events = shift_to_test_window(events, keep_dates=keep_dates)
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
                # The sheet's end date is inclusive; Outlook's End is the midnight after it.
                # Without this a multi-day event collapses to its first day.
                item.End = (event.end + dt.timedelta(days=1)).strftime("%Y-%m-%d")
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
                # Off for placeholders: attendees see no "Please respond" and
                # the organiser gets no flood of replies.
                item.ResponseRequested = response_requested
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


OL_FOLDER_OUTBOX = 4


def _stores(namespace, mailbox: str | None = None):
    """Every store, or only the one whose name or address matches ``mailbox``."""
    wanted = mailbox.strip().lower() if mailbox else None
    for store in namespace.Stores:
        try:
            name = str(store.DisplayName).strip().lower()
        except Exception:
            name = ""
        if wanted is None or name == wanted:
            yield store


def _outbox_ids(namespace, mailbox: str | None = None) -> set[str]:
    ids = set()
    for store in _stores(namespace, mailbox):
        try:
            ids.add(store.GetDefaultFolder(OL_FOLDER_OUTBOX).EntryID)
        except Exception:
            continue
    return ids


def _test_items(
    namespace, outboxes: set[str], *, include_outbox: bool = False, mailbox: str | None = None
):
    """(folder, item) for every test item outside the Outbox - or, with
    ``include_outbox``, only those inside it. What sits in the Outbox is mail not
    yet sent; deleting it silently un-sends it."""
    for store in _stores(namespace, mailbox):
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        for folder in _walk(root):
            try:
                in_outbox = folder.EntryID in outboxes
            except Exception:
                in_outbox = False
            if in_outbox != include_outbox:
                continue
            for item in _candidates(folder):
                if _is_test_item(item):
                    yield folder, item


MSG_CANCELED = "IPM.Schedule.Meeting.Canceled"
#: Deleted Items and the Sync Issues folders (Conflicts, Local Failures, Server
#: Failures). An organiser meeting there is already gone from the calendar - or
#: a conflict copy - so deleting it cannot leave an attendee holding anything.
_SAFE_TO_DELETE = (3, 19, 20, 21, 22)


def _safe_folder_ids(namespace, mailbox: str | None = None) -> set[str]:
    ids = set()
    for store in _stores(namespace, mailbox):
        for kind in _SAFE_TO_DELETE:
            try:
                ids.add(str(store.GetDefaultFolder(kind).EntryID))
            except Exception:
                continue
    return ids
OL_FOLDER_SENT = 5


def _outgoing_folders(namespace, mailbox: str | None = None) -> list:
    """Outbox and Sent Items of each store: where a sent message shows up."""
    folders = []
    for store in _stores(namespace, mailbox):
        for kind in (OL_FOLDER_OUTBOX, OL_FOLDER_SENT):
            try:
                folders.append(store.GetDefaultFolder(kind))
            except Exception:
                continue
    return folders


def _outgoing_ids(folders, subject: str) -> set[str]:
    ids = set()
    for folder in folders:
        for message in _candidates(folder):
            try:
                if subject in str(message.Subject):
                    ids.add(str(message.EntryID))
            except Exception:
                continue
    return ids


def _new_outgoing_kind(folders, subject: str, before: set[str], wait: float) -> str | None:
    """MessageClass of the message a Send() just produced, or None if none appeared."""
    deadline = time.monotonic() + wait
    while True:
        for folder in folders:
            for message in _candidates(folder):
                try:
                    if subject in str(message.Subject) and str(message.EntryID) not in before:
                        return str(message.MessageClass)
                except Exception:
                    continue
        if time.monotonic() >= deadline:
            return None
        time.sleep(0.5)


def _is_live_organiser_meeting(item) -> bool:
    try:
        return (
            getattr(item, "MeetingStatus", 0) == OL_MEETING
            and item.Recipients.Count > 0
        )
    except Exception:
        return False


def _meeting_key(item) -> str:
    for attr in ("GlobalAppointmentID", "EntryID"):
        try:
            value = getattr(item, attr)
            if value:
                return str(value)
        except Exception:
            continue
    return str(id(item))


def purge_tests(
    *,
    apply: bool = False,
    max_passes: int = 6,
    on_progress=None,
    namespace=None,
    mailbox: str | None = None,
    pace: float = SEND_INTERVAL,
    outbox_timeout: float = 600,
    poll: float = 5,
    verify_timeout: float = 30,
) -> dict[str, int]:
    """Remove every trace of a test run, from every mailbox and every folder.

    Three phases, in this order, because each one depends on the last:

    1. **Cancel** every organiser meeting, once each, paced like sending, so
       attendees are told the meeting is gone rather than left holding a ghost.
    2. **Wait** until the Outbox has actually sent those cancellations. Exchange
       throttles at ~30 messages a minute, so they queue.
    3. **Delete** every test item, in every folder except the Outbox, in
       repeated passes (Outlook does not always delete everything first time).

    The first version did all three in one sweep and swept the Outbox too: pass
    two found the queued "Canceled: [AUTOCAL-TEST] ..." messages by their tag and
    deleted them before they left, so most attendees were never told (2026-09-21:
    only the first handful of 73 cancellations reached the attendee).

    If the Outbox does not drain within ``outbox_timeout`` seconds nothing is
    deleted - better a leftover test item than an attendee with a ghost meeting.

    ``mailbox`` limits all of this to one mailbox. Without it every mailbox in
    the profile is swept, which removes the attendee side too when a test
    recipient's mailbox is also open in this Outlook - but also removes their
    cancellations before anyone has checked they arrived.

    Sent Items is included in phase 3. The first cleanup written for this project
    missed it and left the request and the cancellation sitting there.
    """
    if namespace is None:
        _, namespace = connect()
    outboxes = _outbox_ids(namespace, mailbox)
    outgoing = _outgoing_folders(namespace, mailbox)
    safe_folders = _safe_folder_ids(namespace, mailbox)
    counts = {"found": 0, "cancelled": 0, "cancel_failed": 0, "resent": 0, "deleted": 0,
              "remaining": 0, "passes": 0, "outbox_stuck": 0}

    def count() -> int:
        return sum(1 for _ in _test_items(namespace, outboxes, mailbox=mailbox))

    counts["found"] = count()
    if not apply:
        counts["remaining"] = counts["found"]
        return counts

    # 1. cancel, once per meeting. Collect ids first and open one item at a time:
    # holding every item open at once is what went wrong on 2026-09-21 - after
    # 33 meetings Outlook silently stopped applying the "cancelled" status, and
    # Send() then re-sent 40 of them as plain invitations.
    targets = []
    for folder, item in _test_items(namespace, outboxes, mailbox=mailbox):
        if _is_live_organiser_meeting(item):
            try:
                targets.append((str(item.EntryID), str(folder.StoreID)))
            except Exception:
                continue
        del item
    cancelled: set[str] = set()
    for entry_id, store_id in targets:
        try:
            item = namespace.GetItemFromID(entry_id, store_id)
            key = _meeting_key(item)
            if key in cancelled or not _is_live_organiser_meeting(item):
                continue
            if on_progress:
                on_progress(None, item)
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Save()
            # Send() on a meeting that is not really cancelled re-sends the
            # invitation - the one thing a purge must never do. Check the open
            # item first (Outlook commits the status only on Send; reopening the
            # item never shows it, not even afterwards).
            if getattr(item, "MeetingStatus", None) != OL_MEETING_CANCELED:
                counts["cancel_failed"] += 1
                continue
            subject = str(item.Subject)
            before = _outgoing_ids(outgoing, subject)
            item.Send()
            # The only reliable proof is the message that went out.
            kind = _new_outgoing_kind(outgoing, subject, before, wait=verify_timeout)
            if kind != MSG_CANCELED:
                counts["resent"] += 1  # an invitation, or nothing we can see: stop
                break
            # Remove the organiser copy now, as Outlook's own "Send Cancellation"
            # does, so a later run cannot cancel it a second time.
            item.Delete()
            del item
            cancelled.add(key)
            counts["cancelled"] += 1
            counts["deleted"] += 1
            if pace:
                time.sleep(pace)
        except Exception:
            counts["cancel_failed"] += 1
            continue

    if counts["resent"]:
        counts["remaining"] = count()
        return counts

    # 2. wait for the Outbox to empty of test mail
    def queued() -> int:
        return sum(1 for _ in _test_items(namespace, outboxes, include_outbox=True, mailbox=mailbox))

    waited = 0.0
    while queued():
        if waited >= outbox_timeout:
            counts["outbox_stuck"] = queued()
            counts["remaining"] = count()
            return counts
        try:
            namespace.SendAndReceive(False)
        except Exception:
            pass
        time.sleep(poll)
        waited += poll

    # 3. delete, never touching the Outbox, and never a live organiser meeting:
    # deleting one of those unsent-cancelled leaves attendees holding a ghost
    # that nobody can cancel any more. It stays behind for the next run.
    remaining = counts["found"]
    for attempt in range(1, max_passes + 1):
        ids = []
        for folder, item in _test_items(namespace, outboxes, mailbox=mailbox):
            try:
                protected = (
                    _is_live_organiser_meeting(item)
                    and str(folder.EntryID) not in safe_folders
                    and _meeting_key(item) not in cancelled
                )
                if not protected:
                    ids.append((str(item.EntryID), str(folder.StoreID)))
            except Exception:
                pass
            del item
        for entry_id, store_id in ids:
            try:
                item = namespace.GetItemFromID(entry_id, store_id)
                if on_progress:
                    on_progress(None, item)
                item.Delete()
                del item
                counts["deleted"] += 1
            except Exception:
                continue
        time.sleep(poll if pace else 0)
        remaining = count()
        counts["passes"] = attempt
        if not remaining:
            break
    counts["remaining"] = remaining
    return counts
