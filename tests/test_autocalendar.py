"""Tests: python -m unittest discover -s tests

Stdlib only, so they run anywhere the tool itself runs.
"""

from __future__ import annotations

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from autocalendar.ics import build_calendar, escape, fold  # noqa: E402
from autocalendar.model import Event  # noqa: E402
from autocalendar.parsing import (  # noqa: E402
    parse_date,
    parse_duration,
    parse_time,
    split_time_range,
)
from autocalendar.outlook import (  # noqa: E402
    TEST_CATEGORY,
    TEST_MARKER,
    shift_to_test_window,
)
from autocalendar.sync import (  # noqa: E402
    NEW,
    ORPHANED,
    UNCHANGED,
    UPDATED,
    Change,
    Plan,
    compare,
)
from autocalendar.reader import (  # noqa: E402
    find_header_row,
    normalise,
    read_events,
    split_people,
)


def write_csv(text: str) -> Path:
    handle = tempfile.NamedTemporaryFile(
        "w", suffix=".csv", delete=False, encoding="utf-8", newline=""
    )
    handle.write(text)
    handle.close()
    return Path(handle.name)


class TestParseDate(unittest.TestCase):
    def test_formats_all_land_on_the_same_day(self):
        expected = dt.date(2026, 9, 3)
        for text in [
            "2026-09-03",
            "03/09/2026",
            "3-9-2026",
            "3.9.26",
            "3 Sep 2026",
            "3. september 2026",
            "Sep 3, 2026",
            "Thu 3 Sep 2026",
            "torsdag d. 3. september 2026",
        ]:
            with self.subTest(text=text):
                self.assertEqual(parse_date(text), expected)

    def test_native_and_excel_serial_values(self):
        self.assertEqual(parse_date(dt.datetime(2026, 9, 3, 10, 0)), dt.date(2026, 9, 3))
        self.assertEqual(parse_date(dt.date(2026, 9, 3)), dt.date(2026, 9, 3))
        self.assertEqual(parse_date(46268), dt.date(2026, 9, 3))

    def test_us_style_month_first_is_tolerated(self):
        self.assertEqual(parse_date("09/25/2026"), dt.date(2026, 9, 25))

    def test_unreadable_date_raises(self):
        for text in ["TBD", "", "week 38", None]:
            with self.subTest(text=text), self.assertRaises(ValueError):
                parse_date(text)


class TestParseTime(unittest.TestCase):
    def test_formats(self):
        cases = {
            "9": dt.time(9, 0),
            "09:00": dt.time(9, 0),
            "9.30": dt.time(9, 30),
            "0900": dt.time(9, 0),
            "930": dt.time(9, 30),
            "13:45": dt.time(13, 45),
            "1:30 pm": dt.time(13, 30),
            "12:00 am": dt.time(0, 0),
            dt.time(8, 15): dt.time(8, 15),
            dt.datetime(2026, 1, 1, 8, 15): dt.time(8, 15),
            0.5: dt.time(12, 0),  # Excel time fraction
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(parse_time(value), expected)

    def test_invalid_time_raises(self):
        for value in ["25:00", "noon", ""]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_time(value)


class TestRangesAndDurations(unittest.TestCase):
    def test_split_time_range(self):
        self.assertEqual(split_time_range("09:00-12:00"), ("09:00", "12:00"))
        self.assertEqual(split_time_range("9 – 12"), ("9", "12"))
        self.assertEqual(split_time_range("13:00 til 16:00"), ("13:00", "16:00"))
        self.assertEqual(split_time_range("09:00"), ("09:00", None))
        self.assertEqual(split_time_range(None), ("", None))

    def test_durations(self):
        cases = {
            "90": dt.timedelta(minutes=90),
            90: dt.timedelta(minutes=90),
            "1.5h": dt.timedelta(minutes=90),
            "1h30": dt.timedelta(minutes=90),
            "1:30": dt.timedelta(minutes=90),
            "2 timer": dt.timedelta(hours=2),
            "45 min": dt.timedelta(minutes=45),
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(parse_duration(value), expected)


class TestIcsFormatting(unittest.TestCase):
    def test_escape(self):
        self.assertEqual(escape("a, b; c\\d\ne"), r"a\, b\; c\\d\ne")

    def test_fold_keeps_lines_within_75_octets(self):
        line = "DESCRIPTION:" + "æ" * 200  # two octets per character
        folded = fold(line)
        self.assertGreater(len(folded), 1)
        for part in folded:
            self.assertLessEqual(len(part.encode("utf-8")), 75)
        for part in folded[1:]:
            self.assertTrue(part.startswith(" "))
        unfolded = folded[0] + "".join(p[1:] for p in folded[1:])
        self.assertEqual(unfolded, line)

    def test_short_line_is_untouched(self):
        self.assertEqual(fold("SUMMARY:Kick-off"), ["SUMMARY:Kick-off"])


class TestCalendar(unittest.TestCase):
    def setUp(self):
        self.timed = Event(
            title="Pitch training",
            start=dt.datetime(2026, 9, 14, 16, 0),
            end=dt.datetime(2026, 9, 14, 19, 0),
            location="Skylab",
        )
        self.all_day = Event(
            title="Supervision",
            start=dt.date(2026, 9, 10),
            end=dt.date(2026, 9, 10),
        )

    def test_structure_and_crlf(self):
        text = build_calendar([self.timed])
        self.assertTrue(text.startswith("BEGIN:VCALENDAR\r\n"))
        self.assertTrue(text.endswith("END:VCALENDAR\r\n"))
        self.assertIn("BEGIN:VTIMEZONE\r\nTZID:Europe/Copenhagen", text)
        self.assertIn("DTSTART;TZID=Europe/Copenhagen:20260914T160000", text)
        self.assertIn("DTEND;TZID=Europe/Copenhagen:20260914T190000", text)
        self.assertEqual(text.count("BEGIN:VEVENT"), 1)

    def test_all_day_end_is_exclusive(self):
        text = build_calendar([self.all_day])
        self.assertIn("DTSTART;VALUE=DATE:20260910", text)
        self.assertIn("DTEND;VALUE=DATE:20260911", text)
        self.assertNotIn("VTIMEZONE", text)  # nothing timed, so none is needed

    def test_alarm_is_optional(self):
        self.assertNotIn("VALARM", build_calendar([self.timed]))
        self.assertIn("TRIGGER:-PT15M", build_calendar([self.timed], alarm_minutes=15))

    def test_uid_is_stable_for_the_same_event(self):
        again = Event(
            title="Pitch training",
            start=dt.datetime(2026, 9, 14, 16, 0),
            end=dt.datetime(2026, 9, 14, 19, 0),
            location="Skylab",
        )
        self.assertEqual(self.timed.uid(), again.uid())
        self.timed.location = "Building 101"
        self.assertNotEqual(self.timed.uid(), again.uid())

    def test_unknown_timezone_is_refused(self):
        with self.assertRaises(ValueError):
            build_calendar([self.timed], tzid="Mars/Olympus")


class TestReader(unittest.TestCase):
    def test_reads_a_sheet_with_a_title_row_and_danish_headers(self):
        path = write_csv(
            "Efterårsprogram 2026\n"
            "\n"
            "Dato,Tid,Emne,Lokale,Indhold,Underviser\n"
            "07/09/2026,09:00-12:00,Kick-off,358 Aud 060,Velkomst,Toke\n"
            "08/09/2026,13:00,Idégenerering,Skylab,Effectuation,Toke\n"
        )
        try:
            result = read_events(path)
        finally:
            path.unlink()

        self.assertEqual(result.header_row, 3)
        self.assertEqual(len(result.events), 2)
        self.assertEqual(result.mapping["title"], "Emne")
        first, second = result.events
        self.assertEqual(first.start, dt.datetime(2026, 9, 7, 9, 0))
        self.assertEqual(first.end, dt.datetime(2026, 9, 7, 12, 0))
        self.assertEqual(first.location, "358 Aud 060")
        # No end time and no duration, so the default hour applies.
        self.assertEqual(second.end - second.start, dt.timedelta(hours=1))
        # The unrecognised column is kept, labelled with its own header.
        self.assertIn("Underviser: Toke", first.description)

    def test_missing_time_columns_make_an_all_day_event(self):
        path = write_csv("Date,Subject\n2026-09-10,Supervision\n")
        try:
            result = read_events(path)
        finally:
            path.unlink()
        event = result.events[0]
        self.assertTrue(event.all_day)
        self.assertEqual(event.start, dt.date(2026, 9, 10))

    def test_session_running_past_midnight(self):
        path = write_csv("Date,Start,End,Subject\n2026-09-10,22:00,01:00,Hackathon\n")
        try:
            result = read_events(path)
        finally:
            path.unlink()
        event = result.events[0]
        self.assertEqual(event.end, dt.datetime(2026, 9, 11, 1, 0))
        self.assertTrue(result.warnings)

    def test_multi_day_all_day_event(self):
        path = write_csv(
            "Date,End date,Subject\n2026-09-10,2026-09-12,Study trip\n"
        )
        try:
            result = read_events(path)
        finally:
            path.unlink()
        event = result.events[0]
        self.assertEqual(event.start, dt.date(2026, 9, 10))
        self.assertEqual(event.end, dt.date(2026, 9, 12))

    def test_bad_rows_are_reported_not_dropped_silently(self):
        path = write_csv(
            "Date,Start,Subject\n"
            "TBD,09:00,Not scheduled yet\n"
            "2026-09-10,09:00,\n"
            "2026-09-11,09:00,Real session\n"
        )
        try:
            result = read_events(path)
        finally:
            path.unlink()
        self.assertEqual([e.title for e in result.events], ["Real session"])
        self.assertEqual(len(result.skipped), 2)
        self.assertIn("TBD", str(result.skipped[0]))

    def test_missing_required_columns_is_reported(self):
        path = write_csv("Room,Notes\nSkylab,nothing here\n")
        try:
            result = read_events(path)
        finally:
            path.unlink()
        self.assertEqual(result.events, [])
        self.assertIn("could not find", result.problems[0].message)

    def test_semicolon_separated_export(self):
        path = write_csv("Date;Start;Subject\n2026-09-10;09:00;Semikolon\n")
        try:
            result = read_events(path)
        finally:
            path.unlink()
        self.assertEqual([e.title for e in result.events], ["Semikolon"])

    def test_extra_columns_can_be_dropped(self):
        path = write_csv("Date,Subject,Owner\n2026-09-10,Session,Toke\n")
        try:
            result = read_events(path, keep_extra_columns=False)
        finally:
            path.unlink()
        self.assertEqual(result.events[0].description, "")


class TestHeaderDetection(unittest.TestCase):
    def test_normalise(self):
        self.assertEqual(normalise(" Start Time "), "starttime")
        self.assertEqual(normalise("Varighed (min)"), "varighedmin")

    def test_header_row_is_found_below_a_title_block(self):
        rows = [["Programme 2026"], [], ["Date", "Subject", "Room"], ["2026-09-10"]]
        self.assertEqual(find_header_row(rows), 2)


class TestExampleSheet(unittest.TestCase):
    """The shipped example should keep converting cleanly."""

    def test_example_csv_round_trip(self):
        example = Path(__file__).parent.parent / "examples" / "programme_autumn2026.csv"
        if not example.exists():  # generated on demand
            self.skipTest("run examples/make_example_sheet.py first")
        result = read_events(example)
        self.assertEqual(len(result.events), 7)
        self.assertEqual(len(result.skipped), 1)  # the "TBD" row
        text = build_calendar(result.events, calendar_name="Example")
        self.assertEqual(text.count("BEGIN:VEVENT"), 7)


class TestAttendeeColumns(unittest.TestCase):
    def test_semicolons_win_over_commas(self):
        # a display name may legitimately contain a comma
        self.assertEqual(
            split_people("Nielsen, Jan; Hansen, Ole"), ["Nielsen, Jan", "Hansen, Ole"]
        )

    def test_commas_split_when_there_is_no_semicolon(self):
        self.assertEqual(
            split_people("a@dtu.dk, b@dtu.dk"), ["a@dtu.dk", "b@dtu.dk"]
        )

    def test_blank_gives_nothing(self):
        self.assertEqual(split_people(""), [])
        self.assertEqual(split_people("   "), [])

    def test_reads_attendees_from_a_sheet(self):
        path = write_csv(
            "Date,Time,Subject,Attendees,Optional,Calendar\n"
            "2026-09-07,09:00,Kick-off,a@dtu.dk; b@dtu.dk,c@dtu.dk,Teaching\n"
        )
        event = read_events(path).events[0]
        self.assertEqual(event.required, ["a@dtu.dk", "b@dtu.dk"])
        self.assertEqual(event.optional, ["c@dtu.dk"])
        self.assertEqual(event.calendar, "Teaching")
        self.assertTrue(event.has_attendees)

    def test_a_speaker_column_is_not_an_invitation_list(self):
        """A column of names must never become a list of people to email."""
        path = write_csv(
            "Date,Time,Subject,Speaker,Underviser\n"
            "2026-09-07,09:00,Lecture,Jane Doe,Toke\n"
        )
        event = read_events(path).events[0]
        self.assertEqual(event.required, [])
        self.assertFalse(event.has_attendees)
        self.assertIn("Speaker: Jane Doe", event.description)


class TestTestWindow(unittest.TestCase):
    def setUp(self):
        self.events = [
            Event(
                title="First",
                start=dt.datetime(2026, 8, 3, 9, 0),
                end=dt.datetime(2026, 8, 3, 10, 0),
            ),
            Event(
                title="Third day",
                start=dt.datetime(2026, 8, 5, 14, 30),
                end=dt.datetime(2026, 8, 5, 15, 0),
            ),
        ]

    def test_shape_is_preserved(self):
        shifted = shift_to_test_window(self.events)
        gap = shifted[1].start - shifted[0].start
        self.assertEqual(gap, dt.timedelta(days=2, hours=5, minutes=30))

    def test_lands_far_from_any_real_calendar(self):
        for event in shift_to_test_window(self.events):
            self.assertEqual(event.start.year, 2099)

    def test_everything_is_tagged_both_ways(self):
        for event in shift_to_test_window(self.events):
            self.assertIn(TEST_CATEGORY, event.categories)
            self.assertTrue(event.title.startswith(TEST_MARKER))

    def test_weekdays_are_preserved(self):
        """A Monday session must not land on a Sunday."""
        for original, shifted in zip(self.events, shift_to_test_window(self.events)):
            self.assertEqual(original.start.weekday(), shifted.start.weekday())

    def test_times_of_day_are_preserved(self):
        for original, shifted in zip(self.events, shift_to_test_window(self.events)):
            self.assertEqual(original.start.time(), shifted.start.time())

    def test_the_originals_are_untouched(self):
        shift_to_test_window(self.events)
        self.assertEqual(self.events[0].start.year, 2026)
        self.assertEqual(self.events[0].title, "First")

    def test_empty_input(self):
        self.assertEqual(shift_to_test_window([]), [])

    def test_keep_dates_stays_put_but_is_still_purgeable(self):
        """Rehearsing on the real dates must still leave --purge-tests a way in."""
        kept = shift_to_test_window(self.events, keep_dates=True)
        for original, event in zip(self.events, kept):
            self.assertEqual(event.start, original.start)
            self.assertEqual(event.end, original.end)
            self.assertIn(TEST_CATEGORY, event.categories)
            self.assertTrue(event.title.startswith(TEST_MARKER))


class FakeRecipients:
    def __init__(self, addresses):
        self._people = [type("R", (), {"Address": a, "Name": a})() for a in addresses]

    def __iter__(self):
        return iter(self._people)

    @property
    def Count(self):
        return len(self._people)


class FakeItem:
    """Enough of an Outlook AppointmentItem for compare() to work on."""

    def __init__(self, **kw):
        self.Subject = kw.get("subject", "Kick-off")
        self.Location = kw.get("location", "Room 1")
        self.Start = kw.get("start", dt.datetime(2026, 8, 3, 9, 0))
        self.End = kw.get("end", dt.datetime(2026, 8, 3, 10, 0))
        self.Body = kw.get("body", "Opening session.")
        self.Recipients = FakeRecipients(kw.get("people", []))


class _FakeFolder:
    def __init__(self, name, items=None):
        self.Name, self.EntryID, self.StoreID = name, name, "store"
        self._items = list(items or [])
        self.Folders = []

    @property
    def Items(self):
        folder = self

        class _Items(list):
            def Restrict(self, _query):
                return list(folder._items)

        return _Items(folder._items)


class _FakeMeeting:
    """An organiser meeting. Send() queues a cancellation in the Outbox."""

    def __init__(self, n, outbox, stubborn=False):
        self.Subject = f"{TEST_MARKER} Session {n}"
        self.Categories = TEST_CATEGORY
        self.EntryID = f"m{n}"
        self._status, self._stubborn = 1, stubborn
        self.GlobalAppointmentID = f"g{n}"
        self.Recipients = FakeRecipients(["someone@example.invalid"])
        self._outbox, self.deleted, self.sends = outbox, False, 0

    @property
    def MeetingStatus(self):
        return self._status

    @MeetingStatus.setter
    def MeetingStatus(self, value):
        if not self._stubborn:  # the 2026-09-21 failure: the set is silently ignored
            self._status = value

    def Save(self):
        pass

    def Send(self):
        self.sends += 1
        kind = "Canceled" if self._status == 5 else "Invitation"
        self._outbox._items.append(_FakeMail(f"{kind}: {self.Subject}", self._outbox))

    def Delete(self):
        self.deleted = True
        for folder in _FakeStore.folders:
            if self in folder._items:
                folder._items.remove(self)


class _FakeMail:
    _n = 0

    def __init__(self, subject, folder):
        self.Subject, self.Categories, self._folder = subject, "", folder
        _FakeMail._n += 1
        self.EntryID = f"x{_FakeMail._n}"
        self.deleted_unsent = False

    def Delete(self):
        if self in self._folder._items and self._folder.Name == "Outbox":
            self.deleted_unsent = True
        self._folder._items.remove(self)


class _FakeStore:
    folders: list = []

    def __init__(self, calendar, outbox, sent, name="organiser@example.invalid"):
        self.DisplayName = name
        self._root = _FakeFolder("root")
        self._root.Folders = [calendar, outbox, sent]
        self._outbox = outbox
        _FakeStore.folders = [calendar, outbox, sent]

    def GetRootFolder(self):
        return self._root

    def GetDefaultFolder(self, n):
        return self._outbox


class _FakeNamespace:
    """SendAndReceive moves everything queued in the Outbox to Sent Items."""

    def __init__(self, n_meetings, stubborn=()):
        self.outbox, self.sent = _FakeFolder("Outbox"), _FakeFolder("Sent")
        self.calendar = _FakeFolder("Calendar")
        self.meetings = [_FakeMeeting(i, self.outbox, i in stubborn) for i in range(n_meetings)]
        self.calendar._items = list(self.meetings)
        self.Stores = [_FakeStore(self.calendar, self.outbox, self.sent)]
        self.sent_mail = []

    def GetItemFromID(self, entry_id, _store_id):
        for folder in (self.calendar, self.outbox, self.sent):
            for item in folder._items:
                if item.EntryID == entry_id:
                    return item
        raise LookupError(entry_id)

    def SendAndReceive(self, _show):
        for mail in list(self.outbox._items):
            self.outbox._items.remove(mail)
            mail._folder = self.sent
            self.sent._items.append(mail)
            self.sent_mail.append(mail)


class TestPurge(unittest.TestCase):
    """The 2026-09-21 bug: the purge deleted its own queued cancellations."""

    def run_purge(self, n=5):
        from autocalendar.outlook import purge_tests

        ns = _FakeNamespace(n)
        counts = purge_tests(apply=True, namespace=ns, pace=0, poll=0, outbox_timeout=1)
        return ns, counts

    def test_every_attendee_is_told_exactly_once(self):
        ns, counts = self.run_purge()
        self.assertEqual(counts["cancelled"], 5)
        self.assertEqual([m.sends for m in ns.meetings], [1] * 5)
        self.assertEqual(len(ns.sent_mail), 5)
        self.assertFalse(any(m.deleted_unsent for m in ns.sent_mail))

    def test_everything_is_gone_afterwards(self):
        ns, counts = self.run_purge()
        self.assertEqual(counts["remaining"], 0)
        self.assertTrue(all(m.deleted for m in ns.meetings))
        self.assertEqual(ns.sent._items, [])  # the sent cancellations are tidied up too

    def test_nothing_is_deleted_while_the_outbox_is_stuck(self):
        from autocalendar.outlook import purge_tests

        ns = _FakeNamespace(3)
        ns.SendAndReceive = lambda _show: None  # offline: nothing ever leaves
        counts = purge_tests(apply=True, namespace=ns, pace=0, poll=0.01, outbox_timeout=0.05)
        self.assertEqual(counts["outbox_stuck"], 3)
        self.assertEqual(counts["deleted"], 0)
        self.assertEqual(len(ns.outbox._items), 3)  # still queued, not destroyed

    def test_a_cancel_that_does_not_stick_is_never_sent_or_deleted(self):
        """Never re-send an invitation, never orphan an attendee's copy."""
        from autocalendar.outlook import purge_tests

        ns = _FakeNamespace(4, stubborn={2})
        counts = purge_tests(apply=True, namespace=ns, pace=0, poll=0, outbox_timeout=1)
        stubborn = ns.meetings[2]
        self.assertEqual(stubborn.sends, 0)
        self.assertFalse(stubborn.deleted)
        self.assertEqual(counts["cancel_failed"], 1)
        self.assertEqual(counts["cancelled"], 3)
        self.assertFalse(any(m.Subject.startswith("Invitation") for m in ns.sent_mail))
        self.assertEqual(counts["remaining"], 1)

    def test_mailbox_filter_leaves_other_mailboxes_alone(self):
        from autocalendar.outlook import purge_tests

        ns = _FakeNamespace(2)
        counts = purge_tests(apply=True, namespace=ns, mailbox="someone-else@example.invalid",
                             pace=0, poll=0, outbox_timeout=1)
        self.assertEqual(counts["found"], 0)
        self.assertFalse(any(m.deleted for m in ns.meetings))


class TestAllDayWrite(unittest.TestCase):
    """A multi-day all-day event must keep its span when written to Outlook."""

    def test_inclusive_sheet_end_becomes_outlooks_exclusive_end(self):
        from autocalendar.sync import _write

        item = FakeItem()
        event = Event(
            title="Live stream",
            start=dt.date(2027, 1, 4),
            end=dt.date(2027, 1, 22),  # last day, inclusive
        )
        _write(item, event)
        self.assertTrue(item.AllDayEvent)
        self.assertEqual(item.Start, "2027-01-04")
        self.assertEqual(item.End, "2027-01-23")

    def test_single_day(self):
        from autocalendar.sync import _write

        item = FakeItem()
        day = dt.date(2027, 1, 12)
        _write(item, Event(title="Check-in", start=day, end=day))
        self.assertEqual(item.End, "2027-01-13")


def sheet_event(**kw):
    return Event(
        title=kw.get("title", "Kick-off"),
        start=kw.get("start", dt.datetime(2026, 8, 3, 9, 0)),
        end=kw.get("end", dt.datetime(2026, 8, 3, 10, 0)),
        location=kw.get("location", "Room 1"),
        description=kw.get("description", "Opening session."),
        required=kw.get("required", []),
        event_id=kw.get("event_id", "aaaa-bbbb-cccc"),
    )


class TestSyncComparison(unittest.TestCase):
    def test_identical_rows_are_unchanged(self):
        self.assertEqual(compare(sheet_event(), FakeItem()), [])

    def test_a_moved_session_is_seen(self):
        moved = sheet_event(start=dt.datetime(2026, 8, 3, 10, 0))
        self.assertIn("start", compare(moved, FakeItem()))

    def test_a_room_change_is_seen(self):
        self.assertIn("location", compare(sheet_event(location="Room 9"), FakeItem()))

    def test_an_added_attendee_is_seen(self):
        event = sheet_event(required=["a@dtu.dk"])
        self.assertIn("attendees", compare(event, FakeItem(people=[])))

    def test_an_unchanged_attendee_list_is_not_a_change(self):
        event = sheet_event(required=["a@dtu.dk"])
        self.assertNotIn("attendees", compare(event, FakeItem(people=["a@dtu.dk"])))


class TestWhoGetsTold(unittest.TestCase):
    """The whole point: only bother people when the change concerns them."""

    def test_a_tidied_description_tells_nobody(self):
        change = Change(sheet_event(required=["a@dtu.dk"]), UPDATED, ["description"])
        self.assertFalse(change.notifies)

    def test_a_moved_session_tells_everyone(self):
        change = Change(sheet_event(required=["a@dtu.dk"]), UPDATED, ["start"])
        self.assertTrue(change.notifies)

    def test_a_room_change_tells_everyone(self):
        change = Change(sheet_event(required=["a@dtu.dk"]), UPDATED, ["location"])
        self.assertTrue(change.notifies)

    def test_a_new_event_without_attendees_tells_nobody(self):
        self.assertFalse(Change(sheet_event(), NEW).notifies)

    def test_an_unchanged_row_tells_nobody(self):
        change = Change(sheet_event(required=["a@dtu.dk"]), UNCHANGED)
        self.assertFalse(change.notifies)

    def test_a_dropped_appointment_that_nobody_saw_tells_nobody(self):
        change = Change(sheet_event(), ORPHANED, detail="appointment")
        self.assertFalse(change.notifies)

    def test_a_dropped_meeting_is_cancelled_to_its_attendees(self):
        change = Change(
            sheet_event(required=["a@dtu.dk"]),
            ORPHANED,
            detail="meeting - would be cancelled",
        )
        self.assertTrue(change.notifies)


class TestPlanTotals(unittest.TestCase):
    def test_a_plan_of_only_unchanged_rows_is_empty(self):
        plan = Plan([Change(sheet_event(), UNCHANGED) for _ in range(5)])
        self.assertTrue(plan.is_empty)
        self.assertEqual(plan.people_contacted, 0)

    def test_only_notifying_changes_count_towards_mail(self):
        plan = Plan(
            [
                Change(sheet_event(required=["a@dtu.dk"]), UPDATED, ["start"]),
                Change(sheet_event(required=["b@dtu.dk"]), UPDATED, ["description"]),
                Change(sheet_event(required=["c@dtu.dk"]), UNCHANGED),
            ]
        )
        self.assertFalse(plan.is_empty)
        self.assertEqual(plan.people_contacted, 1)


if __name__ == "__main__":
    unittest.main()
