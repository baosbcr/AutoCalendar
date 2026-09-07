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


if __name__ == "__main__":
    unittest.main()
