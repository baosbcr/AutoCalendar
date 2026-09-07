# AutoCalendar

Turn the Excel planning sheet for a programme — dates, times, rooms, subject
names, content — into a single `.ics` calendar file that can be imported
("installed") into Outlook, Google Calendar or Apple Calendar.

First mock-up for DTU Entrepreneurship.

```
examples/programme_autumn2026.xlsx  ->  examples/programme_autumn2026.ics
```

## Try it in one minute

```powershell
python examples/make_example_sheet.py                  # build the example sheet
python -m autocalendar examples/programme_autumn2026.xlsx
```

```
Read programme_autumn2026.xlsx [Programme]
  Header row 3. Columns used:
    Date         <- Date
    Start time   <- Time
    End time     <- End time
    Duration     <- Duration
    Subject      <- Subject
    Location     <- Location
    Description  <- Content
    Category     <- Track
  Other columns (added to the event description): Underviser
  7 event(s) found
  ! row 11: skipped - could not read 'TBD' as a date (for 'Follow-up session ...')
  07 Sep 2026 - 18 Sep 2026, timezone Europe/Copenhagen

Wrote examples\programme_autumn2026.ics
```

Then double-click the `.ics`, or import it (see [Installing the calendar](#installing-the-calendar)).

## Requirements

Python 3.10 or newer. `openpyxl` is needed only to read `.xlsx`:

```powershell
pip install -r requirements.txt
```

CSV files need nothing beyond Python itself, and the `.ics` writer is
dependency-free.

## Using it

```powershell
python -m autocalendar SHEET.xlsx [options]
```

Run it with **no arguments** and it opens a file picker — handy for anyone who
would rather not use a terminal.

| Option | What it does |
| --- | --- |
| `-o out.ics` | where to write the calendar (default: next to the input) |
| `-s Programme` | which worksheet to read (name or `0`-based index) |
| `-n "Autumn 2026"` | the calendar name the calendar app shows |
| `--list` | show what was read and write nothing (dry run) |
| `--duration 90` | length of a session when the sheet does not say (default 60 min) |
| `--alarm 15` | add a reminder 15 minutes before every session |
| `--tz Europe/Berlin` | timezone of the times in the sheet (default `Europe/Copenhagen`) |
| `--no-extra-columns` | do not append unrecognised columns to the description |
| `--strict` | fail with an error code if any row could not be converted |

Start with `--list`: it prints how each column was understood and which rows
were skipped, without writing anything.

## What the spreadsheet may look like

There is no template to fill in. The converter looks for the column names
people already use, in English or Danish, and it does not matter which order
they are in or how far down the sheet the header row sits.

| We look for | Column names recognised (any of) |
| --- | --- |
| **Subject** *(required)* | Subject, Title, Topic, Session, Activity, Course, Emne, Titel, Fag, Aktivitet |
| **Date** *(required)* | Date, Day, Start date, Dato, Dag, Startdato |
| Start time | Start, Start time, Time, From, Tid, Starttid, Kl, Tidspunkt, Fra |
| End time | End, End time, To, Until, Sluttid, Til |
| Duration | Duration, Length, Minutes, Varighed, Længde, Minutter |
| Location | Location, Room, Venue, Building, Address, Lokale, Sted, Bygning, Lokation |
| Content | Content, Description, Notes, Agenda, Indhold, Beskrivelse, Noter |
| End date | End date, Slutdato (for events spanning several days) |
| Category | Category, Type, Track, Tag, Kategori |
| Link | URL, Link, Teams, Zoom |
| All day | All day, Heldag |

Anything it does not recognise — `Underviser`, `Responsible`, `Sign-up link` —
is not thrown away: it is appended to the event description as
`Underviser: Toke Reichstein`, so nothing from the sheet is lost.

**Formats it accepts**

- Dates: real Excel dates, `2026-09-03`, `03/09/2026`, `3.9.26`,
  `3 Sep 2026`, `3. september 2026`, `torsdag d. 3. september 2026`
- Times: `9`, `09:00`, `9.30`, `0900`, `1:30 pm`, Excel time cells
- A whole range in one cell: `09:00-12:00`, `9 – 12`, `13:00 til 16:00`
- Durations: `90`, `1h30`, `1.5h`, `2 timer`, `45 min`

**What it does with an incomplete row**

- No start time → an **all-day** event on that date
- Start but no end → the default duration (`--duration`, one hour)
- End earlier than start → assumed to run past midnight, and reported
- No subject, or a date it cannot read (`TBD`) → the row is **skipped and
  listed**, never dropped in silence

## Sending real invitations (Windows + classic Outlook)

A `.ics` file cannot invite anyone. Outlook reads its attendee list when importing,
stores it for reference, and sends nothing - the event looks completely correct and
silently invites nobody. To actually send meeting requests the tool drives Outlook
itself.

Add an `Attendees` column to the sheet (and optionally `Optional` and `Calendar`),
then:

```powershell
# create every event as a draft - nothing is sent
python -m autocalendar programme.xlsx --outlook

# send the invitations, for the first 10 rows only
python -m autocalendar programme.xlsx --outlook --send --limit 10
```

Saving is the default. `--send` is opt-in and asks for a typed confirmation.

### Running it again after the sheet changes

The second run is the one that matters, because a course edition is edited over
weeks. `--sync` compares the sheet against the calendar and reports what it would
do, without touching anything:

```
CHANGES SINCE THE LAST RUN

     1  new                      (would contact 1)
     4  changed                  (would contact 3)
     1  no longer in the sheet   (nobody contacted)
     5  unchanged                (nobody contacted)

  ~ SF, Day 2 Team Formation  [location]     -> 1 person
  ~ MR, Day 3 ...             [description]  (silent)
```

Start, end, location, title and attendee changes notify the people involved.
Everything else is saved quietly, so tidying a description does not mail 200
students. `--apply` carries the plan out; add `--send` to let it notify anyone.

To make this work the tool adds an **`AutoCalendar ID`** column to the sheet and
fills it on the first run. That column is how a later run knows which row is which
event, so a session can be retitled, moved to another week or dragged elsewhere in
the sheet and still be recognised. **Do not edit or delete it.**

### Rehearsing, and cleaning up

`--test-mode` shifts an entire programme into a tagged window in 2099, preserving
weekdays and times, so the whole thing can be run at full scale without any risk of
being mistaken for real work. `--purge-tests` then removes every trace - from every
folder of every mailbox, cancelling before deleting so nobody is left holding a
meeting that no longer exists.

```powershell
python -m autocalendar programme.xlsx --outlook --test-mode
python -m autocalendar --purge-tests
```

### What this needs

Windows, with **classic** Outlook installed, running and signed in. The new Outlook
supports no automation of this kind. The events themselves are stored on the server,
so they appear in every client - new Outlook, the web, a phone - it is only the tool
that needs the classic one.

Also `pip install pywin32`.

## Installing the calendar

- **Outlook (desktop)** — File → Open & Export → Import/Export → *Import an
  iCalendar (.ics) file* → choose **Import** to add the events to your own
  calendar, or **Open as New** to keep them in a separate calendar you can
  switch off.
- **Outlook on the web / Microsoft 365** — Calendar → Add calendar → Upload
  from file.
- **Google Calendar** — Settings → Import & export → Import, and pick which
  calendar to add them to.
- **Apple Calendar** — File → Import.

Re-exporting an unchanged sheet produces the same event IDs, so importing
again updates the existing entries instead of duplicating them. Changing a
session's time, room or title makes it a new entry — for a mock-up, importing
a revised programme into a *fresh* calendar is the safest way to work.

Times are written as local wall-clock times with a `Europe/Copenhagen`
timezone definition, so summer/winter time is handled by the calendar app.

## Repository layout

```
autocalendar/
  cli.py         command line, dry-run report, file picker
  reader.py      finds the header row, maps columns, turns rows into events
  parsing.py     the messy-cell parsers: dates, times, ranges, durations
  ics.py         RFC 5545 writer - escaping, 75-octet folding, CRLF
  timezones.py   VTIMEZONE blocks for European zones
  model.py       the Event dataclass and its stable UID
examples/        the example sheet and the script that generates it
tests/           27 tests, stdlib unittest
```

Run the tests with:

```powershell
python -m unittest discover -s tests
```

## Notes and next steps

Deliberately out of scope for a first mock-up, in rough order of usefulness:

1. **Recurring sessions** — `RRULE` for "every Tuesday for 12 weeks" instead
   of one row per occurrence.
2. **A published calendar feed** — host the `.ics` at a URL and subscribe to
   it, so changes to the sheet reach everyone's calendar without re-importing.
3. **Invitations** — adding participants as `ATTENDEE`s so the sessions land
   in students' calendars as invites rather than an import they must do
   themselves.
4. **A drag-and-drop page** — for handing to someone without Python.
5. **Room booking cross-check** — validating locations against DTU's room list.

## Licence

MIT — see [LICENSE](LICENSE).
