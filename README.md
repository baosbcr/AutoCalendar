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
