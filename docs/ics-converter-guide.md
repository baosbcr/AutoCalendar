# AutoCalendar: from a calendar sheet to your Outlook calendar

**What it does.** You edit a list of events in Excel. AutoCalendar turns it into one
calendar file (`.ics`), and Outlook imports every event in one go.

**What it does not do: it never invites or notifies anyone.** Importing the file only puts the
events in *your own* calendar, with you as the organizer. Invitees in the sheet are kept
for the later invitation run and are left out of the file on purpose. This makes it safe
for placeholders.

---

## 1. Get the sheet

Start from an Outlook calendar export, either the one we prepare for you or your own:

> Outlook (classic): **File > Open & Export > Import/Export > Export to a file >
> Comma Separated Values**, choose the calendar, then the date range.

Open the `.csv` in Excel and save it as an Excel workbook: **File > Save As > Excel
Workbook (.xlsx)**. The raw `.csv` works too, but `.xlsx` survives editing better.

## 2. Edit it

One row is one event. Here is what you can change:

- **Delete** the rows (events) you do not want.
- **Rename** events in *Subject*. Remove last edition's speaker initials
  (e.g. `SF, Day 5 ...`) so they do not carry over.
- **Move** events: change *Start Date* / *End Date* and the times.
- **Invitees** (*Required Attendees*, *Optional Attendees*): edit them freely. They are
  kept for the real invitation run later and are **not** imported now.

The column names must stay as they are. Their order does not matter, and any other
column in Outlook's export is ignored.

| Column | Needed? | Write it like |
| --- | --- | --- |
| **Subject** | yes | `CR, Day 2 Introduction Morning - Innovation in Engineering DTU` |
| **Start Date** | yes | `05-01-2027` (day-month-year) |
| Start Time / End Time | for timed events | `09:00` or `09:00:00` |
| End Date | if it ends on another day | `05-01-2027` |
| All day event | for all-day events | `TRUE`, see below |
| Location | optional | `Byg 303A, Aud. 42 DTU Lyngby` |
| Description | optional | any text, several lines are fine |
| Categories | optional | `Placeholder` (several: `Placeholder;Complete`) |
| Required / Optional Attendees | optional, not imported | `thow@dtu.dk; name@dtu.dk` |

**All-day events.** Set *All day event* to `TRUE`, then give the end one of two ways:

- as Outlook exports it: *End Date* is the **day after** the last day, and *End Time* is `00:00`.
  A 4-22 January event is `04-01-2027` to `23-01-2027 00:00`.
- or leave *End Time* empty and put the **last day** itself in *End Date* (`22-01-2027`).

A row with neither *Start Time* nor *End Time* also becomes an all-day event.

## 3. Convert

Open **https://baosbcr.github.io/AutoCalendar/** and drop your sheet on the page (or click
to choose it). The page runs the converter **inside your browser: the file is never
uploaded anywhere.** It shows:

- how many events it read, and a list of them to check,
- which columns it used,
- **every row it could not read, and why.** Nothing is dropped without being named.
  Fix those rows in Excel and drop the sheet again.

**Last-minute changes on the page.** Untick the events you do not want, and click any
subject, date, time or location to change it. Clear both times to make an event all-day.
**Download .ics** then takes only the ticked events, with your changes.

These changes are not saved into your original file. Click **Download edited sheet** to
keep them: that sheet is the one to keep for the later invitation run, so the sheet and
your calendar stay the same.

*Offline alternative:* **AutoCalendar.exe** does the same thing. Double-click it and pick
the sheet, and it writes the `.ics` next to the sheet. Windows may block it, because it is
new and not signed. The web page avoids that.

## 4. Import into Outlook

In **classic Outlook**: **File > Open & Export > Import/Export > Import an iCalendar
(.ics) or vCalendar file**, choose the file, then click **Import**.

- **Import** adds the events to your main calendar.
- **Open as New** puts them in a separate calendar first. That is a good way to look them over, and
  you can then delete that calendar. If they look right, import again with **Import**.

Use the menu, not a double-click on the `.ics`: on some PCs a double-click opens the
*new* Outlook, which gets stuck at a sign-in screen.

## Good to know

- **Import once, when the sheet is final.** Importing the same file again updates the
  same events. An event you **changed** in the sheet comes in as a **new** event next to the
  old one. To redo an event, delete the old one in Outlook first.
- After importing, edit events **in Outlook** as usual: add invitees, write the agenda, attach files, send.
- **Categories:** the category *name* is in the file. Whether Outlook's import keeps it,
  and colours it from your own category list (`Placeholder` in yellow), is still to be
  confirmed in the first import.
