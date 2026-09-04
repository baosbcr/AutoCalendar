# Delivery options — sheet to Outlook, with invitations

Status: **investigation complete, no path chosen.** Written 2026-09-04.
Awaiting: an expert read on the three original options, a Copilot 365 answer on
what the DTU tenant allows, and Toke's answer on send timing.

---

## What Toke asked for

From his 2 min 09 s Teams video (transcript: `from-toke/video-transcript.txt`;
auto-generated, one typo corrected — "media organisers" → "meeting organisers"):

> "I have downloaded my own calendar and managed to get most of the information
> out ... can we create a new Excel file that lists all of these events up? We can
> adjust the dates super quickly, start and end date, we can add **meeting
> organisers and additional attendees** and so on ... once the course is created,
> we generate a calendar file that very quickly just **installs all of these
> events, instead of us going in and creating it manually, which allows us to make
> mistakes** ... I think we only should do a little test — a **sample of 10
> events**. How can we get those into the calendar? The end goal: a full course
> plan like this, all events in, very easy, **just with one click**."

Broken into requirements:

| # | Requirement | Feasible? |
|---|---|---|
| R1 | Whole programme as one spreadsheet, one row per event | Yes — `reader.py` already does this |
| R2 | Edit there: shift dates, set times, set organiser + attendees | Yes — columns |
| R3 | One click, all events land in his Outlook calendar | Yes — `.ics`, already built |
| R4 | **Invitations actually sent to attendees** | **No — not via `.ics`. This is the whole problem.** |
| R5 | Start with a 10-event sample, not all 117 | Yes |

---

## The constraint that drives everything

Outlook parses `ATTENDEE` properties when importing a `.ics` file, but **stores
them for reference only and sends no invitations.** Forcing `METHOD:REQUEST` does
not change this. Confirmed independently of our own testing.

So R4 requires something that *drives Outlook* rather than handing it a file.
That is a different and larger piece of work than the current `.ics` writer, and
it is the single fact Toke most needs to know before he sets expectations.

The `.ics` writer stays useful either way as a zero-risk preview / dry-run.

---

## The four options

### A. Desktop automation — Python + `pywin32` (COM)

Script drives Toke's own installed Outlook: creates `AppointmentItem`, sets
`MeetingStatus = olMeeting`, adds `Recipients`, calls `.Send()`. Real meetings,
real invitations, real RSVP tracking.

- **Sends invitations:** yes
- **IT approval needed:** none
- **Resolves display-name attendees:** **yes** — `Recipients.Resolve()` queries the
  Exchange Global Address List, exactly as typing a name into Outlook does
- **Fit with his setup:** excellent. Video shows *classic* Outlook (new-Outlook
  toggle Off) on Exchange, Developer tab already enabled
- **Cost:** runs only on his Windows PC with Outlook open; we maintain the script
- **Reuses:** all of AutoCalendar's existing sheet reader and parsers

### B. Excel VBA macro button

Same COM mechanism, packaged as a button inside the workbook. Closest to Toke's
"one click"; nothing for him to install.

- **Sends invitations:** yes
- **IT approval needed:** none (macro-enabled workbook must be trusted)
- **Resolves display names:** yes (same GAL resolution)
- **Cost:** hard to version-control, test, or hand over; macro security warnings

### C. Power Automate flow

Excel table on OneDrive/SharePoint → "List rows present in a table" → "Apply to
each" → Office 365 Outlook "Create event (V4)". Microsoft publishes this recipe
itself. Graph underneath, but Microsoft has already done the app registration.

- **Sends invitations:** yes
- **IT approval needed:** usually none, if the connector is not restricted
- **Resolves display names:** **no** — needs SMTP addresses
- **Cost:** weak at update-on-re-run logic; clunky to version-control; lives in
  his account, so handover is easy but our ability to debug it is limited

### D. Microsoft Graph API

`POST /me/events` with an `attendees` array. Microsoft's docs confirm invitations
are sent automatically when `attendees` is present — no extra flag.

- **Sends invitations:** yes
- **IT approval needed:** **Azure AD app registration at DTU, possibly admin
  consent.** Unknown timeline — this is the risk
- **Resolves display names:** **no** — needs SMTP addresses
- **Cost:** cleanest and headless; works without Outlook open; best long-term
- **Trap:** `sendUpdates` is *not* a supported query parameter on `POST /me/events`
  — it is silently ignored, so send behaviour looks controllable but is not

### Ruled out

**`.ics` alone** — cannot satisfy R4. Retained as preview/dry-run only.

---

## The finding that re-ranks these

Toke's exported attendees are overwhelmingly **display names, not addresses**:

```
distinct attendee tokens : 227
  with an email address  :  58   (26%)  — mostly external speakers
  display name only      : 169   (74%)  — DTU staff, resolved via the address book
```

Outlook resolved these against DTU's Exchange address book and the CSV export
wrote back the *resolved name*. Consequences:

- **A and B work as-is** — desktop Outlook re-resolves names against the GAL.
- **C and D do not** — both need SMTP addresses. Either Toke hand-fills ~169
  addresses, or we build a name-to-address lookup step first.

This partly reverses the intuitive "Graph is cleanest" conclusion and is the
strongest evidence currently on the table.

---

## The unsolved part: re-running after edits

Every published recipe **punts on this**, which means it is real engineering, not
something we get for free:

| Source | How it handles a second run |
|---|---|
| Slipstick VBA | No duplicate detection; mark rows "Imported" and skip |
| TheStaticTurtle (Python/COM) | Deletes all future events and recreates; reports Outlook deletion is *flaky*, needing repeated passes |
| Power Automate thread | Avoids duplicates only by filtering to a narrow date window |

Notably, TheStaticTurtle's script — the closest public match to our tool —
**deliberately omits attendees and invitations** to avoid disrupting shared
calendars. The one person who built our tool avoided our hard part.

"Update in place, notify only the affected attendees" is where the actual value
is, and it is unsolved in public. Decision deferred until the 10-event sample
lands.

---

## Gotchas register

- **Localised calendar folder names** — TheStaticTurtle's script broke on
  `Calendrier` vs `Calendar`. Toke is Danish at a Danish university: assume
  `Kalender` is possible, never hard-code the folder name.
- **Multiple calendars** — the video shows at least six (his main work calendar plus
  `Innovation in En...`, `Facilitator Info`, `innovator info`, ...). "Which
  calendar" is a column, not a constant.
- **Timezone shifting** when start and end timezones are set independently.
  AutoCalendar already writes `Europe/Copenhagen`.
- **Fully-qualified addresses** required by C and D — see the display-name finding.

---

## What the source data actually is

`docs/from-toke/toke-outlook-calendar-aug2026.csv` (**local only — gitignored,
see note below**) — an **Outlook CSV export**, not a
hand-built sheet. 117 events, 03–31 Aug 2026, the real Innovation in Engineering
August 2026 edition.

- Standard Outlook export columns (Subject, Start/End Date + Time, Required /
  Optional Attendees, Meeting Resources, Categories, Description, Location,
  Show time as, Reminder ...)
- Subjects carry a speaker/role prefix: `SF, Day 5 Working in Multicultural...`,
  `CR, Kick-off...`, `PA, Team Allocation Support`
- **Categories are used as a workflow status**: `Complete`, `Placeholder`,
  `Missing Agenda/Content`, `Programme Assistant`, `Course_01` — 61 of 117 tagged
- Descriptions are full speaker-briefing emails (longest ~15 000 chars)
- Reminders off on 114/117; 8 all-day events

---

## Open questions for Toke

1. **Send timing** — invitations immediately on import, a separate deliberate send
   step, or a per-row "send on" date? Sending months early to external speakers
   may not be wanted.
2. **What is a "course edition"** concretely — same structure with shifted dates
   and swapped speakers? Should the tool offer a "shift all dates by N weeks"?
3. **Should the Categories gate sending?** e.g. only `Complete` rows get invites;
   `Placeholder` rows become calendar blockers with no attendees.
4. **Which calendar** should events land in, given he keeps at least six?
5. **The ~169 missing email addresses** — is there an export that includes them,
   or is address-book resolution (options A/B) the answer?

---

## Testing constraint

João's personal mailbox is **@hotmail, used via Outlook web only** — no desktop
Outlook for it, so the A/B invite test cannot run against it as things stand.
Classic Outlook *is* installed on BPC and accepts Outlook.com accounts, so adding
it there is the cheap fix. Graph also works against personal Microsoft accounts,
so D could be tested without touching DTU at all.

---

## A note on the source files

`docs/from-toke/` is **gitignored and never pushed.** This repository is public,
and Toke's calendar export carries 227 named staff and students, 58 live email
addresses (several personal), and full briefing text. The folder holds the CSV,
the video transcript and two screenshots; they live on the working machine only.
Anyone picking this up on another machine needs those files re-supplied by hand.

---

## Sources

- Slipstick — Create Appointments Using Spreadsheet Data: https://www.slipstick.com/developer/create-appointments-spreadsheet-data/
- Microsoft Power Platform Blog — Creating appointments from an Excel table: https://www.microsoft.com/en-us/power-platform/blog/power-automate/flow-of-the-week-creating-appointments-from-excel-table/
- Microsoft Graph — Create event: https://learn.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0
- TheStaticTurtle — Automating my Outlook calendars: https://blog.thestaticturtle.fr/automating-my-outlook-calendars/
- Power Platform community — Flow to send invites from Excel: https://community.powerplatform.com/forums/thread/details/?threadid=48875206-6a16-4d8c-b194-d71e94c2ee92
- nishanthrj/Outlook-Meeting-Invite (closest prior art): https://github.com/nishanthrj/Outlook-Meeting-Invite
- Outlook ICS METHOD:REQUEST behaviour: https://en.ittrip.xyz/ms-office/outlook/ics-file-issue-outlook
