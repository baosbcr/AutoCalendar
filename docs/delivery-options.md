# Delivery options — sheet to Outlook, with invitations

Status: **path chosen 2026-09-07 — desktop Outlook automation (option A).**
Investigation written 2026-09-04; verified by live testing 2026-09-07.
Two earlier conclusions in this document were wrong and are corrected below.

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

## CORRECTION — the display-name finding was an artefact

The 2026-09-04 version of this document treated the following as decisive:

```
distinct attendee tokens : 227
  with an email address  :  58   (26%)  — mostly external speakers
  display name only      : 169   (74%)  — DTU staff, resolved via the address book
```

and concluded that A and B win because desktop Outlook re-resolves names against the
address book while C and D need real SMTP addresses.

**This was reading a property of the export as a property of the workflow.** Those names
are display names *because Outlook resolved them on the way out* when Toke downloaded his
own calendar. Future editions source attendees from DTU Learn exports and speaker lists,
which carry real email addresses. The 74% describes one CSV, not the job.

The finding is retained here because it was the stated reason for an earlier
recommendation, and the reasoning should not silently disappear. It is no longer a
differentiator between the options.

## CORRECTION — Graph's forced sending is not a real objection

Microsoft's documentation is unambiguous: *"When you create an event that includes
attendees, the server sends invitations to all attendees ... and can't be configured."*
That was recorded here as a loss of control.

It is not. Control lives one level up: the tool chooses **which rows to send in this run**,
so batching gives the same pacing a deferred-send flag would. Sending months ahead is also
normal for this course — invitations to the finals routinely go out before the course
begins. The "send timing" worry was our assumption, not Toke's requirement.

## The other APIs, for completeness

"The Outlook API" is not one thing. Two of the five are already options above under other
names.

| API | Status | Relation to the options |
|---|---|---|
| Outlook Object Model (COM) | Alive, classic Outlook only | **is option A** |
| Microsoft Graph | Current, Microsoft's designated successor | **is option D** |
| Exchange Web Services (EWS) | **Disablement began October 2026; fully disabled April 2027** | do not use |
| Outlook REST API v2.0 | HTTP 410 Gone since March 2024 | dead |
| Office.js Outlook add-in | Alive, survives the new-Outlook transition | wrong shape — bulk creation calls Graph anyway, and it needs a manifest deployed by IT |

EWS deserves the warning: it was the obvious "Python plus the Outlook API" answer for a
decade, `exchangelib` tutorials still recommend it, and it is now the worst possible choice.

## Verification results — 2026-09-07

Tested on João's own machine, DTU student account `<the DTU test account>` (the same Entra tenant
as Toke's mailbox) with classic Outlook and a personal @hotmail account as the guinea pig.

| # | Test | Result |
|---|---|---|
| 1 | COM reachable; real folder names | **Pass.** Both mailboxes visible, DTU is Exchange Online, default calendar folder is `Calendar` |
| 2 | COM creates and sends a real meeting | **Pass.** `Send()` returned in 0.9 s, no security prompt, invitation arrived with working Accept / Tentative / Decline |
| 3 | Graph via Microsoft Graph PowerShell | **Blocked.** `AADSTS50105` — the Graph CLI app requires explicit assignment in this tenant |
| 4 | Register a custom Entra app | **Blocked.** App registrations blade returns 401 |
| 5 | Cancel and purge every trace | **Pass.** Cancellation sent, then 7 artefacts across 4 folders in 2 mailboxes removed, verified zero remaining |
| 6 | Does importing a .ics with `ATTENDEE` send anything? | **No - confirmed firsthand.** See below |

Details worth keeping:

- **No object-model security prompt fired.** This was listed as an unknown; it is not a
  problem on this machine.
- **Timezones behaved.** An event set for 10:00 Copenhagen displayed as 09:00 in a mailbox
  on Lisbon time — correct.
- **Cleanup must cancel before deleting.** Calling `Delete()` on the organiser's copy
  removes it locally and sends nothing, leaving attendees with a ghost meeting forever. The
  correct sequence is `MeetingStatus = olMeetingCanceled` -> `Save()` -> `Send()` ->
  `Delete()`.
- **A cleanup pass cannot catch its own cancellation notice.** The cancellation arrives
  after the sweep has read the folder, so it needs a second pass.
- **Sent Items is easy to forget.** The first cleanup attempt missed the request and the
  cancellation sitting there. Purge must walk every folder, not just Inbox and Calendar.
- **Exchange's Recoverable Items dumpster is out of reach from COM.** Everything
  user-visible can be removed; the retention copy ages out on its own, or is purged
  manually via Folder -> Recover Deleted Items From Server.

### The .ics import, observed rather than assumed

A hand-built `.ics` carrying `METHOD:REQUEST`, an `ORGANIZER` and one external
`ATTENDEE` was imported through the Outlook UI (File > Open & Export > Import) into
the DTU calendar. After 45 seconds, a sweep of every folder in both mailboxes found
exactly one item - the calendar entry itself. Nothing in the attendee's inbox,
nothing in Sent Items, nothing in the Outbox.

What the import produced:

```
class         : IPM.Appointment
MeetingStatus : 1   olMeeting ("I organise this")
organizer     : <the DTU account>
recipients    : 2  - the organiser, resolved against the Exchange directory,
                     and the external attendee, Required, correct address
```

**This is worse than the documentation implies.** Outlook does not degrade the event
into a plain appointment with the attendees dropped, which would at least look wrong.
It creates a *proper meeting*, keeps `MeetingStatus = olMeeting`, resolves the
organiser against the directory and preserves the attendee list - and sends nothing.
On screen the event is indistinguishable from one that did invite people.

That silence is the real risk to Toke: he could build an entire course edition, see
attendees listed on every event, and discover the problem only when nobody arrives.

For contrast, cancelling those same meetings through COM *did* send cancellations to
the same attendee. It is not a delivery problem; importing simply never initiates a
send.

Note also that AutoCalendar's own writer emits `METHOD:PUBLISH` with no `ORGANIZER`
and no `ATTENDEE` at all, so the shipped `.ics` cannot invite anyone by construction -
this test covers the harder case of a file that genuinely asks for invitations.

### Sending limits that matter at course scale

Exchange Online: **30 messages/minute**, **10,000 recipients/day**, **1,000 recipients per
message**, plus a tenant-wide external recipient limit. A 200-student cohort is roughly
seven minutes of steady sending. The tool should pace deliberately rather than fire as fast
as COM allows.

## The decision

**Build option A — Python driving desktop Outlook, shipped as a self-contained executable.**

Not because it is the most elegant. Graph is the better architecture and it survives the
eventual retirement of classic Outlook, which A does not. A wins on availability: two
independent routes into Graph are closed by deliberate DTU policy, so every Graph path
starts with an IT request of unknown duration, while COM did the entire job today with no
permission from anyone.

Honest record of how this conclusion moved, because it moved twice:

1. Favoured A on the display-name finding — which was an artefact.
2. Corrected toward D once that fell, and D also handles bulk sending better and gives
   stable event IDs for the re-run problem.
3. Settled on A when the tenant turned out to be locked.

The caveat on test 4: the 401 is the Entra *portal* refusing access, and Microsoft has a
separate setting that hides the portal while leaving registration possible via API. That
cannot be probed here because the CLI app is also blocked. Both reachable doors are shut;
whether a third exists is a question for DTU IT, written up in the Bookkeeping repo at
`statuses/dtu-it-questions.md`.

Option B (the VBA button) remains a possible nicety later — same engine, nicer for Toke —
but only if DTU permits macro-enabled workbooks, which is unknown and also on the IT list.

## Test safety rules for the build

Agreed 2026-09-07, after proving that test artefacts can be fully removed:

1. Tag every generated test item with a dedicated **category and a hidden user property**,
   never a subject-substring match — that would eventually eat a real event.
2. Confine test events to a **distinctive far-future date window**, so cleanup filters on
   tag *and* date: two independent conditions.
3. Purge **recursively across every folder** in both mailboxes, cancel-before-delete, with
   a second pass for cancellation notices.
4. Use **`Save()` for volume runs and `Send()` only for small deliberate sends.** "Does the
   loop survive 117 rows" and "does an invitation arrive" are separate questions, and only
   the second needs mail to leave the building.

## Re-running after edits - solved 2026-09-07

Every published recipe punts on this, which is why it was the last real risk:

| Source | How it handles a second run |
|---|---|
| Slipstick VBA | No duplicate detection; mark rows "Imported" and skip |
| TheStaticTurtle (Python/COM) | Deletes all future events and recreates; reports Outlook deletion is *flaky*, needing repeated passes |
| Power Automate thread | Avoids duplicates only by filtering to a narrow date window |

None of those work for a course edition, which is edited over weeks. "Skip
imported rows" never applies a change. "Delete and recreate" re-invites everyone
to everything, every time.

### Identity: one column, written once

Nothing already in the sheet can say which row is which event. Titles repeat
across days (`PA, Team Allocation Support` appears on several), dates are exactly
what gets edited, and row numbers move when a line is inserted. So the tool adds
an **`AutoCalendar ID`** column and fills it on the first run, and stamps the same
id on the Outlook item as a hidden user property. A row can then be retitled,
moved to another week or dragged elsewhere in the sheet and still be recognised.

This is the only place AutoCalendar writes to a file the user owns. It writes to a
temporary file in the same folder and swaps it in only on success, and never
overwrites an id that is already there.

Considered and rejected: a sidecar state file (breaks silently the moment the
sheet is renamed, copied or emailed - all normal for Toke), and matching on title
plus nearest date (genuinely ambiguous on his real data).

### What a re-run does

It reports, and sends nothing, until told twice:

```
CHANGES SINCE THE LAST RUN

     1  new                      (would contact 1)
     4  changed                  (would contact 3)
     1  no longer in the sheet   (nobody contacted)
     5  unchanged                (nobody contacted)

  + NEW: Debrief                          -> 1 person
  ~ CR, Kick-off ...          [start]     -> 1 person
  ~ SF, Day 2 Team Formation  [location]  -> 1 person
  ~ MR, Day 3 ...             [description]  (silent)
  ~ Student Room ...          [attendees] -> 1 person
  - Finals                    [appointment]  (silent)

  Applying this would put mail in 4 inboxes.
```

The line that matters is the silent one. **Start, end, location, title and
attendees notify; everything else is saved quietly.** A tidied description does
not put mail in 200 inboxes.

An event dropped from the sheet is cancelled properly if people were invited, and
simply deleted if it was never sent - so an unsent draft bothers nobody.

### Verified behaviour

Run against a live DTU mailbox on 2026-09-07 with a ten-row sheet:

- **Run 1** created 10 events and wrote 10 ids back into the sheet.
- The sheet was then edited the way Toke edits one: a session moved an hour, a
  room changed, a description tidied, an attendee added, a row added, a row
  deleted.
- **Run 2** classified all six correctly and named the three people who would be
  contacted; applying it created 1, updated 4, removed 1, and sent exactly 4
  messages.
- **Run 3, with nothing edited, reported 10 unchanged and zero inboxes.**

That last one is the property that makes it safe to run repeatedly, and it is the
one every published approach fails.

## Gotchas register

- **Localised calendar folder names** — a public script broke on `Calendrier` vs `Calendar`,
  and this document assumed `Kalender` was likely for a Danish university. **Corrected:**
  Toke's Outlook runs an English display language with Danish regional formats, and his
  folders read `Calendar — <his DTU address>`. The DTU test account is the same. Still do not
  hard-code the name — but this is not the trap it was written up as.
- **Multiple calendars** — the video shows at least six (his main work calendar plus
  `Innovation in En...`, `Facilitator Info`, `innovator info`, ...), plus shared team
  calendars. "Which calendar" is a column, not a constant.
- **Timezone shifting** when start and end timezones are set independently. Verified
  behaving correctly in testing; AutoCalendar writes `Europe/Copenhagen`.
- **Cancel before delete**, always — see the verification section.
- **Sent Items holds copies** of both the request and the cancellation.
- **Outlook deletions need repeat passes.** Prior art warned of it and we reproduced it:
  one run cleared 59 items and left 10, another needed four passes. The purge now loops
  until a verification sweep comes back empty rather than assuming a fixed number.
- **`.ics` files may be associated with the new Outlook**, which on a stock Windows 11
  machine with both versions installed dead-ends at a sign-in prompt whose Continue
  button does nothing. "Just double-click the file" is not a safe instruction to give.
- **`OpenSharedItem` on a `METHOD:REQUEST` file returns a `MeetingItem`, not an
  appointment** - no `.Start`, and `GetAssociatedAppointment` returns `None` while the
  item is not in a store. Relevant to anyone automating around .ics files.
- **`OUTLOOK.EXE /f` expects a `.msg`**, not a `.ics`; it fails with a misleading
  "file is already open, or you don't have permission" dialog.

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

1. ~~**Send timing**~~ — **resolved, no longer blocking.** Invitations normally go out well
   in advance for this course, and the tool controls pacing by choosing which rows to send.
2. **What is a "course edition"** concretely — same structure with shifted dates and
   swapped speakers? Should the tool offer a "shift all dates by N weeks"?
3. **Should the Categories gate sending?** e.g. only `Complete` rows get invites;
   `Placeholder` rows become calendar blockers with no attendees.
4. **Which calendar** should events land in, given he keeps at least six? *This is now the
   main open question.*
5. ~~**The ~169 missing email addresses**~~ — **withdrawn.** See the correction above; the
   names were an artefact of his export, and future editions come from lists that carry
   real addresses.

## Testing constraint — resolved

Previously: João's personal mailbox is @hotmail and was web-only, so the invite test could
not run. **Resolved 2026-09-07** — the account was added to classic Outlook on BPC
alongside the DTU student account, which is what made the end-to-end invitation test
possible. Both mailboxes are reachable from one COM session, which is also what allowed the
purge to verify both sides.

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
- Microsoft — Deprecation of Exchange Web Services in Exchange Online: https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online
- Microsoft — Outlook REST API v2.0 decommissioning: https://devblogs.microsoft.com/microsoft365dev/outlook-rest-api-v2-0-and-beta-endpoints-decommissioning-update/
- Microsoft — Recipients.Add (accepts display name, alias or SMTP): https://learn.microsoft.com/en-us/office/vba/api/outlook.recipients.add
- Microsoft — AppointmentItem.MeetingStatus: https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.meetingstatus
- Microsoft — Macros from the internet are blocked by default: https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked
- Microsoft — Exchange Online limits: https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits
- Microsoft Graph permissions — Calendars.ReadWrite (delegated: no admin consent): https://graphpermissions.merill.net/permission/Calendars.ReadWrite
