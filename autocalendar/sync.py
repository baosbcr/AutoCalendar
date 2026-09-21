"""Work out what a second run should do to events that already exist.

This is the part every published recipe skips. The ones that exist either
mark rows as "imported" and never update them again, or delete every future
event and recreate the lot - which re-invites everybody, for every event,
every time.

Neither is acceptable here. A course edition is edited over weeks: a room
changes, a session moves an hour, a speaker is swapped. The people already
invited need to hear about the first two and must not be spammed about the
hundred events that did not change.

So this module compares the sheet against the calendar and classifies every
row. It sends nothing. It produces a plan, which is printed for a human to
read, and only a second explicit command carries it out.

The distinction that matters is between changes attendees must be told about
- when it starts, when it ends, where it is, what it is called, who is
invited - and changes that are nobody's business but the organiser's, like a
tidied description. The first kind sends an update. The second is saved
quietly.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .model import Event
from .outlook import OL_APPOINTMENT, PROP_UID

#: Changing any of these means every attendee needs to hear about it.
NOTIFIABLE = ("start", "end", "location", "title", "attendees")

NEW = "new"
UNCHANGED = "unchanged"
UPDATED = "updated"
ORPHANED = "orphaned"


@dataclass
class Change:
    event: Event
    kind: str
    fields: list[str] = field(default_factory=list)
    item = None  # the live Outlook item, when there is one
    detail: str = ""

    @property
    def notifies(self) -> bool:
        """Would carrying this out put mail in somebody's inbox?"""
        if self.kind == NEW:
            return bool(self.event.required or self.event.optional)
        if self.kind == ORPHANED:
            return self.detail.startswith("meeting")
        if self.kind == UPDATED:
            return any(f in NOTIFIABLE for f in self.fields)
        return False

    @property
    def people(self) -> int:
        return len(self.event.required) + len(self.event.optional)


@dataclass
class Plan:
    changes: list[Change] = field(default_factory=list)

    def of(self, kind: str) -> list[Change]:
        return [c for c in self.changes if c.kind == kind]

    @property
    def notifying(self) -> list[Change]:
        return [c for c in self.changes if c.notifies]

    @property
    def people_contacted(self) -> int:
        return sum(c.people for c in self.notifying)

    @property
    def is_empty(self) -> bool:
        return not any(c.kind != UNCHANGED for c in self.changes)


def _text(value) -> str:
    return (value or "").strip()


def _moment(value) -> dt.datetime | None:
    """Outlook hands back tz-aware datetimes; the sheet is wall-clock."""
    if value is None:
        return None
    try:
        return dt.datetime(
            value.year, value.month, value.day, value.hour, value.minute
        )
    except Exception:
        return None


def read_ids(folder) -> dict[str, object]:
    """Map our id onto the live Outlook items in a calendar folder."""
    found: dict[str, object] = {}
    try:
        items = folder.Items
    except Exception:
        return found
    for item in items:
        try:
            prop = item.UserProperties.Find(PROP_UID)
        except Exception:
            prop = None
        if prop is None:
            continue
        try:
            value = str(prop.Value).strip()
        except Exception:
            continue
        if value:
            found[value] = item
    return found


def _attendees_of(item) -> set[str]:
    people = set()
    try:
        for recipient in item.Recipients:
            address = _text(getattr(recipient, "Address", "")).lower()
            name = _text(getattr(recipient, "Name", "")).lower()
            people.add(address or name)
    except Exception:
        pass
    return people


def compare(event: Event, item) -> list[str]:
    """Which fields of a live Outlook item disagree with the sheet."""
    differences = []

    if _text(item.Subject) != _text(event.title):
        differences.append("title")
    if _text(getattr(item, "Location", "")) != _text(event.location):
        differences.append("location")

    if not event.all_day:
        if _moment(getattr(item, "Start", None)) != event.start:
            differences.append("start")
        if _moment(getattr(item, "End", None)) != event.end:
            differences.append("end")

    body = _text(getattr(item, "Body", ""))
    if body.replace("\r\n", "\n") != _text(event.description):
        differences.append("description")

    wanted = {p.strip().lower() for p in (event.required + event.optional) if p.strip()}
    if wanted:
        current = _attendees_of(item)
        # The organiser is always in Recipients; only additions matter here.
        if not wanted.issubset(current):
            differences.append("attendees")

    return differences


def build_plan(events: list[Event], folder) -> Plan:
    """Compare the sheet against the calendar. Changes nothing."""
    live = read_ids(folder)
    plan = Plan()
    seen: set[str] = set()

    for event in events:
        if not event.event_id:
            plan.changes.append(Change(event, NEW, detail="no id yet"))
            continue
        item = live.get(event.event_id)
        if item is None:
            plan.changes.append(Change(event, NEW))
            continue
        seen.add(event.event_id)
        differences = compare(event, item)
        change = Change(
            event,
            UPDATED if differences else UNCHANGED,
            fields=differences,
        )
        change.item = item
        plan.changes.append(change)

    for event_id, item in live.items():
        if event_id in seen:
            continue
        ghost = Event(
            title=_text(item.Subject),
            start=_moment(getattr(item, "Start", None)) or dt.datetime.now(),
            end=_moment(getattr(item, "End", None)) or dt.datetime.now(),
            event_id=event_id,
        )
        try:
            is_meeting = item.MeetingStatus == 1 and item.Recipients.Count > 1
        except Exception:
            is_meeting = False
        change = Change(
            ghost,
            ORPHANED,
            detail="meeting - would be cancelled" if is_meeting else "appointment",
        )
        change.item = item
        try:
            change.event.required = [
                str(r.Name) for r in item.Recipients if str(r.Name)
            ][1:]
        except Exception:
            pass
        plan.changes.append(change)

    return plan


def describe(plan: Plan) -> str:
    """The report a human reads before deciding to apply anything."""
    lines = ["CHANGES SINCE THE LAST RUN", ""]
    counts = [
        (NEW, "new"),
        (UPDATED, "changed"),
        (ORPHANED, "no longer in the sheet"),
        (UNCHANGED, "unchanged"),
    ]
    for kind, label in counts:
        group = plan.of(kind)
        if not group:
            continue
        contacted = sum(c.people for c in group if c.notifies)
        note = f"would contact {contacted}" if contacted else "nobody contacted"
        lines.append(f"  {len(group):>4}  {label:<24} ({note})")

    for kind, marker in ((NEW, "+"), (UPDATED, "~"), (ORPHANED, "-")):
        group = plan.of(kind)
        if not group:
            continue
        lines.append("")
        for change in group:
            what = ", ".join(change.fields) if change.fields else change.detail
            tail = f"  [{what}]" if what else ""
            n = change.people
            who = f"  -> {n} {'person' if n == 1 else 'people'}" if change.notifies else "  (silent)"
            lines.append(f"  {marker} {change.event.title[:52]}{tail}{who}")

    lines.append("")
    lines.append(
        f"  Applying this would put mail in {plan.people_contacted} "
        f"{'inbox' if plan.people_contacted == 1 else 'inboxes'}."
    )
    return "\n".join(lines)


def apply_plan(
    plan: Plan, folder, *, send: bool = False, response_requested: bool = True, on_progress=None
) -> dict:
    """Carry out a plan. Only sends where the change warrants it."""
    from .outlook import OL_MEETING, OL_MEETING_CANCELED, OL_REQUIRED, OL_OPTIONAL

    tally = {"created": 0, "updated": 0, "cancelled": 0, "notified": 0, "failed": 0}

    for change in plan.changes:
        if change.kind == UNCHANGED:
            continue
        try:
            if change.kind == ORPHANED:
                item = change.item
                if change.detail.startswith("meeting") and send:
                    item.MeetingStatus = OL_MEETING_CANCELED
                    item.Save()
                    # Outlook can silently ignore the status change; Send() would then
                    # re-send the invitation. Skip it and leave the meeting in place.
                    if getattr(item, "MeetingStatus", None) != OL_MEETING_CANCELED:
                        tally["failed"] += 1
                        continue
                    item.Send()
                    tally["notified"] += change.people
                item.Delete()
                tally["cancelled"] += 1

            elif change.kind == NEW:
                item = folder.Items.Add(OL_APPOINTMENT)
                _write(item, change.event)
                _stamp_id(item, change.event)
                if change.event.required or change.event.optional:
                    item.MeetingStatus = OL_MEETING
                    item.ResponseRequested = response_requested
                    for who in change.event.required:
                        item.Recipients.Add(who).Type = OL_REQUIRED
                    for who in change.event.optional:
                        item.Recipients.Add(who).Type = OL_OPTIONAL
                    item.Recipients.ResolveAll()
                if send and change.notifies:
                    item.Send()
                    tally["notified"] += change.people
                else:
                    item.Save()
                tally["created"] += 1

            else:  # UPDATED
                item = change.item
                _write(item, change.event)
                if "attendees" in change.fields:
                    item.MeetingStatus = OL_MEETING
                    existing = _attendees_of(item)
                    for who in change.event.required:
                        if who.strip().lower() not in existing:
                            item.Recipients.Add(who).Type = OL_REQUIRED
                    for who in change.event.optional:
                        if who.strip().lower() not in existing:
                            item.Recipients.Add(who).Type = OL_OPTIONAL
                    item.Recipients.ResolveAll()
                if send and change.notifies:
                    item.Send()
                    tally["notified"] += change.people
                else:
                    item.Save()
                tally["updated"] += 1
        except Exception as exc:  # noqa: BLE001
            tally["failed"] += 1
            change.detail = str(exc)[:100]

        if on_progress:
            on_progress(change)

    return tally


def _write(item, event: Event) -> None:
    item.Subject = event.title
    item.Location = event.location
    item.Body = event.description
    if event.all_day:
        item.AllDayEvent = True
        item.Start = event.start.strftime("%Y-%m-%d")
        # inclusive sheet end -> Outlook's exclusive End (see outlook.push)
        item.End = (event.end + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        item.Start = event.start.strftime("%Y-%m-%d %H:%M")
        item.End = event.end.strftime("%Y-%m-%d %H:%M")
    if event.categories:
        item.Categories = ", ".join(event.categories)


def _stamp_id(item, event: Event) -> None:
    try:
        item.UserProperties.Add(PROP_UID, 1).Value = event.event_id
    except Exception:
        pass
