"""Give every row a permanent identity, written back into the sheet.

Re-running is the whole point of the tool: a course edition is edited over
weeks, and each run has to know which spreadsheet row is which calendar
event. Nothing already in the sheet can answer that. Titles repeat across
days, dates are exactly what gets edited, and row numbers move the moment
somebody inserts a line.

So the tool adds one column, ``AutoCalendar ID``, and fills it once. After
that a row can be retitled, moved to another week, or dragged to a different
position and still be recognised as the same event.

This is the only place AutoCalendar writes to a file the user owns, so it is
deliberately careful: it writes to a temporary file in the same folder and
replaces the original only once that has succeeded, it never touches a cell
that already has an id, and it refuses to guess about formats it cannot
round-trip.
"""

from __future__ import annotations

import csv
import os
import uuid
from pathlib import Path

from .model import Event
from .reader import ALIASES, normalise

ID_HEADER = "AutoCalendar ID"

#: The header names that already mean "this is our id column".
ID_ALIASES = set(ALIASES["event_id"])


class SheetWriteError(RuntimeError):
    """The sheet could not be updated, and nothing was changed."""


def new_id() -> str:
    """Short, unique, and readable enough to compare by eye in a cell."""
    raw = uuid.uuid4().hex
    return f"{raw[:4]}-{raw[4:8]}-{raw[8:12]}"


def assign_ids(events: list[Event]) -> int:
    """Fill in an id for every event that has none. Returns how many."""
    added = 0
    for event in events:
        if not event.event_id:
            event.event_id = new_id()
            added += 1
    return added


def _find_id_column(header: list) -> int | None:
    for index, cell in enumerate(header):
        if normalise(cell) in ID_ALIASES:
            return index
    return None


def _atomic_replace(tmp: Path, target: Path) -> None:
    """Swap the new file in only after it is completely written."""
    os.replace(tmp, target)


def write_ids_csv(path: Path, header_row: int, events: list[Event]) -> None:
    by_row = {e.source_row: e.event_id for e in events if e.event_id}
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))

    header = rows[header_row - 1] if header_row - 1 < len(rows) else []
    column = _find_id_column(header)
    if column is None:
        column = len(header)
        header.append(ID_HEADER)
        rows[header_row - 1] = header

    for number, row in enumerate(rows, start=1):
        if number <= header_row:
            continue
        while len(row) <= column:
            row.append("")
        if not row[column].strip() and number in by_row:
            row[column] = by_row[number]

    tmp = path.with_suffix(path.suffix + ".autocalendar-tmp")
    with open(tmp, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(rows)
    _atomic_replace(tmp, path)


def write_ids_xlsx(path: Path, sheet: str | None, header_row: int, events: list[Event]) -> None:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise SheetWriteError("openpyxl is needed to write ids into an .xlsx") from exc

    by_row = {e.source_row: e.event_id for e in events if e.event_id}
    book = load_workbook(path)
    page = book[sheet] if sheet and sheet in book.sheetnames else book.worksheets[0]

    header = [cell.value for cell in page[header_row]]
    column = _find_id_column(header)
    if column is None:
        column = len(header)
        page.cell(row=header_row, column=column + 1, value=ID_HEADER)

    for row_number, event_id in by_row.items():
        cell = page.cell(row=row_number, column=column + 1)
        if cell.value in (None, ""):
            cell.value = event_id

    tmp = path.with_suffix(path.suffix + ".autocalendar-tmp")
    book.save(tmp)
    _atomic_replace(tmp, path)


def ensure_ids(
    path: Path, events: list[Event], *, sheet: str | None, header_row: int
) -> int:
    """Give every event an id and persist it back into the sheet.

    Returns the number of ids newly written. Zero means the sheet already
    knew every row, which is the normal case from the second run onwards.
    """
    added = assign_ids(events)
    if not added:
        return 0

    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            write_ids_csv(path, header_row, events)
        elif suffix in {".xlsx", ".xlsm"}:
            write_ids_xlsx(path, sheet, header_row, events)
        else:
            raise SheetWriteError(f"cannot write ids into a {suffix} file")
    except SheetWriteError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise SheetWriteError(f"could not update {path.name}: {exc}") from exc
    return added
