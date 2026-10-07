"""Keeping a description's formatting on its way to Outlook.

The planning sheet's Description cells carry Excel formatting - bold headings,
the odd italic or coloured word - and links typed as plain URLs. A plain
DESCRIPTION throws all of it away, so the .ics also gets an HTML copy in
X-ALT-DESC, which classic Outlook shows instead (verified 2026-10-07: bold,
bullets, links and tables all survive File > Import).
"""

from __future__ import annotations

import html
import re

URL = re.compile(r"https?://[^\s<>\"]+")
_URL_TRAILING = ".,;:!?)]}'"

# One run of text and its formatting: {"b", "i", "u", "s"} flags plus "color".
Run = tuple[str, dict]


class Formatted(str):
    """A cell's text that also remembers its formatted runs.

    A str, so every other part of the reader treats it like any other cell.
    """

    runs: list[Run]

    def __new__(cls, runs: list[Run]):
        self = super().__new__(cls, "".join(text for text, _ in runs))
        self.runs = runs
        return self


def from_rich_cell(value) -> Formatted:
    """An openpyxl CellRichText as Formatted text."""
    runs: list[Run] = []
    for part in value:
        if isinstance(part, str):
            runs.append((part, {}))
        else:
            runs.append((part.text, _style(part.font)))
    return Formatted(runs)


def _style(font) -> dict:
    if font is None:
        return {}
    style = {
        "b": bool(font.b),
        "i": bool(font.i),
        "u": bool(font.u) and font.u != "none",
        "s": bool(font.strike),
    }
    color = getattr(font, "color", None)
    rgb = getattr(color, "rgb", None) if getattr(color, "type", None) == "rgb" else None
    # Black is left to the reader's theme, so dark mode still shows it.
    if isinstance(rgb, str) and len(rgb) == 8 and rgb[2:].upper() != "000000":
        style["color"] = "#" + rgb[2:]
    return {key: value for key, value in style.items() if value}


def runs_of(value) -> list[Run]:
    """The runs of a cell, formatted or not."""
    return list(value.runs) if isinstance(value, Formatted) else [(str(value), {})]


def strip_runs(runs: list[Run]) -> list[Run]:
    """Trim whitespace at both ends, the way the plain text is trimmed."""
    runs = [(text.replace("\r\n", "\n").replace("\r", "\n"), style) for text, style in runs]
    while runs and not runs[0][0].strip():
        runs.pop(0)
    while runs and not runs[-1][0].strip():
        runs.pop()
    if runs:
        runs[0] = (runs[0][0].lstrip(), runs[0][1])
        runs[-1] = (runs[-1][0].rstrip(), runs[-1][1])
    return runs


def _text_html(text: str) -> str:
    out, last = [], 0
    for match in URL.finditer(text):
        url = match.group().rstrip(_URL_TRAILING)
        end = match.start() + len(url)
        out.append(html.escape(text[last:match.start()], quote=False))
        out.append(f'<a href="{html.escape(url)}">{html.escape(url, quote=False)}</a>')
        last = end
    out.append(html.escape(text[last:], quote=False))
    body = "".join(out)
    # HTML collapses whitespace; Excel text relies on tabs and runs of spaces.
    body = body.replace("\t", "&nbsp;" * 4).replace("  ", " &nbsp;")
    return body.replace("\n", "<br>")


def runs_html(runs: list[Run]) -> str:
    parts = []
    for text, style in runs:
        piece = _text_html(text)
        if not piece:
            continue
        for flag, tag in (("b", "b"), ("i", "i"), ("u", "u"), ("s", "s")):
            if style.get(flag):
                piece = f"<{tag}>{piece}</{tag}>"
        if style.get("color"):
            piece = f'<span style="color:{style["color"]}">{piece}</span>'
        parts.append(piece)
    return "".join(parts)


def document(body: str) -> str:
    """Wrap body HTML the way Outlook expects an X-ALT-DESC."""
    return (
        '<html><body><div style="font-family:Aptos,Calibri,sans-serif;font-size:11pt">'
        f"{body}</div></body></html>"
    )
