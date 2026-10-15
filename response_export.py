"""
title: Response Export
author: tovren
author_url: https://github.com/tovren/open-webui-plugins
version: 1.0.0
license: MIT
required_open_webui_version: 0.11.4
description: Adds four buttons below every assistant message - export as Word (.docx), export its tables as Excel (.xlsx), print / save as PDF, and open as an e-mail draft. Files are built in memory and handed straight to the browser; nothing is stored on the server.

Buttons
- Word (.docx) with headings, lists (one nesting level), tables, bold/italic/code/links, code blocks
  and quotes.
- Excel (.xlsx) with one sheet per Markdown table, a styled and frozen header row, auto filter,
  fitted column widths and real numbers (also German formats such as "1.234,56 EUR" or "12,5 %").
- PDF via the browser's print dialog (A4 layout, small footer), so no PDF library is needed.
- E-mail via a mailto link; answers too long for a mailto link are downloaded as an .eml draft
  (marked "X-Unsent", so Outlook opens it as a new editable message).
Reasoning and tool-call blocks are removed before exporting.

Security & privacy
- No network access, no shell, nothing is written to disk or stored on the server. Only the
  standard library plus python-docx and openpyxl (both bundled with Open WebUI) are used.
- The browser receives the generated file (base64), a mailto link or the print HTML, wrapped in
  a fixed JavaScript snippet run via the "execute" event. All dynamic values are JSON-encoded.
- The page's Content-Security-Policy must allow 'unsafe-eval' for the "execute" event to work.
- Raw HTML in the response is always escaped, links are limited to http, https, mailto and tel,
  and images are never loaded (they become links). The print page also carries its own strict
  CSP. Spreadsheet cells are written as text or numbers, never as formulas.
"""

import asyncio
import base64
import html
import io
import json
import logging
import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field, replace
from datetime import datetime
from email.message import EmailMessage
from typing import TYPE_CHECKING, Any, Literal
from urllib.parse import quote
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.text.paragraph import Paragraph as DocxParagraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

if TYPE_CHECKING:
    from docx.styles.style import ParagraphStyle

log = logging.getLogger(__name__)

EventCallback = Callable[[dict[str, Any]], Awaitable[Any]]

# ------------------------------------------------------------------------------------------ constants --

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
EML_MIME = "message/rfc822"

BODY_FONT = "Calibri"
BODY_FONT_SIZE = Pt(11)
CODE_FONT = "Consolas"
CODE_FONT_SIZE = Pt(9.5)
MAX_DOCX_HEADING_LEVEL = 4
ORDERED_LIST_STYLES = ("List Number", "List Number 2")
BULLET_LIST_STYLES = ("List Bullet", "List Bullet 2")
TABLE_ALIGNMENTS = {  # (starts with ":", ends with ":") -> alignment
    (True, True): "center",
    (False, True): "right",
    (True, False): "left",
}
DOCX_ALIGNMENT = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}

HEADER_FILL = "DDE4EE"
MIN_COLUMN_WIDTH = 8
MAX_COLUMN_WIDTH = 60
MAX_EXACT_DIGITS = 15  # Excel keeps 15 significant digits; longer numbers stay text
EURO_FORMAT = '#,##0.00 "€"'
DOLLAR_FORMAT = '"$"#,##0.00'
POUND_FORMAT = '"£"#,##0.00'
CURRENCY_FORMATS = {  # symbol or ISO code before or after the amount -> Excel number format
    "€": EURO_FORMAT,
    "EUR": EURO_FORMAT,
    "$": DOLLAR_FORMAT,
    "USD": DOLLAR_FORMAT,
    "£": POUND_FORMAT,
    "GBP": POUND_FORMAT,
}

MAX_FILENAME_CHARS = 60
SAFE_LINK_PREFIXES = ("http://", "https://", "mailto:", "tel:")
PRINT_FRAME_LIFETIME_MS = 60_000
PRINT_CSP = "default-src 'none'; style-src 'unsafe-inline'"
PRINT_CSS = f"""
@page {{ size: A4; margin: 2cm 2cm 2.2cm 2cm; }}
body {{ font-family: {BODY_FONT}, Carlito, Arial, sans-serif; font-size: {BODY_FONT_SIZE.pt:g}pt; line-height: 1.4; color: #000; }}
h1 {{ font-size: 18pt; margin: 0 0 10pt; }}
h2 {{ font-size: 14pt; margin: 14pt 0 6pt; }}
h3, h4, h5, h6 {{ font-size: 12pt; margin: 12pt 0 4pt; }}
table {{ border-collapse: collapse; margin: 8pt 0; width: auto; }}
th, td {{ border: 1px solid #888; padding: 4pt 6pt; vertical-align: top; }}
th {{ background: #{HEADER_FILL}; }}
tr {{ page-break-inside: avoid; }}
pre {{ background: #f4f4f4; padding: 6pt; white-space: pre-wrap; font-size: {CODE_FONT_SIZE.pt:g}pt; }}
code {{ font-family: {CODE_FONT}, monospace; }}
blockquote {{ border-left: 3pt solid #bbb; margin: 6pt 0; padding-left: 8pt; color: #333; }}
hr {{ border: 0; border-top: 1px solid #ccc; }}
.footer {{ margin-top: 18pt; padding-top: 4pt; border-top: 1px solid #ccc; font-size: 8.5pt; color: #666; }}
"""
EMAIL_HTML = (  # format template with a {body} placeholder
    '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'
    f'<div style="font-family:{BODY_FONT},Arial,sans-serif;font-size:{BODY_FONT_SIZE.pt:g}pt">{{body}}</div>'
    "</body></html>"
)

# Button labels are language-neutral on purpose: Open WebUI builds the button list without loading
# the valves, so translated labels would only switch after the first click.

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "empty_message": "This response is empty, so there is nothing to export.",
        "no_tables": "This response does not contain a table.",
        "eml_downloaded": (
            "This response is too long to hand over to your mail program directly, so an e-mail draft "
            "(.eml) was downloaded. Open the file to edit and send it - Outlook shows it as a new message."
        ),
        "export_failed": "The export failed. Please try again or contact your administrator.",
        "default_title": "AI response",
        "email_draft_name": "E-mail draft",
        "default_subject": "Draft",
        "sheet_name": "Table {n}",
        "print_footer": "Created with the AI assistant on {date}. Please check the content before use.",
        "date_format": "%Y-%m-%d",
    },
    "de": {
        "empty_message": "Diese Antwort ist leer, es gibt nichts zu exportieren.",
        "no_tables": "Diese Antwort enthält keine Tabelle.",
        "eml_downloaded": (
            "Die Antwort ist zu lang für die direkte Übergabe an das Mailprogramm. Deshalb wurde ein "
            "E-Mail-Entwurf (.eml) heruntergeladen – Datei öffnen, dann erscheint sie in Outlook als "
            "neue E-Mail."
        ),
        "export_failed": "Der Export ist fehlgeschlagen. Bitte erneut versuchen oder an die IT wenden.",
        "default_title": "KI-Antwort",
        "email_draft_name": "E-Mail-Entwurf",
        "default_subject": "Entwurf",
        "sheet_name": "Tabelle {n}",
        "print_footer": "Erstellt mit dem KI-Assistenten am {date} – Inhalte bitte prüfen.",
        "date_format": "%d.%m.%Y",
    },
}


def _svg_icon(paths: str) -> str:
    """Return a stroke-style 24x24 SVG icon as a base64 data URI."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        f'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">{paths}</svg>'
    )
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()


# action id -> (button label, icon); each id is handled by Action._export_<id>
ACTIONS = {
    "docx": (
        "Word (.docx)",
        _svg_icon(
        '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>'
        '<path d="M8.5 12l1.2 5 1.3-3.5 1.3 3.5 1.2-5"/>'
        ),
    ),
    "xlsx": (
        "Excel (.xlsx)",
        _svg_icon('<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 9.5h16M4 14.5h16M10 4v16"/>'),
    ),
    "pdf": (
        "PDF",
        _svg_icon(
            '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>'
        ),
    ),
    "email": (
        "E-Mail",
        _svg_icon('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3.5 6.5l8.5 6.5 8.5-6.5"/>'),
    ),
}

# ---------------------------------------------------------------------------------- message content --

HIDDEN_BLOCK_RE = re.compile(r"<(details|think)\b[^>]*>.*?</\1>", re.DOTALL | re.IGNORECASE)
XML_ILLEGAL_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def find_message(body: dict[str, Any]) -> dict[str, Any]:
    """Return the clicked message, falling back to the last assistant message."""
    messages = [m for m in body.get("messages") or [] if isinstance(m, dict)]
    clicked = next((m for m in messages if m.get("id") == body.get("id")), None)
    fallback = next((m for m in reversed(messages) if m.get("role") == "assistant"), None)
    return clicked or fallback or {}


def message_text(message: dict[str, Any]) -> str:
    """Return the text of a message whose content is a string or a list of content parts."""
    content = message.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return "\n\n".join(
            part.get("text") or ""
            for part in content
            if isinstance(part, dict) and part.get("type") == "text"
        )
    return ""


def clean_markdown(text: str) -> str:
    """Drop reasoning/tool-call blocks and characters that are invalid in XML (docx/xlsx)."""
    return XML_ILLEGAL_RE.sub("", HIDDEN_BLOCK_RE.sub("", text)).strip()


# ----------------------------------------------------------------------------------- inline markdown --

INLINE_RE = re.compile(
    r"\\(?P<escaped>[\\`*_{}\[\]()#+\-.!|~<>])"
    r"|`(?P<code>[^`]+)`"
    r"|(?P<br><br\s*/?>)"
    r"|!?\[(?P<label>[^\]]*)\]\((?P<url>(?:[^()\s]|\([^()\s]*\))+)(?:\s+\"[^\"]*\")?\)"
    r"|<(?P<autolink>(?:https?://|mailto:)[^>\s]+)>"
    r"|\*\*\*(?P<bold_italic>.+?)\*\*\*"
    r"|\*\*(?P<bold>(?:[^*]|\*[^*\s][^*]*\*|\*)+?)\*\*"  # may contain *italic*
    r"|(?<!\w)__(?P<bold_u>.+?)__(?!\w)"
    r"|~~(?P<strike>.+?)~~"
    r"|\*(?P<italic>[^*\s](?:[^*]*[^*\s])?)\*"
    r"|(?<!\w)_(?P<italic_u>[^_\s](?:[^_]*[^_\s])?)_(?!\w)",
    re.IGNORECASE | re.DOTALL,  # emphasis may span the lines of a paragraph
)
EMPHASIS_GROUPS: dict[str, dict[str, bool]] = {
    "bold_italic": {"bold": True, "italic": True},
    "bold": {"bold": True},
    "bold_u": {"bold": True},
    "strike": {"strike": True},
    "italic": {"italic": True},
    "italic_u": {"italic": True},
}


@dataclass(frozen=True)
class Run:
    """A piece of inline text with uniform formatting. Line breaks are kept as '\\n' in the text."""

    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    url: str | None = None


PLAIN = Run("")


def parse_inline(text: str, style: Run = PLAIN) -> list[Run]:
    """Split inline Markdown into formatted runs (emphasis may be nested)."""
    runs: list[Run] = []
    position = 0
    for match in INLINE_RE.finditer(text):
        if match.start() > position:
            runs.append(replace(style, text=text[position : match.start()]))
        runs.extend(_runs_for_match(match, style))
        position = match.end()
    if position < len(text):
        runs.append(replace(style, text=text[position:]))
    return runs


def _runs_for_match(match: re.Match[str], style: Run) -> list[Run]:
    kind = match.lastgroup or ""
    if kind == "escaped":
        return [replace(style, text=match["escaped"])]
    if kind == "code":
        return [replace(style, text=match["code"], code=True)]
    if kind == "br":
        return [replace(style, text="\n")]
    if kind == "url":
        label = to_plain(match["label"]) or match["url"]
        return [replace(style, text=label, url=match["url"])]
    if kind == "autolink":
        return [replace(style, text=match["autolink"], url=match["autolink"])]
    return parse_inline(match[kind], replace(style, **EMPHASIS_GROUPS[kind]))


def _run_text(run: Run, with_url: bool) -> str:
    if with_url and run.url and run.url != run.text:
        return f"{run.text} ({run.url})"
    return run.text


def to_plain(text: str, with_urls: bool = False) -> str:
    """Strip inline Markdown; optionally keep link targets as 'label (url)'."""
    return "".join(_run_text(run, with_urls) for run in parse_inline(text)).strip()


def is_safe_url(url: str) -> bool:
    return url.strip().lower().startswith(SAFE_LINK_PREFIXES)


# ------------------------------------------------------------------------------------ block markdown --

FENCE_RE = re.compile(r"^\s*(?P<fence>`{3,}|~{3,})")
HEADING_RE = re.compile(r"^\s{0,3}(?P<hashes>#{1,6})\s+(?P<text>.*?)(?:\s+#+)?\s*$")
RULE_RE = re.compile(r"^\s{0,3}([-*_])(?:\s*\1){2,}\s*$")
LIST_ITEM_RE = re.compile(r"^(?P<indent>[ \t]*)(?:[-*+]|(?P<number>\d{1,9})[.)])\s+(?P<text>.*)$")
QUOTE_RE = re.compile(r"^\s{0,3}>\s?(?P<text>.*)$")
TABLE_DELIMITER_RE = re.compile(r"^:?-+:?$")
UNESCAPED_PIPE_RE = re.compile(r"(?<!\\)\|")
UNSAFE_FILENAME_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f\x7f]+')


@dataclass
class Heading:
    level: int
    text: str


@dataclass
class Paragraph:
    text: str  # lines are stripped and joined with "\n"


@dataclass
class ListItem:
    text: str
    sublists: list["List"] = field(default_factory=list)  # nested lists; deeper nesting is flattened


@dataclass
class List:
    ordered: bool
    start: int
    items: list[ListItem]


@dataclass
class Quote:
    lines: list[str]


@dataclass
class CodeBlock:
    code: str


@dataclass
class Table:
    rows: list[list[str]]  # first row is the header; all rows have the header's width
    aligns: list[str | None]


@dataclass
class Rule:
    pass


Block = Heading | Paragraph | List | Quote | CodeBlock | Table | Rule


def _split_row(line: str) -> list[str]:
    line = line.strip().removeprefix("|")
    if line.endswith("|") and not line.endswith("\\|"):
        line = line[:-1]
    # GFM unescapes "\|" in cells before inline parsing, so it also works inside `code`
    return [cell.strip().replace("\\|", "|") for cell in UNESCAPED_PIPE_RE.split(line)]


def _table_alignments(lines: list[str], i: int) -> list[str | None] | None:
    """Return column alignments if a GFM table starts at line i, else None."""
    if "|" not in lines[i] or i + 1 >= len(lines) or "|" not in lines[i + 1]:
        return None
    header, delimiter = _split_row(lines[i]), [c.replace(" ", "") for c in _split_row(lines[i + 1])]
    if len(header) != len(delimiter) or not all(TABLE_DELIMITER_RE.match(c) for c in delimiter):
        return None
    return [TABLE_ALIGNMENTS.get((cell.startswith(":"), cell.endswith(":"))) for cell in delimiter]


def _parse_table(lines: list[str], i: int, aligns: list[str | None]) -> tuple[Table, int]:
    width = len(aligns)
    rows = [_split_row(lines[i])]
    i += 2
    while i < len(lines) and lines[i].strip() and "|" in lines[i]:
        rows.append(_split_row(lines[i]))
        i += 1
    # GFM: missing cells are empty, surplus cells are ignored
    return Table([(row + [""] * width)[:width] for row in rows], aligns), i


def _parse_code(lines: list[str], i: int, fence: str) -> tuple[CodeBlock, int]:
    closing = re.compile(rf"^\s*{re.escape(fence[0])}{{{len(fence)},}}\s*$")
    code: list[str] = []
    i += 1
    while i < len(lines) and not closing.match(lines[i]):
        code.append(lines[i])
        i += 1
    return CodeBlock("\n".join(code)), i + 1


def _starts_block(lines: list[str], i: int) -> bool:
    """True if line i is blank or opens a block, i.e. ends a running paragraph."""
    line = lines[i]
    return bool(
        not line.strip()
        or FENCE_RE.match(line)
        or HEADING_RE.match(line)
        or RULE_RE.match(line)
        or LIST_ITEM_RE.match(line)
        or QUOTE_RE.match(line)
        or _table_alignments(lines, i) is not None
    )


def _list_line(lines: list[str], i: int) -> re.Match[str] | None:
    """Match line i as a list item, unless a block that takes precedence starts there."""
    line = lines[i]
    if FENCE_RE.match(line) or HEADING_RE.match(line) or RULE_RE.match(line):
        return None
    if _table_alignments(lines, i) is not None:
        return None
    return LIST_ITEM_RE.match(line)


def _add_list_item(lists: list[List], match: re.Match[str], text: str) -> None:
    ordered = match["number"] is not None
    nested = bool(lists) and len(match["indent"].expandtabs(4)) >= 2
    siblings = lists[-1].items[-1].sublists if nested else lists
    if not siblings or siblings[-1].ordered != ordered:  # a change of list type starts a new list
        siblings.append(List(ordered, int(match["number"] or 1), []))
    siblings[-1].items.append(ListItem(text))


def _parse_lists(lines: list[str], i: int, match: re.Match[str]) -> tuple[list[List], int]:
    """Parse consecutive list items (blank lines between them allowed) into top-level lists."""
    lists: list[List] = []
    while True:
        text = match["text"].strip()
        i += 1
        # indented lines directly below an item continue its text
        while i < len(lines) and lines[i][:1].isspace() and not _starts_block(lines, i):
            text += " " + lines[i].strip()
            i += 1
        _add_list_item(lists, match, text)
        while i < len(lines) and not lines[i].strip():
            i += 1
        if i >= len(lines) or not (match := _list_line(lines, i)):
            return lists, i


def parse_blocks(markdown_text: str) -> list[Block]:
    """Parse the subset of Markdown that chat models produce into a flat list of blocks."""
    lines = markdown_text.splitlines()
    blocks: list[Block] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
        elif fence := FENCE_RE.match(line):
            code, i = _parse_code(lines, i, fence["fence"])
            blocks.append(code)
        elif (aligns := _table_alignments(lines, i)) is not None:
            table, i = _parse_table(lines, i, aligns)
            blocks.append(table)
        elif heading := HEADING_RE.match(line):
            if heading["text"]:
                blocks.append(Heading(len(heading["hashes"]), heading["text"]))
            i += 1
        elif RULE_RE.match(line):
            blocks.append(Rule())
            i += 1
        elif item_match := LIST_ITEM_RE.match(line):
            lists, i = _parse_lists(lines, i, item_match)
            blocks.extend(lists)
        elif QUOTE_RE.match(line):
            quote = Quote([])
            while i < len(lines) and (quote_match := QUOTE_RE.match(lines[i])):
                quote.lines.append(quote_match["text"])
                i += 1
            blocks.append(quote)
        else:
            paragraph = [line.strip()]
            i += 1
            while i < len(lines) and not _starts_block(lines, i):
                paragraph.append(lines[i].strip())
                i += 1
            blocks.append(Paragraph("\n".join(paragraph)))
    return blocks


def document_title(blocks: list[Block]) -> str:
    """Plain text of the first heading, or '' if there is none."""
    return next((to_plain(b.text) for b in blocks if isinstance(b, Heading)), "")


def safe_filename(text: str, fallback: str) -> str:
    """Turn text into a short file name that is valid on Windows, macOS and Linux."""
    name = " ".join(UNSAFE_FILENAME_RE.sub(" ", text).split())
    return name[:MAX_FILENAME_CHARS].strip(" .") or fallback


# ------------------------------------------------------------------------------------------------ Word --


def _add_runs(paragraph: DocxParagraph, runs: list[Run]) -> None:
    for run in runs:
        docx_run = paragraph.add_run(_run_text(run, with_url=True))
        if run.bold:
            docx_run.bold = True
        if run.italic:
            docx_run.italic = True
        if run.strike:
            docx_run.font.strike = True
        if run.code:
            docx_run.font.name = CODE_FONT


class DocxWriter:
    """Render parsed blocks into a Word document."""

    def __init__(self) -> None:
        self.doc = Document()
        normal_font = self.doc.styles["Normal"].font
        normal_font.name = BODY_FONT
        normal_font.size = BODY_FONT_SIZE

    def render(self, blocks: list[Block]) -> bytes:
        for block in blocks:
            match block:
                case Heading(level, text):
                    self.doc.add_heading(to_plain(text), level=min(level, MAX_DOCX_HEADING_LEVEL))
                case Paragraph(text):
                    _add_runs(self.doc.add_paragraph(), parse_inline(text))
                case List():
                    self._list(block, level=0)
                case Quote(lines):
                    for line in filter(str.strip, lines):
                        _add_runs(self.doc.add_paragraph(style="Quote"), parse_inline(line))
                case CodeBlock(code):
                    run = self.doc.add_paragraph().add_run(code)
                    run.font.name = CODE_FONT
                    run.font.size = CODE_FONT_SIZE
                case Table():
                    self._table(block)
        buffer = io.BytesIO()
        self.doc.save(buffer)
        return buffer.getvalue()

    def _list(self, block: List, level: int) -> None:
        style = (ORDERED_LIST_STYLES if block.ordered else BULLET_LIST_STYLES)[level]
        num_id: int | None = None
        for item in block.items:
            paragraph = self.doc.add_paragraph(style=style)
            if block.ordered:
                if num_id is None:  # each list gets its own numbering, so the next one restarts
                    num_id = self._new_numbering(paragraph.style, block.start)
                paragraph._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num_id
            _add_runs(paragraph, parse_inline(item.text))
            for sublist in item.sublists:
                self._list(sublist, level=1)

    def _new_numbering(self, style: "ParagraphStyle", start: int) -> int:
        """Create a numbering instance of the style's list that starts at `start`.

        python-docx has no public API for this; without it every numbered list in the
        document would continue the previous list's numbering.
        """
        numbering = self.doc.part.numbering_part.element
        style_num_id = style.element.pPr.numPr.numId.val
        num = numbering.add_num(numbering.num_having_numId(style_num_id).abstractNumId.val)
        num.add_lvlOverride(ilvl=0).add_startOverride(start)
        return num.numId

    def _table(self, table: Table) -> None:
        doc_table = self.doc.add_table(rows=len(table.rows), cols=len(table.aligns))
        doc_table.style = "Table Grid"
        for row_index, (doc_row, row) in enumerate(zip(doc_table.rows, table.rows, strict=True)):
            for cell, text, align in zip(doc_row.cells, row, table.aligns, strict=True):
                paragraph = cell.paragraphs[0]
                paragraph.alignment = DOCX_ALIGNMENT.get(align or "")
                _add_runs(paragraph, parse_inline(text, Run("", bold=row_index == 0)))
        self.doc.add_paragraph()


def markdown_to_docx(blocks: list[Block]) -> bytes:
    return DocxWriter().render(blocks)


# ----------------------------------------------------------------------------------------------- Excel --

CURRENCY_RE = "|".join(map(re.escape, CURRENCY_FORMATS))
NUMBER_RE = re.compile(  # thousands separators: . , ' space, no-break space, narrow no-break space
    rf"(?P<prefix>{CURRENCY_RE})?\s*"
    r"(?P<sign>[-+]?)"
    r"(?P<int>\d{1,3}(?:(?P<group>[.,' \u00a0\u202f])\d{3}(?:(?P=group)\d{3})*)?|\d+)"
    r"(?:(?!(?P=group))(?P<decimal>[.,])(?P<fraction>\d+))?"  # decimal mark differs from the separator
    rf"\s*(?P<suffix>%|{CURRENCY_RE})?"
)


def parse_number(text: str) -> tuple[str | int | float, str | None]:
    """Convert '1.234,56 €', '1.234,56 EUR', '-3,2 %', '1,234.5' etc. to (number, Excel number format).

    A dot or comma followed by exactly three digits is read as a thousands separator, so
    "2.000" is 2000 and "1,234" is 1234. Values that do not look like a plain number, and
    integers with leading zeros (postcodes, IDs), are returned unchanged as (text, None).
    """
    match = NUMBER_RE.fullmatch(text.strip().replace("\u2212", "-"))  # Unicode minus sign
    if not match or (match["prefix"] and match["suffix"]):
        return text, None
    digits, fraction = re.sub(r"\D", "", match["int"]), match["fraction"] or ""
    if (len(digits) > 1 and digits.startswith("0")) or len(digits) + len(fraction) > MAX_EXACT_DIGITS:
        return text, None
    number: int | float = (
        float(f"{match['sign']}{digits}.{fraction}") if fraction else int(match["sign"] + digits)
    )
    unit = match["prefix"] or match["suffix"]
    if unit == "%":
        return number / 100, "0%" if float(number).is_integer() else "0.0#%"
    if unit:
        return float(number), CURRENCY_FORMATS[unit]
    return number, None


def tables_to_xlsx(tables: list[Table], sheet_name: str) -> bytes:
    """One worksheet per table; `sheet_name` is a format string with an {n} placeholder."""
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_fill = PatternFill("solid", fgColor=HEADER_FILL)
    for n, table in enumerate(tables, start=1):
        sheet = workbook.create_sheet(title=sheet_name.format(n=n))
        widths = [0] * len(table.aligns)
        for row_index, row in enumerate(table.rows, start=1):
            for column, raw in enumerate(row, start=1):
                text = to_plain(raw)
                value, number_format = (text, None) if row_index == 1 else parse_number(text)
                cell = sheet.cell(row=row_index, column=column, value=value if value != "" else None)
                if isinstance(value, str) and value:
                    cell.data_type = "s"  # never let model output become a live formula ("=...")
                if number_format:
                    cell.number_format = number_format
                if "\n" in text:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                if row_index == 1:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                widths[column - 1] = max(widths[column - 1], *(len(line) for line in text.split("\n")))
        for column, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(column)].width = min(
                max(width + 2, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH
            )
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


# ------------------------------------------------------------------------------------ HTML and text --


def _runs_html(runs: list[Run]) -> str:
    parts: list[str] = []
    for run in runs:
        part = html.escape(run.text).replace("\n", "<br>")
        if run.code:
            part = f"<code>{part}</code>"
        if run.strike:
            part = f"<s>{part}</s>"
        if run.italic:
            part = f"<em>{part}</em>"
        if run.bold:
            part = f"<strong>{part}</strong>"
        if run.url and is_safe_url(run.url):
            part = f'<a href="{html.escape(run.url.strip())}">{part}</a>'
        parts.append(part)
    return "".join(parts)


def _inline_html(text: str) -> str:
    return _runs_html(parse_inline(text))


def _list_html(block: List) -> str:
    tag = "ol" if block.ordered else "ul"
    start = f' start="{block.start}"' if block.ordered and block.start != 1 else ""
    items = "".join(
        f"<li>{_inline_html(item.text)}{''.join(map(_list_html, item.sublists))}</li>" for item in block.items
    )
    return f"<{tag}{start}>{items}</{tag}>"


def _table_html(table: Table) -> str:
    def row_html(row: list[str], tag: str) -> str:
        cells = []
        for text, align in zip(row, table.aligns, strict=True):
            style = f' style="text-align:{align}"' if align else ""
            cells.append(f"<{tag}{style}>{_inline_html(text)}</{tag}>")
        return f"<tr>{''.join(cells)}</tr>"

    body = "".join(row_html(row, "td") for row in table.rows[1:])
    return (
        '<table border="1" cellpadding="5" style="border-collapse:collapse">'
        f"<thead>{row_html(table.rows[0], 'th')}</thead><tbody>{body}</tbody></table>"
    )


def _block_html(block: Block) -> str:
    match block:
        case Heading(level, text):
            return f"<h{level}>{_inline_html(text)}</h{level}>"
        case Paragraph(text):
            return f"<p>{_inline_html(text)}</p>"
        case List():
            return _list_html(block)
        case Quote(lines):
            return f"<blockquote>{'<br>'.join(_inline_html(line) for line in lines)}</blockquote>"
        case CodeBlock(code):
            return f"<pre><code>{html.escape(code)}</code></pre>"
        case Table():
            return _table_html(block)
        case Rule():
            return "<hr>"
    raise TypeError(f"Unexpected block: {block!r}")


def blocks_to_html(blocks: list[Block]) -> str:
    """Render blocks as HTML. Every piece of text is escaped; no raw HTML passes through."""
    return "\n".join(map(_block_html, blocks))


def _list_text(block: List, indent: str = "") -> str:
    lines: list[str] = []
    for number, item in enumerate(block.items, start=block.start):
        marker = f"{number}." if block.ordered else "•"
        lines.append(f"{indent}{marker} {to_plain(item.text, with_urls=True)}")
        lines.extend(_list_text(sublist, indent + "    ") for sublist in item.sublists)
    return "\n".join(lines)


def _block_text(block: Block) -> str:
    match block:
        case Heading(_, text):
            return to_plain(text, with_urls=True)
        case Paragraph(text):
            return to_plain(text, with_urls=True)
        case List():
            return _list_text(block)
        case Quote(lines):
            return "\n".join(f"> {to_plain(line, with_urls=True)}" for line in lines)
        case CodeBlock(code):
            return code
        case Table(rows):
            return "\n".join(
                "   ".join(to_plain(cell, with_urls=True) for cell in row).rstrip() for row in rows
            )
    raise TypeError(f"Unexpected block: {block!r}")


def blocks_to_text(blocks: list[Block]) -> str:
    """Render blocks as readable plain text (used for mailto links and the .eml text part)."""
    return "\n\n".join(_block_text(block) for block in blocks if not isinstance(block, Rule))


# ---------------------------------------------------------------------------------------------- E-mail --

SUBJECT_RE = re.compile(r"^(?:subject|betreff)\s*:\s*(?P<subject>.+)$", re.IGNORECASE)


def split_email(blocks: list[Block]) -> tuple[str, list[Block]]:
    """Return (subject, body blocks). A 'Subject:'/'Betreff:' line in a heading or paragraph becomes
    the subject and everything before it (the model's preamble) is dropped; without such a line
    the subject is '' and all blocks are the body."""
    for i, block in enumerate(blocks):
        if not isinstance(block, Heading | Paragraph):
            continue
        lines = block.text.split("\n")
        for j, line in enumerate(lines):
            if match := SUBJECT_RE.match(to_plain(line)):
                rest = "\n".join(lines[j + 1 :])
                return match["subject"].strip(), ([Paragraph(rest)] if rest else []) + blocks[i + 1 :]
    return "", blocks


def build_mailto(subject: str, body_text: str) -> str:
    body = body_text.replace("\r\n", "\n").replace("\n", "\r\n")
    return f"mailto:?subject={quote(subject, safe='')}&body={quote(body, safe='')}"


def build_eml(subject: str, blocks: list[Block], body_text: str) -> bytes:
    """Build an unsent e-mail draft with a plain-text and an HTML alternative."""
    message = EmailMessage()
    message["Subject"] = " ".join(subject.split())
    message["X-Unsent"] = "1"  # Outlook opens the file as a new, editable message
    message.set_content(body_text)
    message.add_alternative(EMAIL_HTML.format(body=blocks_to_html(blocks)), subtype="html")
    return message.as_bytes()


# ------------------------------------------------------------------------------------------------- PDF --


def build_print_html(blocks: list[Block], title: str, footer: str, language: str) -> str:
    """Standalone A4 print page; the browser uses <title> as the suggested PDF file name."""
    footer_html = f'<div class="footer">{html.escape(footer)}</div>' if footer else ""
    return (
        f'<!DOCTYPE html><html lang="{html.escape(language)}"><head><meta charset="utf-8">'
        f'<meta http-equiv="Content-Security-Policy" content="{PRINT_CSP}">'
        f"<title>{html.escape(title)}</title><style>{PRINT_CSS}</style></head>"
        f"<body>{blocks_to_html(blocks)}{footer_html}</body></html>"
    )


# ------------------------------------------------------------------------------------- browser snippets --
# The snippets run inside an async function in the user's browser. They must return quickly:
# the server waits for the result with a timeout, hence print() is scheduled with setTimeout.


def _click_link_js(href: str, filename: str | None = None) -> str:
    """Click a temporary link; `href` is a JavaScript expression, `filename` turns it into a download."""
    download = f"\nlink.download = {json.dumps(filename)};" if filename is not None else ""
    return f"""
const link = document.createElement('a');
link.href = {href};{download}
document.body.appendChild(link);
link.click();
link.remove();
"""


def download_js(data: bytes, filename: str, mime: str) -> str:
    return f"""
const bytes = Uint8Array.from(atob({json.dumps(base64.b64encode(data).decode())}), (c) => c.charCodeAt(0));
const url = URL.createObjectURL(new Blob([bytes], {{ type: {json.dumps(mime)} }}));
{_click_link_js("url", filename)}
setTimeout(() => URL.revokeObjectURL(url), 10000);
return true;
"""


def open_url_js(url: str) -> str:
    return f"""{_click_link_js(json.dumps(url))}
return true;
"""


def print_js(page_html: str) -> str:
    return f"""
const frame = document.createElement('iframe');
frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
document.body.appendChild(frame);
frame.onload = () => setTimeout(() => {{
  frame.contentWindow.focus();
  frame.contentWindow.print();
  setTimeout(() => frame.remove(), {PRINT_FRAME_LIFETIME_MS});
}}, 100);
frame.srcdoc = {json.dumps(page_html)};
return true;
"""


# ---------------------------------------------------------------------------------------------- Action --


@dataclass
class Browser:
    """Thin wrapper around Open WebUI's event callbacks."""

    emitter: EventCallback | None
    caller: EventCallback | None

    async def notify(self, level: Literal["info", "success", "warning", "error"], text: str) -> None:
        if self.emitter is not None:
            await self.emitter({"type": "notification", "data": {"type": level, "content": text}})

    async def run_js(self, code: str) -> None:
        if self.caller is None:
            raise RuntimeError("__event_call__ is not available, cannot run code in the browser")
        await self.caller({"type": "execute", "data": {"code": code}})

    async def download(self, data: bytes, filename: str, mime: str) -> None:
        await self.run_js(download_js(data, filename, mime))


class Action:
    class Valves(BaseModel):
        language: Literal["en", "de"] = Field(
            default="en",
            description="Language of the buttons, messages, default file names and the print footer.",
        )
        timezone: str = Field(
            default="",
            description="IANA time zone for file name timestamps and the print footer date, "
            "e.g. 'Europe/Berlin'. Empty = server time.",
        )
        max_mailto_length: int = Field(
            default=1900,
            ge=0,
            description="Longest mailto link to hand to the mail program (Outlook on Windows handles about "
            "2000 characters). Longer answers are downloaded as an .eml draft instead.",
        )
        always_eml: bool = Field(
            default=False,
            description="Always download a formatted .eml draft instead of opening a plain-text mailto link.",
        )
        print_footer: bool = Field(
            default=True,
            description="Add a small 'created with the AI assistant' note with the date to printed pages.",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()

    @property
    def actions(self) -> list[dict[str, str]]:
        """Sub-actions shown as buttons (language-neutral labels, see ACTIONS)."""
        return [
            {"id": action_id, "name": label, "icon_url": icon} for action_id, (label, icon) in ACTIONS.items()
        ]

    async def action(
        self,
        body: dict[str, Any],
        __id__: str | None = None,
        __event_emitter__: EventCallback | None = None,
        __event_call__: EventCallback | None = None,
    ) -> None:
        texts = TRANSLATIONS[self.valves.language]
        browser = Browser(__event_emitter__, __event_call__)
        action_id = __id__ or "docx"
        if action_id not in ACTIONS:
            log.warning("Unknown response export action %r", action_id)
            return
        exporter = getattr(self, f"_export_{action_id}")
        content = clean_markdown(message_text(find_message(body)))
        if not content:
            await browser.notify("warning", texts["empty_message"])
            return
        try:
            await exporter(content, texts, browser)
        except Exception:
            log.exception("Response export %r failed", action_id)
            await browser.notify("error", texts["export_failed"])

    async def _export_docx(self, content: str, texts: dict[str, str], browser: Browser) -> None:
        blocks = parse_blocks(content)
        data = await asyncio.to_thread(markdown_to_docx, blocks)
        await browser.download(data, f"{self._file_stem(blocks, texts)}.docx", DOCX_MIME)

    async def _export_xlsx(self, content: str, texts: dict[str, str], browser: Browser) -> None:
        blocks = parse_blocks(content)
        tables = [block for block in blocks if isinstance(block, Table)]
        if not tables:
            await browser.notify("warning", texts["no_tables"])
            return
        data = await asyncio.to_thread(tables_to_xlsx, tables, texts["sheet_name"])
        await browser.download(data, f"{self._file_stem(blocks, texts)}.xlsx", XLSX_MIME)

    async def _export_pdf(self, content: str, texts: dict[str, str], browser: Browser) -> None:
        blocks = parse_blocks(content)
        footer = ""
        if self.valves.print_footer:
            footer = texts["print_footer"].format(date=self._now().strftime(texts["date_format"]))
        page = build_print_html(blocks, self._file_stem(blocks, texts), footer, self.valves.language)
        await browser.run_js(print_js(page))

    async def _export_email(self, content: str, texts: dict[str, str], browser: Browser) -> None:
        subject, blocks = split_email(parse_blocks(content))
        subject = subject or document_title(blocks)
        body_text = blocks_to_text(blocks)
        if not self.valves.always_eml:
            url = build_mailto(subject, body_text)
            if len(url) <= self.valves.max_mailto_length:
                await browser.run_js(open_url_js(url))
                return
        data = await asyncio.to_thread(build_eml, subject or texts["default_subject"], blocks, body_text)
        await browser.download(data, f"{safe_filename(subject, texts['email_draft_name'])}.eml", EML_MIME)
        if not self.valves.always_eml:
            await browser.notify("info", texts["eml_downloaded"])

    def _file_stem(self, blocks: list[Block], texts: dict[str, str]) -> str:
        """File name without extension: first heading (or a default) plus a timestamp."""
        return f"{safe_filename(document_title(blocks), texts['default_title'])} {self._now():%Y-%m-%d %H%M}"

    def _now(self) -> datetime:
        if timezone := self.valves.timezone.strip():
            try:
                return datetime.now(ZoneInfo(timezone))
            except (ZoneInfoNotFoundError, ValueError):
                log.warning("Unknown time zone %r, using server time", self.valves.timezone)
        return datetime.now()
