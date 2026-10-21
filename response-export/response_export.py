"""
title: Response Export
author: tovren
author_url: https://github.com/tovren/open-webui-plugins
version: 1.1.0
license: MIT
required_open_webui_version: 0.11.4
description: Adds four buttons below every assistant message - export as Word (.docx), export its tables as Excel (.xlsx), print / save as PDF, and open as an e-mail draft. Files are built in memory and handed straight to the browser; nothing is stored on the server.

Buttons
- Word (.docx) with headings, lists (one nesting level), tables, bold/italic/code, clickable links,
  code blocks, quotes and horizontal rules, in the design of Office Documents (accent_color and
  font valves) and optionally on a company template such as a letterhead (word_template valve).
- Excel (.xlsx) with one sheet per Markdown table, a styled and frozen header row, auto filter,
  fitted column widths and real numbers (also German and Swiss formats such as "1.234,56 EUR",
  "12,5 %" or "CHF 1'234.50"); phone numbers such as "+49 711 123456" stay text.
- PDF via the browser's print dialog (A4 layout, small footer), so no PDF library is needed.
- E-mail via a mailto link; answers too long for a mailto link are downloaded as an .eml draft
  (marked "X-Unsent", so Outlook opens it as a new editable message).
Reasoning and tool-call blocks are removed before exporting.

Security & privacy
- No network access, no shell, nothing is written to disk or stored on the server. Only the
  standard library plus python-docx and openpyxl (both bundled with Open WebUI) are used.
- The browser receives the generated file (base64), a mailto link or the print HTML, wrapped in
  a fixed JavaScript snippet run via the "execute" event. All dynamic values are JSON-encoded.
- With language 'auto' or an empty timezone, a fixed snippet first reads the interface language
  and time zone of the browser; nothing else is read from the page.
- The page's Content-Security-Policy must allow 'unsafe-eval' for the "execute" event to work. If
  the browser does not run the snippet (or does not answer within 30 seconds), an error is shown.
- The word_template valve reads one Word file from the server (at most 20 MB); its text is removed,
  only header, footer, page size and margins are kept.
- Raw HTML in the response is always escaped, links are limited to http, https, mailto and tel
  (in Word to http, https and mailto), and images are never loaded (they become links). The
  print page also carries its own strict CSP. Spreadsheet cells are written as text or numbers,
  never as formulas.
"""

import asyncio
import base64
import copy
import html
import io
import json
import logging
import re
import zipfile
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field, replace
from datetime import datetime, tzinfo
from email.message import EmailMessage
from typing import TYPE_CHECKING, Any, Literal
from urllib.parse import quote
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import CONTENT_TYPE as DOCX_CT
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.parts.numbering import NumberingPart
from docx.shared import Cm, Pt, RGBColor
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

BODY_FONT = "Calibri"  # print page and e-mail; Word uses the font valve
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
FRANC_FORMAT = '"CHF" #,##0.00'
CURRENCY_FORMATS = {  # symbol or ISO code before or after the amount -> Excel number format
    "€": EURO_FORMAT,
    "EUR": EURO_FORMAT,
    "$": DOLLAR_FORMAT,
    "USD": DOLLAR_FORMAT,
    "£": POUND_FORMAT,
    "GBP": POUND_FORMAT,
    "CHF": FRANC_FORMAT,
}

MAX_FILENAME_CHARS = 60
SAFE_LINK_PREFIXES = ("http://", "https://", "mailto:", "tel:")
WORD_LINK_PREFIXES = ("http://", "https://", "mailto:")  # as in Office Documents
MAILTO_FIELDS = {"subject", "body", "cc", "bcc"}  # others, e.g. "attach", are acted on by some mail clients
BROWSER_TIMEOUT_S = (
    30  # Open WebUI waits for the browser without a limit unless WEBSOCKET_EVENT_CALLER_TIMEOUT
)
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
        "browser_failed": (
            "Your browser did not run the export. Please reload the page and try again. If it keeps "
            "failing, contact your administrator: the page's Content Security Policy has to allow "
            "'unsafe-eval' for Open WebUI's code events."
        ),
        "default_title": "AI response",
        "email_draft_name": "E-mail draft",
        "default_subject": "Draft",
        "sheet_name": "Table {n}",
        "print_footer": "Created with the AI assistant on {date}. Please check the content before use.",
        "date_format": "%Y-%m-%d",
        "paper": "letter",
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
        "browser_failed": (
            "Der Browser hat den Export nicht ausgeführt. Bitte die Seite neu laden und erneut versuchen. "
            "Klappt es weiterhin nicht, bitte an die IT wenden: Die Content-Security-Policy der Seite muss "
            "'unsafe-eval' für die Code-Events von Open WebUI erlauben."
        ),
        "default_title": "KI-Antwort",
        "email_draft_name": "E-Mail-Entwurf",
        "default_subject": "Entwurf",
        "sheet_name": "Tabelle {n}",
        "print_footer": "Erstellt mit dem KI-Assistenten am {date} – Inhalte bitte prüfen.",
        "date_format": "%d.%m.%Y",
        "paper": "a4",
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

# innermost block first (no opening tag of the same kind inside), so nested <details> are removed completely
HIDDEN_BLOCK_RE = re.compile(r"<(details|think)\b[^>]*>(?:(?!<\1\b).)*?</\1>", re.DOTALL | re.IGNORECASE)
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
    """Drop reasoning/tool-call blocks (also nested ones) and characters invalid in XML (docx/xlsx)."""
    count = 1
    while count:
        text, count = HIDDEN_BLOCK_RE.subn("", text)
    return XML_ILLEGAL_RE.sub("", text).strip()


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


def is_word_link(url: str) -> bool:
    """Links Word gets as real hyperlinks (like Office Documents): http, https and plain mailto links."""
    url = url.strip()
    if not url.lower().startswith(WORD_LINK_PREFIXES):
        return False
    if url.lower().startswith("mailto:") and "?" in url:
        fields = [part.split("=", 1)[0].lower() for part in url.split("?", 1)[1].split("&") if part]
        return all(field in MAILTO_FIELDS for field in fields)
    return True


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
# Same design, valves (accent_color, font, word_template) and link rules as the Word files of Office
# Documents, copied from it, so both plugins give the same result.

DEFAULT_ACCENT = "1F4E79"
WHITE, BLACK = "FFFFFF", "000000"
TEXT_COLOR = "262626"
MUTED_COLOR = "6B6B6B"
RULE_COLOR = "D9D9D9"  # thin lines
FONT_NAME_RE = re.compile(r"[^\W_][\w .-]{0,40}")  # 'Calibri', 'Segoe UI', 'Source Sans 3'
THEME_FONTS_RE = (
    r'(<a:(?:major|minor)Font>\s*<a:latin typeface=")[^"]*(")'  # headings and body text of a theme
)


def mix(color: str, other: str, share: float) -> str:
    """`color` with `share` (0 to 1) of `other` mixed in, as RRGGBB."""
    return "".join(
        f"{round(a + (b - a) * share):02X}"
        for a, b in zip(bytes.fromhex(color), bytes.fromhex(other), strict=True)
    )


@dataclass(frozen=True)
class Design:
    """Accent colour (accent_color valve) and font (font valve) of the Word file."""

    accent: str = DEFAULT_ACCENT
    font: str = BODY_FONT

    @property
    def band(self) -> str:  # every second table row
        return mix(self.accent, WHITE, 0.93)


DEFAULT_DESIGN = Design()

# Word wants the children of w:pPr and w:tblPr in a fixed order; new ones go before these.
PPR_AFTER_BORDERS = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
    "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
    "w:rPr", "w:sectPr", "w:pPrChange",
)  # fmt: skip
TBLPR_AFTER_BORDERS = (
    "w:shd",
    "w:tblLayout",
    "w:tblCellMar",
    "w:tblLook",
    "w:tblCaption",
    "w:tblDescription",
)
TBLPR_AFTER_CELL_MARGINS = ("w:tblLook", "w:tblCaption", "w:tblDescription")
THEME_FONT_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme")
PAPER_SIZES = {"a4": (Cm(21), Cm(29.7)), "letter": (Cm(21.59), Cm(27.94))}

MAX_TEMPLATE_BYTES = 20_000_000
# With a template, its styles for these are replaced: they get the design. The content uses Body Text, not the
# template's Normal style, which its header and footer build on.
LIST_STYLES = ("List Bullet", "List Bullet 2", "List Number", "List Number 2")
TEMPLATE_STYLES = ("Body Text", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Quote", *LIST_STYLES)
# Settings of a template that would follow every document: a mail merge (Word asks to run its SQL
# command), write protection, and the path of the template it was made from (Word looks for it, e.g. on a
# network share).
TEMPLATE_SETTINGS_REMOVED = ("w:attachedTemplate", "w:mailMerge", "w:documentProtection", "w:writeProtection")
DOTX_MAIN = b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"


def read_word_template(path: str) -> bytes | None:
    """The Word file of the word_template valve, a .dotx turned into a .docx; None if there is none or it
    cannot be read (logged: the documents are made without it)."""
    path = path.strip()
    if not path:
        return None
    try:
        with open(path, "rb") as file:
            data = file.read(MAX_TEMPLATE_BYTES + 1)
        if len(data) > MAX_TEMPLATE_BYTES:
            raise ValueError(f"the file is larger than {MAX_TEMPLATE_BYTES // 1_000_000} MB")
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            if sum(item.file_size for item in package.infolist()) > 5 * MAX_TEMPLATE_BYTES:
                raise ValueError("the file is too large unpacked")
            types = package.read("[Content_Types].xml")
            if DOTX_MAIN not in types:
                return data
            buffer = io.BytesIO()  # python-docx opens documents only
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as document:
                for item in package.infolist():
                    content = package.read(item)
                    if item.filename == "[Content_Types].xml":
                        content = content.replace(DOTX_MAIN, DOCX_CT.WML_DOCUMENT_MAIN.encode())
                    document.writestr(item, content)
            return buffer.getvalue()
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as e:
        log.warning("The Word template %r cannot be used, documents are made without it: %s", path, e)
        return None


def _set_style(style: Any, size: float, color: str, bold: bool = False, italic: bool = False,
               before: float | None = None, after: float | None = None,
               font: str = BODY_FONT) -> None:  # fmt: skip
    """Font, colour and spacing of a built-in style; its theme font and colour are removed, they would win."""
    style.font.name = font
    for attr in THEME_FONT_ATTRS:
        style.element.rPr.rFonts.attrib.pop(qn(attr), None)
    style.font.size, style.font.bold, style.font.italic = Pt(size), bold, italic
    style.font.color.rgb = RGBColor.from_string(color)
    for attr in ("w:themeShade", "w:themeTint"):
        style.element.rPr.color.attrib.pop(qn(attr), None)
    if before is not None:
        style.paragraph_format.space_before = Pt(before)
    if after is not None:
        style.paragraph_format.space_after = Pt(after)


def _set_borders(
    parent: Any, tag: str, successors: tuple[str, ...], sides: list[tuple[str, int, str]]
) -> None:
    """Replace the w:pBdr of a paragraph (or style) or the w:tblBorders of a table. Sides in schema order
    (top, left, bottom, right, insideH, insideV), size in eighths of a point, 0 for no line."""
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    borders = OxmlElement(tag)
    for side, size, color in sides:
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:val"), "single" if size else "nil")
        if size:
            edge.set(qn("w:sz"), str(size))
            edge.set(qn("w:space"), "4" if tag == "w:pBdr" else "0")
            edge.set(qn("w:color"), color)
        borders.append(edge)
    parent.insert_element_before(borders, *successors)


def _set_cell_margins(table: Any, top: int, bottom: int) -> None:
    """Cell margins in twips; left and right stay at Word's 108."""
    margins = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("left", 108), ("bottom", bottom), ("right", 108)):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:w"), str(value))
        edge.set(qn("w:type"), "dxa")
        margins.append(edge)
    table._tbl.tblPr.insert_element_before(margins, *TBLPR_AFTER_CELL_MARGINS)


def _set_cell_shading(cell: Any, fill: str) -> None:
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shading)


class DocxWriter:
    """Render parsed blocks into a Word document: headings in the accent colour, tables with an accent
    header row, real hyperlinks, optionally on a company template (header, footer, page size, margins)."""

    def __init__(
        self,
        design: Design = DEFAULT_DESIGN,
        template: bytes | None = None,
        paper: str = "letter",
        title: str = "",
    ) -> None:
        self.design, self.paper, self.title = design, paper, title
        self.template = False
        if template:
            try:
                self.doc, self.template = Document(io.BytesIO(template)), True
                self._take_template()
                self._styles()
                return
            except Exception:
                log.warning(
                    "The Word template cannot be used, the document is made without it", exc_info=True
                )
                self.template = False
        self.doc = Document()
        self._styles()

    @property
    def body_style(self) -> str | None:
        return "Body Text" if self.template else None  # None: Normal

    def render(self, blocks: list[Block]) -> bytes:
        self._page_setup()
        for block in blocks:
            match block:
                case Heading(level, text):
                    self.doc.add_heading(to_plain(text), level=min(level, MAX_DOCX_HEADING_LEVEL))
                case Paragraph(text):
                    self._add_runs(self.doc.add_paragraph(style=self.body_style), parse_inline(text))
                case List():
                    self._list(block, level=0)
                case Quote(lines):
                    for line in filter(str.strip, lines):
                        self._add_runs(self.doc.add_paragraph(style="Quote"), parse_inline(line))
                case CodeBlock(code):
                    run = self.doc.add_paragraph(style=self.body_style).add_run(code)
                    run.font.name = CODE_FONT
                    run.font.size = CODE_FONT_SIZE
                case Table():
                    self._table(block)
                case Rule():  # a thin line: an empty paragraph with a bottom border
                    pPr = self.doc.add_paragraph(style=self.body_style)._p.get_or_add_pPr()
                    _set_borders(pPr, "w:pBdr", PPR_AFTER_BORDERS, [("bottom", 6, RULE_COLOR)])
        buffer = io.BytesIO()
        self.doc.save(buffer)
        return buffer.getvalue()

    def _take_template(self) -> None:
        """Keep the template's header, footer, page size and margins; remove its content and give the content
        styles the design (TEMPLATE_STYLES)."""
        body = self.doc.element.body
        last = body.get_or_add_sectPr()  # the last section's settings, with its header and footer
        for tag in (
            "w:headerReference",
            "w:footerReference",
        ):  # a section without its own takes the previous one's
            for earlier in reversed(body.xpath("./w:p/w:pPr/w:sectPr")):
                for reference in earlier.findall(qn(tag)):
                    if not any(
                        own.get(qn("w:type")) == reference.get(qn("w:type")) for own in last.findall(qn(tag))
                    ):
                        last.insert(0, copy.deepcopy(reference))
        for child in list(body):
            if child is not last:
                body.remove(child)
        default = Document()
        styles = self.doc.styles.element
        # The template's default character style, which Word names after its language
        default_font = next(
            (style.styleId for style in styles.findall(qn("w:style"))
             if style.get(qn("w:type")) == "character" and style.get(qn("w:default")) in ("1", "true")),
            None,
        )  # fmt: skip
        renamed = {"Normal": "BodyText", "DefaultParagraphFont": default_font}
        for name in TEMPLATE_STYLES:
            style = default.styles[name].element
            link = style.find(qn("w:link"))
            linked = default.styles.element.get_by_id(link.get(qn("w:val"))) if link is not None else None
            for new in (style, linked) if linked is not None else (style,):
                for old in styles.findall(qn("w:style")):
                    if (
                        old.name_val or ""
                    ).casefold() == new.name_val.casefold() or old.styleId == new.styleId:
                        old_link = old.find(qn("w:link"))  # e.g. "Überschrift 1 Zchn" of "Überschrift 1"
                        old_linked = (
                            styles.get_by_id(old_link.get(qn("w:val"))) if old_link is not None else None
                        )
                        if old_linked is not None:
                            styles.remove(old_linked)
                        if old.getparent() is not None:
                            styles.remove(old)
                new = copy.deepcopy(new)
                for tag in ("w:basedOn", "w:next"):
                    reference = new.find(qn(tag))
                    if reference is None or reference.get(qn("w:val")) not in renamed:
                        continue
                    if new.styleId == "BodyText" or not renamed[reference.get(qn("w:val"))]:
                        new.remove(reference)  # nothing of the template's Normal
                    else:
                        reference.set(qn("w:val"), renamed[reference.get(qn("w:val"))])
                styles.append(new)
        self._take_numbering(default)
        try:
            settings_part = self.doc.part.part_related_by(RELATIONSHIP_TYPE.SETTINGS)
        except KeyError:
            settings_part = None  # a file without settings
        for tag in TEMPLATE_SETTINGS_REMOVED if settings_part is not None else ():
            for element in settings_part.element.findall(qn(tag)):
                settings_part.element.remove(element)
                for node in element.iter():
                    if node.get(qn("r:id")):
                        settings_part.drop_rel(node.get(qn("r:id")))
        core = self.doc.core_properties  # the template's author, title ... are not the document's
        core.author = core.last_modified_by = core.subject = core.keywords = core.comments = core.category = (
            ""
        )
        core.title = self.title

    def _take_numbering(self, default: Any) -> None:
        """The list styles now point to the default template's numbering: copy it in, with new numbers."""
        source = default.part.numbering_part.element
        try:
            numbering = self.doc.part.part_related_by(RELATIONSHIP_TYPE.NUMBERING).element
        except KeyError:
            package, numbering = self.doc.part.package, parse_xml(f"<w:numbering {nsdecls('w')}/>")
            partname = package.next_partname("/word/numbering%d.xml")
            self.doc.part.relate_to(
                NumberingPart(partname, DOCX_CT.WML_NUMBERING, numbering, package),
                RELATIONSHIP_TYPE.NUMBERING,
            )
        abstract_ids = [int(a.get(qn("w:abstractNumId"))) for a in numbering.findall(qn("w:abstractNum"))]
        for abstract_id, name in enumerate(LIST_STYLES, start=max(abstract_ids, default=-1) + 1):
            num_id = self.doc.styles[name].element.pPr.numPr.numId
            used = source.num_having_numId(num_id.val).abstractNumId.val
            abstract = copy.deepcopy(source.xpath(f"w:abstractNum[@w:abstractNumId='{used}']")[0])
            abstract.set(qn("w:abstractNumId"), str(abstract_id))
            for nsid in abstract.findall(qn("w:nsid")):  # Word's list identity; the template may use the same
                abstract.remove(nsid)
            numbering.insert_element_before(abstract, "w:num", "w:numIdMacAtCleanup")
            num_id.val = numbering.add_num(abstract_id).numId

    def _styles(self) -> None:
        """A calm, modern look: dark grey body text with some air, headings in the accent colour."""
        styles, accent, font = self.doc.styles, self.design.accent, self.design.font
        body = styles[self.body_style or "Normal"]
        _set_style(body, 10.5, TEXT_COLOR, after=6, font=font)
        body.paragraph_format.line_spacing = 1.15
        _set_style(styles["Heading 1"], 16, accent, bold=True, before=20, after=6, font=font)
        _set_style(styles["Heading 2"], 13, TEXT_COLOR, bold=True, before=14, after=4, font=font)
        _set_style(styles["Heading 3"], 11, accent, bold=True, before=10, after=2, font=font)
        _set_style(
            styles["Heading 4"], 10.5, TEXT_COLOR, bold=True, italic=True, before=8, after=2, font=font
        )
        quote = styles["Quote"]
        _set_style(quote, 10.5, MUTED_COLOR, italic=True, before=6, after=10, font=font)
        quote.paragraph_format.left_indent = Cm(0.4)
        _set_borders(quote.element.get_or_add_pPr(), "w:pBdr", PPR_AFTER_BORDERS, [("left", 18, accent)])
        self._for_previews()

    def _for_previews(self) -> None:
        """Word takes fonts and colours from the paragraph styles. Other viewers also read the document
        defaults, the theme fonts and the character styles linked to the paragraph styles, and they may have
        no Symbol font for Word's bullets (boxes instead of dots). All get the design."""
        font, styles = self.design.font, self.doc.styles
        for fonts in (
            styles.element.xpath("w:docDefaults/w:rPrDefault/w:rPr/w:rFonts") if not self.template else ()
        ):
            for attr in THEME_FONT_ATTRS:
                fonts.attrib.pop(qn(attr), None)
            for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                fonts.set(qn(attr), font)
        for part in (
            self.doc.part.package.iter_parts() if not self.template else ()
        ):  # the template's header and
            if str(part.partname).startswith("/word/theme/"):  # footer may use its theme fonts
                part._blob = re.sub(THEME_FONTS_RE.encode(), lambda m: m[1] + font.encode() + m[2], part.blob)
        for style in styles:
            link = style.element.find(qn("w:link"))
            if style.type != WD_STYLE_TYPE.PARAGRAPH or link is None or style.element.rPr is None:
                continue
            linked = styles.element.get_by_id(link.get(qn("w:val")))
            if linked is not None:  # e.g. "Heading 1 Char": same look as its paragraph style
                for old in linked.findall(qn("w:rPr")):
                    linked.remove(old)
                linked.append(copy.deepcopy(style.element.rPr))
        for level in self.doc.part.numbering_part.element.iter(qn("w:lvl")):
            kind = level.find(qn("w:numFmt"))
            if kind is None or kind.get(qn("w:val")) != "bullet":
                continue
            indent = level.find(f"{qn('w:pPr')}/{qn('w:ind')}")
            sub_item = indent is not None and int(indent.get(qn("w:left"), "0")) > 720
            level.find(qn("w:lvlText")).set(qn("w:val"), "–" if sub_item else "•")
            for old in level.findall(qn("w:rPr")):
                level.remove(old)
            props = OxmlElement("w:rPr")
            fonts = OxmlElement("w:rFonts")
            fonts.set(qn("w:ascii"), font)
            fonts.set(qn("w:hAnsi"), font)
            color = OxmlElement("w:color")
            color.set(qn("w:val"), self.design.accent)
            props.append(fonts)
            props.append(color)
            level.append(props)

    def _page_setup(self) -> None:
        """Paper and margins; a template keeps its own."""
        section = self.doc.sections[0]
        width, height = PAPER_SIZES[self.paper]
        page = {"page_width": width, "page_height": height, "left_margin": Cm(2.5), "right_margin": Cm(2.5),
                "top_margin": Cm(2.5), "bottom_margin": Cm(2.2), "header_distance": Cm(1.25),
                "footer_distance": Cm(1.25)}  # fmt: skip
        for name, value in page.items():
            if not self.template or getattr(section, name) is None:  # what a template lacks
                setattr(section, name, value)

    def _list(self, block: List, level: int) -> None:
        style = (ORDERED_LIST_STYLES if block.ordered else BULLET_LIST_STYLES)[level]
        num_id: int | None = None
        for item in block.items:
            paragraph = self.doc.add_paragraph(style=style)
            paragraph.paragraph_format.space_after = Pt(3)
            if block.ordered:
                if num_id is None:  # each list gets its own numbering, so the next one restarts
                    num_id = self._new_numbering(paragraph.style, block.start)
                paragraph._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num_id
            self._add_runs(paragraph, parse_inline(item.text))
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
        """Header row in the accent colour, every second row lightly tinted, thin lines between rows only."""
        doc_table = self.doc.add_table(rows=len(table.rows), cols=len(table.aligns))
        lines = [("top", 0, ""), ("left", 0, ""), ("bottom", 8, self.design.accent), ("right", 0, ""),
                 ("insideH", 4, RULE_COLOR), ("insideV", 0, "")]  # fmt: skip
        _set_borders(doc_table._tbl.tblPr, "w:tblBorders", TBLPR_AFTER_BORDERS, lines)
        _set_cell_margins(doc_table, top=50, bottom=50)
        for row_index, (doc_row, row) in enumerate(zip(doc_table.rows, table.rows, strict=True)):
            if row_index == 0:  # repeat the header row on every page
                header_flag = OxmlElement("w:tblHeader")
                header_flag.set(qn("w:val"), "true")
                doc_row._tr.get_or_add_trPr().append(header_flag)
            fill = self.design.accent if row_index == 0 else self.design.band if row_index % 2 == 0 else None
            for cell, text, align in zip(doc_row.cells, row, table.aligns, strict=True):
                paragraph = cell.paragraphs[0]
                if self.template:
                    paragraph.style = self.body_style
                paragraph.alignment = DOCX_ALIGNMENT.get(align or "")
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                runs = parse_inline(text, Run("", bold=row_index == 0))
                self._add_runs(paragraph, runs, WHITE if row_index == 0 else None)
                if fill:
                    _set_cell_shading(cell, fill)
        self.doc.add_paragraph(style=self.body_style).paragraph_format.space_after = Pt(4)

    def _add_runs(self, paragraph: DocxParagraph, runs: list[Run], color: str | None = None) -> None:
        for run in runs:
            link = run.url if run.url and is_word_link(run.url) else None
            docx_run = paragraph.add_run(run.text if link else _run_text(run, with_url=True))
            docx_run.bold = run.bold or None
            docx_run.italic = run.italic or None
            if color:
                docx_run.font.color.rgb = RGBColor.from_string(color)
            if run.strike:
                docx_run.font.strike = True
            if run.code:
                docx_run.font.name = CODE_FONT
            if link:
                self._make_link(paragraph, docx_run, link.strip(), color or self.design.accent)

    @staticmethod
    def _make_link(paragraph: DocxParagraph, docx_run: Any, url: str, color: str) -> None:
        docx_run.font.underline = True
        docx_run.font.color.rgb = RGBColor.from_string(color)
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True))
        docx_run._r.addprevious(link)
        link.append(docx_run._r)


def markdown_to_docx(
    blocks: list[Block], design: Design = DEFAULT_DESIGN, template: bytes | None = None, paper: str = "letter"
) -> bytes:
    return DocxWriter(design, template, paper, document_title(blocks)).render(blocks)


# ----------------------------------------------------------------------------------------------- Excel --

CURRENCY_RE = "|".join(map(re.escape, CURRENCY_FORMATS))
NUMBER_RE = re.compile(  # thousands separators: . , ' ’ space, no-break space, narrow no-break space
    rf"(?P<prefix>{CURRENCY_RE})?\s*"
    r"(?P<sign>[-+]?)"
    r"(?P<int>\d{1,3}(?:(?P<group>[.,'\u2019 \u00a0\u202f])\d{3}(?:(?P=group)\d{3})*)?|\d+)"
    r"(?:(?!(?P=group))(?P<decimal>[.,])(?P<fraction>\d+))?"  # decimal mark differs from the separator
    rf"\s*(?P<suffix>%|{CURRENCY_RE})?"
)


def parse_number(text: str) -> tuple[str | int | float, str | None]:
    """Convert '1.234,56 €', '1.234,56 EUR', '-3,2 %', '1,234.5' etc. to (number, Excel number format).

    A dot or comma followed by exactly three digits is read as a thousands separator, so
    "2.000" is 2000 and "1,234" is 1234. Values that do not look like a plain number, integers
    with leading zeros (postcodes, IDs) and phone numbers ("+49 711 123", "+1 234") are returned
    unchanged as (text, None).
    """
    match = NUMBER_RE.fullmatch(text.strip().replace("\u2212", "-"))  # Unicode minus sign
    if not match or (match["prefix"] and match["suffix"]):
        return text, None
    if match["sign"] == "+" and (match["group"] or "").isspace() and not (match["prefix"] or match["suffix"]):
        return text, None  # international phone number
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


# The interface language (Open WebUI keeps it in localStorage.locale) and the time zone of the browser.
BROWSER_SETTINGS_JS = """
let locale = '';
let zone = '';
try { locale = localStorage.getItem('locale') || ''; } catch {}
try { zone = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; } catch {}
return [String(locale || navigator.language || ''), String(zone)];
"""


# ---------------------------------------------------------------------------------------------- Action --


class BrowserError(Exception):
    """The browser did not run a snippet: blocked by the page's CSP, tab closed or disconnected, no answer."""


def zone(name: str) -> tzinfo | None:
    """The IANA time zone, None if the name is empty or unknown."""
    try:
        return ZoneInfo(name.strip()) if name.strip() else None
    except (ZoneInfoNotFoundError, ValueError, OSError):  # OSError: on some systems for a name too long for a file
        return None


@dataclass
class Browser:
    """Thin wrapper around Open WebUI's event callbacks."""

    emitter: EventCallback | None
    caller: EventCallback | None

    async def notify(self, level: Literal["info", "success", "warning", "error"], text: str) -> None:
        if self.emitter is not None:
            await self.emitter({"type": "notification", "data": {"type": level, "content": text}})

    async def call_js(self, code: str) -> Any:
        """Run a snippet in the browser and return its result. A snippet that throws never calls back, a
        disconnected session answers {'error': ...}: both raise BrowserError."""
        if self.caller is None:
            raise BrowserError("__event_call__ is not available, cannot run code in the browser")
        try:
            result = await asyncio.wait_for(
                self.caller({"type": "execute", "data": {"code": code}}), BROWSER_TIMEOUT_S
            )
        except TimeoutError as e:
            raise BrowserError(f"the browser did not answer within {BROWSER_TIMEOUT_S} s") from e
        if isinstance(result, dict) and "error" in result:
            raise BrowserError(str(result["error"])[:200])
        return result

    async def run_js(self, code: str) -> None:
        """Run one of the export snippets, which all return true."""
        result = await self.call_js(code)
        if result is not True:
            raise BrowserError(f"unexpected result from the browser: {str(result)[:200]}")

    async def download(self, data: bytes, filename: str, mime: str) -> None:
        await self.run_js(download_js(data, filename, mime))

    async def settings(self) -> tuple[str, str]:
        """The interface language (e.g. 'de-DE') and time zone of the user's browser."""
        result = await self.call_js(BROWSER_SETTINGS_JS)
        if not (isinstance(result, list) and len(result) == 2):
            raise BrowserError(f"unexpected result from the browser: {str(result)[:200]}")
        return str(result[0] or ""), str(result[1] or "")


@dataclass
class Export:
    """What an export needs besides the content: texts in the user's language, time zone and browser."""

    texts: dict[str, str]
    language: str
    timezone: tzinfo | None
    browser: Browser

    def now(self) -> datetime:
        return datetime.now(self.timezone)


class Action:
    class Valves(BaseModel):
        language: Literal["auto", "en", "de"] = Field(
            default="auto",
            description="Language of the messages, default file names, the print footer and the Word paper "
            "size (en: Letter, de: A4). auto = the user's interface language (English if there is no "
            "translation).",
        )
        timezone: str = Field(
            default="",
            description="IANA time zone for file name timestamps and the print footer date, "
            "e.g. 'Europe/Berlin'. Empty or unknown = the time zone of the user's browser.",
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
        accent_color: str = Field(
            default=DEFAULT_ACCENT,
            description="Accent colour of Word documents (headings, table header rows, links) as a hex value "
            "without '#', e.g. 1F4E79. Same as the valve of Office Documents.",
        )
        font: str = Field(
            default=BODY_FONT,
            description="Font of Word documents, e.g. Calibri, Arial or Aptos. It has to be installed where "
            "the files are opened; otherwise Office shows a similar one.",
        )
        word_template: str = Field(
            default="",
            description="Optional Word file (.docx or .dotx) on the Open WebUI server, e.g. a letterhead: "
            "Word documents get its header, footer, page size and margins, e.g. "
            "/app/backend/data/templates/letterhead.dotx.",
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
        browser = Browser(__event_emitter__, __event_call__)
        # Without a sub-action (e.g. an API call) Open WebUI passes the function's own ID: export to Word.
        action_id = __id__ if __id__ in ACTIONS else "docx"
        try:
            export = await self._export_context(browser)
        except BrowserError as e:
            log.warning("Response export: the browser cannot run the export: %s", e)
            language = self.valves.language if self.valves.language != "auto" else "en"
            await browser.notify("error", TRANSLATIONS[language]["browser_failed"])
            return
        content = clean_markdown(message_text(find_message(body)))
        if not content:
            await browser.notify("warning", export.texts["empty_message"])
            return
        try:
            await getattr(self, f"_export_{action_id}")(content, export)
        except BrowserError as e:
            log.warning("Response export %r: the browser did not run it: %s", action_id, e)
            await browser.notify("error", export.texts["browser_failed"])
        except Exception:
            log.exception("Response export %r failed", action_id)
            await browser.notify("error", export.texts["export_failed"])

    async def _export_context(self, browser: Browser) -> Export:
        """Language and time zone from the valves; 'auto' and an empty time zone come from the browser."""
        language, timezone = self.valves.language, zone(self.valves.timezone)
        if language == "auto" or timezone is None:
            locale, browser_zone = await browser.settings()
            if language == "auto":
                code = locale.replace("_", "-").split("-")[0].strip().lower()
                language = code if code in TRANSLATIONS else "en"
            timezone = timezone or zone(browser_zone)
        return Export(TRANSLATIONS[language], language, timezone, browser)

    def _design(self) -> Design:
        accent = self.valves.accent_color.strip().lstrip("#")
        font = " ".join(self.valves.font.split())
        return Design(
            accent.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", accent) else DEFAULT_ACCENT,
            font if FONT_NAME_RE.fullmatch(font) else BODY_FONT,
        )

    async def _export_docx(self, content: str, export: Export) -> None:
        blocks = parse_blocks(content)

        def build() -> bytes:
            template = read_word_template(self.valves.word_template)
            return markdown_to_docx(blocks, self._design(), template, export.texts["paper"])

        data = await asyncio.to_thread(build)
        await export.browser.download(data, f"{self._file_stem(blocks, export)}.docx", DOCX_MIME)

    async def _export_xlsx(self, content: str, export: Export) -> None:
        blocks = parse_blocks(content)
        tables = [block for block in blocks if isinstance(block, Table)]
        if not tables:
            await export.browser.notify("warning", export.texts["no_tables"])
            return
        data = await asyncio.to_thread(tables_to_xlsx, tables, export.texts["sheet_name"])
        await export.browser.download(data, f"{self._file_stem(blocks, export)}.xlsx", XLSX_MIME)

    async def _export_pdf(self, content: str, export: Export) -> None:
        blocks = parse_blocks(content)
        footer = ""
        if self.valves.print_footer:
            footer = export.texts["print_footer"].format(
                date=export.now().strftime(export.texts["date_format"])
            )
        page = build_print_html(blocks, self._file_stem(blocks, export), footer, export.language)
        await export.browser.run_js(print_js(page))

    async def _export_email(self, content: str, export: Export) -> None:
        texts, browser = export.texts, export.browser
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

    @staticmethod
    def _file_stem(blocks: list[Block], export: Export) -> str:
        """File name without extension: first heading (or a default) plus a timestamp."""
        title = safe_filename(document_title(blocks), export.texts["default_title"])
        return f"{title} {export.now():%Y-%m-%d %H%M}"
