"""Tests for the Open WebUI plugins in the parent folder.

Run from anywhere (Python 3.11+):
    pip install python-docx openpyxl pydantic pytest
    python -I -m pytest -q path/to/tests

Optional checks run automatically when the tools are installed: Node.js (executes the generated
browser snippets), a Chromium-based browser (renders the print page to PDF) and LibreOffice
(converts the .docx to PDF); pdftotext/pdfinfo (poppler) are needed for the last two.
"""

import ast
import asyncio
import base64
import email
import email.policy
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from types import ModuleType, SimpleNamespace
from typing import Any
from urllib.parse import parse_qs, urlsplit
from zoneinfo import ZoneInfo

import pytest
from docx import Document
from openpyxl import load_workbook

PLUGIN_DIR = Path(__file__).resolve().parent.parent
STUB_JS = Path(__file__).resolve().parent / "browser_stub.js"
PLUGIN_FILES = ("response_export.py", "context_guard.py", "company_context.py", "reasoning_toggle.py")
ALLOWED_THIRD_PARTY = {"docx", "openpyxl", "pydantic"}
LAZY_IMPORTS = {"open_webui"}  # only imported inside functions, when running in Open WebUI
FRONTMATTER_KEYS = {
    "title",
    "author",
    "author_url",
    "version",
    "license",
    "required_open_webui_version",
    "description",
}


def load_plugin(file_name: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(f"plugin_{Path(file_name).stem}", PLUGIN_DIR / file_name)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


rx = load_plugin("response_export.py")
cg = load_plugin("context_guard.py")
cc = load_plugin("company_context.py")
rt = load_plugin("reasoning_toggle.py")

SAMPLE = """<details type="reasoning" done="true"><summary>Thought for 3 seconds</summary>
SECRET REASONING
</details>
# Quarterly report C#

Intro line one
continues **bold *nested*** and `code` and [docs](https://example.org/a_(b)).
Steps:
1. First
   continued
2. Second
   - sub a
   - sub b
3. Third

Between lists.

1. Restart
2. Again

| Item | Amount | Share |
|:-----|-------:|:-----:|
| A \\| B | 1.234,56 € | -3,2 % |
| ragged |
| C | 2.000 | 1,5 | surplus |

> A quote
> second line

```bash
# not a heading
echo "<b>hi</b>"
```

<script>alert(1)</script> [bad](javascript:alert(1))
"""


class Recorder:
    """Collects the events an action/filter sends to Open WebUI."""

    def __init__(self) -> None:
        self.events: list[dict[str, Any]] = []
        self.calls: list[dict[str, Any]] = []

    async def emit(self, event: dict[str, Any]) -> None:
        self.events.append(event)

    async def call(self, event: dict[str, Any]) -> bool:
        self.calls.append(event)
        return True

    def notifications(self) -> list[tuple[str, str]]:
        return [(e["data"]["type"], e["data"]["content"]) for e in self.events if e["type"] == "notification"]

    def code(self) -> str:
        assert len(self.calls) == 1, self.calls
        assert self.calls[0]["type"] == "execute"
        return self.calls[0]["data"]["code"]


def js_value(code: str, prefix: str) -> Any:
    """Decode the JSON literal that follows `prefix` in a generated JS snippet."""
    start = code.index(prefix) + len(prefix)
    return json.JSONDecoder().raw_decode(code, start)[0]


def make_filter(module: ModuleType, **valves: Any) -> Any:
    flt = module.Filter()
    flt.valves = flt.Valves(**valves)
    return flt


def run_action(
    action_id: str | None,
    content: Any = SAMPLE,
    with_call: bool = True,
    body: dict[str, Any] | None = None,
    **valves: Any,
) -> Recorder:
    action = rx.Action()
    action.valves = action.Valves(**valves)
    recorder = Recorder()
    body = body or {
        "id": "m2",
        "messages": [
            {"id": "m1", "role": "user", "content": "Please write a report"},
            {"id": "m2", "role": "assistant", "content": content},
        ],
    }
    asyncio.run(
        action.action(
            body,
            __id__=action_id,
            __event_emitter__=recorder.emit,
            __event_call__=recorder.call if with_call else None,
        )
    )
    return recorder


def downloaded(recorder: Recorder) -> tuple[bytes, str, str]:
    code = recorder.code()
    data = base64.b64decode(js_value(code, "atob("))
    return data, js_value(code, "link.download = "), js_value(code, "{ type: ")


# ------------------------------------------------------------------------------------- all plugins --


@pytest.mark.parametrize("file_name", PLUGIN_FILES)
def test_frontmatter(file_name: str) -> None:
    """Replicates Open WebUI's frontmatter parser: no stray 'key:' lines in the docstring."""
    lines = (PLUGIN_DIR / file_name).read_text(encoding="utf-8").splitlines()
    assert lines[0] == '"""'
    pattern = re.compile(r"^\s*([a-z_]+):\s*(.*)\s*$", re.IGNORECASE)
    meta = {}
    for line in lines[1:]:
        if '"""' in line:
            break
        if match := pattern.match(line):
            meta[match[1]] = match[2]
    assert set(meta) == FRONTMATTER_KEYS
    assert re.fullmatch(r"\d+\.\d+\.\d+", meta["version"])
    assert meta["license"] == "MIT"
    assert meta["required_open_webui_version"] == "0.11.4"
    assert "Security & privacy" in "\n".join(lines[: lines.index('"""', 1)])


@pytest.mark.parametrize("file_name", PLUGIN_FILES)
def test_only_bundled_libraries(file_name: str) -> None:
    tree = ast.parse((PLUGIN_DIR / file_name).read_text(encoding="utf-8"))
    modules = {alias.name for node in ast.walk(tree) if isinstance(node, ast.Import) for alias in node.names}
    modules |= {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom) and node.module}
    top_level = {name.split(".")[0] for name in modules}
    assert top_level <= set(sys.stdlib_module_names) | ALLOWED_THIRD_PARTY | LAZY_IMPORTS


# -------------------------------------------------------------------------------- response export --


def test_action_buttons_are_language_neutral() -> None:
    action = rx.Action()
    expected = ["Word (.docx)", "Excel (.xlsx)", "PDF", "E-Mail"]
    assert [a["id"] for a in action.actions] == ["docx", "xlsx", "pdf", "email"]
    assert [a["name"] for a in action.actions] == expected
    action.valves = action.Valves(language="de")
    assert [a["name"] for a in action.actions] == expected
    for item in action.actions:
        svg = base64.b64decode(item["icon_url"].split(",", 1)[1]).decode()
        assert item["icon_url"].startswith("data:image/svg+xml;base64,")
        assert 'fill="none"' in svg and 'stroke="currentColor"' in svg and 'stroke-width="1.5"' in svg


def test_docx_structure() -> None:
    data, filename, mime = downloaded(run_action("docx"))
    assert mime == rx.DOCX_MIME
    assert re.fullmatch(r"Quarterly report C# \d{4}-\d{2}-\d{2} \d{4}\.docx", filename)
    doc = Document(io.BytesIO(data))
    paragraphs = [(p.style.name, p.text) for p in doc.paragraphs if p.text]
    all_text = "\n".join(text for _, text in paragraphs)
    assert ("Heading 1", "Quarterly report C#") in paragraphs
    assert "SECRET REASONING" not in all_text and "Thought for" not in all_text
    assert ("Normal", '# not a heading\necho "<b>hi</b>"') in paragraphs
    assert ("List Number", "First continued") in paragraphs
    assert ("List Bullet 2", "sub a") in paragraphs
    assert ("Quote", "A quote") in paragraphs and ("Quote", "second line") in paragraphs
    intro = next(p for p in doc.paragraphs if p.text.startswith("Intro"))
    assert intro.text == (
        "Intro line one\ncontinues bold nested and code and docs (https://example.org/a_(b)).\nSteps:"
    )
    assert [r.text for r in intro.runs if r.bold] == ["bold ", "nested"]
    assert [r.text for r in intro.runs if r.italic] == ["nested"]
    assert [r.text for r in intro.runs if r.font.name == "Consolas"] == ["code"]
    code_run = next(p for p in doc.paragraphs if p.text.startswith("# not")).runs[0]
    assert code_run.font.name == "Consolas"
    assert "<script>alert(1)</script> bad (javascript:alert(1))" in all_text  # plain text, harmless

    table = doc.tables[0]
    assert table.style.name == "Table Grid"
    assert [c.text for c in table.rows[0].cells] == ["Item", "Amount", "Share"]
    assert all(run.bold for cell in table.rows[0].cells for run in cell.paragraphs[0].runs)
    assert [c.text for c in table.rows[1].cells] == ["A | B", "1.234,56 €", "-3,2 %"]
    assert [c.text for c in table.rows[2].cells] == ["ragged", "", ""]
    assert [c.text for c in table.rows[3].cells] == ["C", "2.000", "1,5"]
    assert not any(run.bold for run in table.rows[1].cells[0].paragraphs[0].runs)


def test_docx_numbered_lists_restart() -> None:
    data, _, _ = downloaded(run_action("docx"))
    doc = Document(io.BytesIO(data))
    numbered = [p for p in doc.paragraphs if p.style.name == "List Number"]
    num_ids = [p._p.pPr.numPr.numId.val for p in numbered]
    assert [p.text for p in numbered] == ["First continued", "Second", "Third", "Restart", "Again"]
    assert num_ids[0] == num_ids[1] == num_ids[2] != num_ids[3] == num_ids[4]
    numbering = doc.part.numbering_part.element
    for num_id in (num_ids[0], num_ids[3]):
        num = numbering.num_having_numId(num_id)
        assert num.xpath("./w:lvlOverride/w:startOverride/@w:val") == ["1"]


def test_docx_list_starting_number_is_kept() -> None:
    data, _, _ = downloaded(run_action("docx", "Intro\n\n3. three\n4. four"))
    doc = Document(io.BytesIO(data))
    item = next(p for p in doc.paragraphs if p.text == "three")
    num = doc.part.numbering_part.element.num_having_numId(item._p.pPr.numPr.numId.val)
    assert num.xpath("./w:lvlOverride/w:startOverride/@w:val") == ["3"]


@pytest.mark.parametrize(
    ("text", "value", "number_format"),
    [
        ("42", 42, None),
        ("-17", -17, None),
        ("1,5", 1.5, None),
        ("0,5", 0.5, None),
        ("3.14159", 3.14159, None),
        ("2.000", 2000, None),
        ("1.234.567", 1234567, None),
        ("1.234,56", 1234.56, None),
        ("1,234.56", 1234.56, None),
        ("1 234,5", 1234.5, None),
        ("1 000", 1000, None),
        ("1'000'000", 1000000, None),
        ("−5", -5, None),
        ("12%", 0.12, "0%"),
        ("-3,2 %", -0.032, "0.0#%"),
        ("1.234,56 €", 1234.56, '#,##0.00 "€"'),
        ("€ -12,50", -12.5, '#,##0.00 "€"'),
        ("$1,234.50", 1234.5, '"$"#,##0.00'),
        ("1.234,56 EUR", 1234.56, '#,##0.00 "€"'),
        ("USD 12", 12.0, '"$"#,##0.00'),
        ("£5 %", "£5 %", None),
        ("01067", "01067", None),
        ("1 2", "1 2", None),
        ("1.2.3", "1.2.3", None),
        ("12.10.2026", "12.10.2026", None),
        ("1.234,567.8", "1.234,567.8", None),
        ("+49 30 123456", "+49 30 123456", None),
        ("1234567890123456", "1234567890123456", None),
        ("=SUM(A1:A2)", "=SUM(A1:A2)", None),
        ("%", "%", None),
        ("abc", "abc", None),
    ],
)
def test_parse_number(text: str, value: Any, number_format: str | None) -> None:
    result, fmt = rx.parse_number(text)
    assert fmt == number_format
    assert result == pytest.approx(value) if isinstance(value, float) else result == value
    assert type(result) is type(value)


def test_xlsx_values_and_formats() -> None:
    content = SAMPLE + (
        "\n\n| Code | Formula | Total |\n|---|---|---|\n| 01067 | =1+1 | 12% |\n"
        "\n```\n| in | code |\n|---|---|\n| x | y |\n```\n"
    )
    data, filename, mime = downloaded(run_action("xlsx", content))
    assert mime == rx.XLSX_MIME and filename.endswith(".xlsx")
    workbook = load_workbook(io.BytesIO(data))
    assert workbook.sheetnames == ["Table 1", "Table 2"]
    sheet = workbook["Table 1"]
    assert [c.value for c in sheet[1]] == ["Item", "Amount", "Share"]
    assert all(c.font.bold and c.fill.fgColor.rgb.endswith("DDE4EE") for c in sheet[1])
    assert sheet.freeze_panes == "A2"
    assert sheet.auto_filter.ref == "A1:C4"
    assert sheet["A2"].value == "A | B"
    assert sheet["B2"].value == pytest.approx(1234.56) and sheet["B2"].number_format == '#,##0.00 "€"'
    assert sheet["C2"].value == pytest.approx(-0.032) and sheet["C2"].number_format == "0.0#%"
    assert sheet["B3"].value is None and sheet["B4"].value == 2000 and sheet["C4"].value == 1.5
    assert sheet.max_column == 3  # surplus cell of a ragged row is dropped like in GFM
    assert 8 <= sheet.column_dimensions["B"].width <= 60
    second = workbook["Table 2"]
    assert second["A2"].value == "01067" and second["A2"].data_type == "s"
    assert second["B2"].value == "=1+1" and second["B2"].data_type == "s"  # text, not a formula
    assert second["C2"].value == pytest.approx(0.12) and second["C2"].number_format == "0%"


def test_xlsx_german_sheet_names() -> None:
    data, _, _ = downloaded(run_action("xlsx", language="de"))
    assert load_workbook(io.BytesIO(data)).sheetnames == ["Tabelle 1"]


@pytest.mark.parametrize(("language", "text"), [("en", "does not contain a table"), ("de", "keine Tabelle")])
def test_xlsx_without_table_warns(language: str, text: str) -> None:
    recorder = run_action("xlsx", "# Just text", language=language)
    assert not recorder.calls
    assert recorder.notifications()[0][0] == "warning" and text in recorder.notifications()[0][1]


EMAIL = """Sure, here is a draft for you:

**Subject:** Meeting on Friday

Dear Ms. Smith,

thank you for your message. Agenda:
- Budget
- Timeline

Best regards"""


def test_email_mailto() -> None:
    recorder = run_action("email", EMAIL)
    url = js_value(recorder.code(), "link.href = ")
    assert url.startswith("mailto:?")
    query = parse_qs(urlsplit(url).query)
    assert query["subject"] == ["Meeting on Friday"]
    body = query["body"][0]
    assert body.startswith("Dear Ms. Smith,\r\n\r\nthank you")
    assert "\r\n• Budget\r\n• Timeline\r\n" in body
    assert "here is a draft" not in body
    assert not recorder.notifications()


@pytest.mark.parametrize(("language", "text"), [("en", "too long"), ("de", "zu lang")])
def test_email_eml_fallback(language: str, text: str) -> None:
    content = EMAIL + "\n\n| Item | Cost |\n|---|---|\n| <b>A</b> | 5 € |\n" + "More text. " * 300
    recorder = run_action("email", content, language=language)
    data, filename, mime = downloaded(recorder)
    assert filename == "Meeting on Friday.eml" and mime == "message/rfc822"
    message = email.message_from_bytes(data, policy=email.policy.default)
    assert message["Subject"] == "Meeting on Friday"
    assert message["X-Unsent"] == "1"
    assert message.get_content_type() == "multipart/alternative"
    parts = {part.get_content_type(): part.get_content() for part in message.iter_parts()}
    assert set(parts) == {"text/plain", "text/html"}
    assert "here is a draft" not in parts["text/plain"] and "• Budget" in parts["text/plain"]
    assert '<table border="1"' in parts["text/html"] and "<li>Budget</li>" in parts["text/html"]
    assert "&lt;b&gt;A&lt;/b&gt;" in parts["text/html"] and "<b>A</b>" not in parts["text/html"]
    assert [kind for kind, _ in recorder.notifications()] == ["info"]
    assert text in recorder.notifications()[0][1]


def test_email_always_eml_without_subject_uses_defaults() -> None:
    recorder = run_action("email", "Just a body", always_eml=True, language="de")
    data, filename, _ = downloaded(recorder)
    assert filename == "E-Mail-Entwurf.eml"
    assert email.message_from_bytes(data, policy=email.policy.default)["Subject"] == "Entwurf"
    assert not recorder.notifications()


def test_email_subject_falls_back_to_heading() -> None:
    content = "```\nSubject: in code\n```\n# Offer *2026*\nText"
    blocks = rx.parse_blocks(content)
    assert rx.split_email(blocks) == ("", blocks)
    url = js_value(run_action("email", content).code(), "link.href = ")
    assert parse_qs(urlsplit(url).query)["subject"] == ["Offer 2026"]


def test_email_subject_line() -> None:
    blocks = rx.parse_blocks("Here you go:\n## Subject: Hi\nDear Bob,\nthanks")
    assert rx.split_email(blocks) == ("Hi", [rx.Paragraph("Dear Bob,\nthanks")])
    blocks = rx.parse_blocks("Draft:\nBetreff: Angebot\nHallo Bob")
    assert rx.split_email(blocks) == ("Angebot", [rx.Paragraph("Hallo Bob")])
    blocks = rx.parse_blocks("#Subject: not a heading")
    assert rx.split_email(blocks)[0] == ""


def test_emphasis_across_lines() -> None:
    blocks = rx.parse_blocks("**bold\ncontinued** x")
    assert rx.blocks_to_html(blocks) == "<p><strong>bold<br>continued</strong> x</p>"
    assert rx.blocks_to_text(blocks) == "bold\ncontinued x"


def print_page(recorder: Recorder) -> str:
    return js_value(recorder.code(), "frame.srcdoc = ")


@pytest.mark.parametrize(
    ("language", "footer"),
    [
        ("en", r"Created with the AI assistant on \d{4}-\d{2}-\d{2}\. Please check"),
        ("de", r"Erstellt mit dem KI-Assistenten am \d{2}\.\d{2}\.\d{4} – Inhalte bitte prüfen\."),
    ],
)
def test_print_html(language: str, footer: str) -> None:
    content = '# Q&A <b>"Report"</b>\n' + SAMPLE
    page = print_page(run_action("pdf", content, language=language))
    title = re.search(r"<title>(.*?)</title>", page)[1]
    assert re.fullmatch(r"Q&amp;A b Report b \d{4}-\d{2}-\d{2} \d{4}", title)
    assert f'<html lang="{language}">' in page
    assert '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'' in page
    assert re.search(footer, page)
    assert "SECRET REASONING" not in page
    assert "<script>" not in page and "&lt;script&gt;alert(1)&lt;/script&gt;" in page
    assert "javascript:" not in page.split("</title>")[1].replace("(javascript:alert(1))", "")
    assert '<a href="https://example.org/a_(b)">docs</a>' in page
    assert "<ol><li>First continued</li><li>Second<ul><li>sub a</li><li>sub b</li></ul></li>" in page
    assert '<th style="text-align:right">Amount</th>' in page
    assert "@page { size: A4;" in page


def test_print_html_escapes_title_and_footer() -> None:
    page = rx.build_print_html([], '<x> & "y"', "<b>footer</b>", 'en"><script>')
    assert "<title>&lt;x&gt; &amp; &quot;y&quot;</title>" in page
    assert '<div class="footer">&lt;b&gt;footer&lt;/b&gt;</div>' in page
    assert '<html lang="en&quot;&gt;&lt;script&gt;">' in page


def test_print_without_footer() -> None:
    assert 'class="footer"' not in print_page(run_action("pdf", print_footer=False))


def test_list_of_parts_content() -> None:
    content = [
        {"type": "text", "text": "# Parts title\n\n| a | b |\n|---|---|\n| 1 | 2 |"},
        {"type": "image_url", "image_url": {"url": "data:image/png;base64,AAAA"}},
        {"type": "text", "text": "Closing line"},
    ]
    data, filename, _ = downloaded(run_action("docx", content))
    assert filename.startswith("Parts title ")
    doc = Document(io.BytesIO(data))
    assert "Closing line" in [p.text for p in doc.paragraphs]
    assert doc.tables[0].rows[1].cells[1].text == "2"


def test_clicked_message_is_exported() -> None:
    body = {
        "id": "a1",
        "messages": [
            {"id": "a1", "role": "assistant", "content": "# First"},
            {"id": "a2", "role": "assistant", "content": "# Second"},
        ],
    }
    assert downloaded(run_action("docx", body=body))[1].startswith("First ")


@pytest.mark.parametrize("content", ["", "   ", None, [], '<details type="reasoning">x</details>'])
def test_empty_message_warns(content: Any) -> None:
    recorder = run_action("docx", content)
    assert not recorder.calls
    assert recorder.notifications() == [("warning", rx.TRANSLATIONS["en"]["empty_message"])]


def test_export_error_shows_toast(monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture) -> None:
    def broken(_blocks: Any) -> bytes:
        raise ValueError("boom")

    monkeypatch.setattr(rx, "markdown_to_docx", broken)
    recorder = run_action("docx", language="de")
    assert recorder.notifications() == [("error", rx.TRANSLATIONS["de"]["export_failed"])]
    assert "boom" in caplog.text


def test_missing_event_call_shows_toast() -> None:
    recorder = run_action("pdf", with_call=False)
    assert recorder.notifications() == [("error", rx.TRANSLATIONS["en"]["export_failed"])]


def test_table_edge_cases() -> None:
    blocks = rx.parse_blocks("Text with | a pipe\n---\n\n| `a\\|b` | c |\n| - | - |\n| 1 |\n")
    assert isinstance(blocks[0], rx.Paragraph) and isinstance(blocks[1], rx.Rule)  # not a table
    table = blocks[2]
    assert table.rows == [["`a|b`", "c"], ["1", ""]]
    assert rx.to_plain(table.rows[0][0]) == "a|b"


def test_filename_helpers() -> None:
    assert rx.safe_filename('a/b\\c:d*e?f"g<h>i|j\nk', "x") == "a b c d e f g h i j k"
    assert rx.safe_filename("  ...  ", "fallback") == "fallback"
    assert rx.safe_filename("x" * 100, "f") == "x" * rx.MAX_FILENAME_CHARS
    assert (
        rx.document_title(rx.parse_blocks("```\n# shell comment\n```\n## Real **title** ##")) == "Real title"
    )


def test_timezone_valve() -> None:
    action = rx.Action()
    action.valves = action.Valves(timezone="Not/AZone")
    assert isinstance(action._now(), datetime)
    action.valves = action.Valves(timezone="Asia/Tokyo")
    assert action._now().utcoffset().total_seconds() == 9 * 3600


# ------------------------------------------------------------------- browser snippets (Node.js) --

node = pytest.mark.skipif(shutil.which("node") is None, reason="Node.js not installed")


def run_in_node(code: str) -> dict[str, Any]:
    result = subprocess.run(
        ["node", str(STUB_JS)], input=code, capture_output=True, text=True, timeout=30, check=True
    )
    return json.loads(result.stdout)


@node
def test_js_download_runs() -> None:
    report = run_in_node(run_action("docx").code())
    assert report["returned"] is True
    assert report["blobs"][0]["type"] == rx.DOCX_MIME and report["blobs"][0]["size"] > 1000
    assert report["clicks"][0]["download"].endswith(".docx")


@node
def test_js_mailto_runs() -> None:
    report = run_in_node(run_action("email", EMAIL).code())
    assert report["clicks"][0]["href"].startswith("mailto:?subject=Meeting%20on%20Friday")


@node
def test_js_print_runs() -> None:
    report = run_in_node(run_action("pdf").code())
    assert report["returned"] is True and report["printed"] is True and report["frameRemoved"] is True
    assert report["srcdoc"].startswith("<!DOCTYPE html>")


# ------------------------------------------------------------------------ optional rendering checks --


def find_browser() -> str | None:
    """Prefer a dedicated headless shell (env PRINT_TEST_BROWSER, PATH, Playwright cache)."""
    candidates = [os.environ.get("PRINT_TEST_BROWSER"), shutil.which("chrome-headless-shell")]
    candidates += sorted(
        map(str, Path.home().glob(".cache/ms-playwright/chromium_headless_shell-*/*/chrome-headless-shell")),
        reverse=True,
    )
    candidates += [shutil.which(name) for name in ("chromium", "google-chrome", "brave")]
    return next((c for c in candidates if c and Path(c).is_file()), None)


BROWSER = find_browser()
HAS_POPPLER = shutil.which("pdftotext") is not None and shutil.which("pdfinfo") is not None


def pdf_text(pdf: Path, layout: bool = False) -> str:
    options = ["-layout"] if layout else []
    return subprocess.run(
        ["pdftotext", *options, str(pdf), "-"], capture_output=True, text=True, check=True
    ).stdout


@pytest.mark.skipif(not (BROWSER and HAS_POPPLER), reason="no headless browser or poppler")
def test_print_page_renders_as_a4_pdf() -> None:
    page = print_page(run_action("pdf", language="de"))
    with tempfile.TemporaryDirectory() as tmp:
        html_file, pdf_file = Path(tmp, "page.html"), Path(tmp, "page.pdf")
        html_file.write_text(page, encoding="utf-8")
        command = [
            BROWSER,
            "--headless=new",
            "--disable-gpu",
            "--no-pdf-header-footer",
            f"--user-data-dir={tmp}/profile",
            f"--print-to-pdf={pdf_file}",
            html_file.as_uri(),
        ]
        try:
            subprocess.run(command, capture_output=True, timeout=90, check=True)
        except subprocess.TimeoutExpired:
            pytest.skip(f"{BROWSER} did not finish printing (headless mode unavailable?)")
        info = subprocess.run(["pdfinfo", str(pdf_file)], capture_output=True, text=True, check=True).stdout
        text = pdf_text(pdf_file)
    assert re.search(r"Page size:.*\(A4\)", info)
    assert re.search(r"Title:\s+Quarterly report C#", info)
    assert "Quarterly report C#" in text and "Erstellt mit dem KI-Assistenten" in text
    assert "SECRET REASONING" not in text
    assert "<script>alert(1)</script>" in "".join(text.split())  # shown as text, not executed


@pytest.mark.skipif(not (shutil.which("soffice") and HAS_POPPLER), reason="LibreOffice or poppler missing")
def test_docx_renders_in_libreoffice() -> None:
    data, _, _ = downloaded(run_action("docx"))
    with tempfile.TemporaryDirectory() as tmp:
        Path(tmp, "report.docx").write_bytes(data)
        subprocess.run(
            [
                "soffice",
                f"-env:UserInstallation=file://{tmp}/profile",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmp,
                str(Path(tmp, "report.docx")),
            ],
            capture_output=True,
            timeout=120,
            check=True,
        )
        text = pdf_text(Path(tmp, "report.pdf"), layout=True)
    lines = [" ".join(line.split()) for line in text.splitlines()]
    assert "3. Third" in lines
    assert "1. Restart" in lines and "2. Again" in lines  # second list restarts at 1


# --------------------------------------------------------------------------------- context guard --


def body_with_chars(chars: int, model: str = "m") -> dict[str, Any]:
    return {"model": model, "messages": [{"role": "user", "content": "x" * chars}]}


def run_guard(
    flt: Any, body: dict[str, Any], metadata: dict[str, Any] | None = None
) -> tuple[dict[str, Any], Recorder]:
    recorder = Recorder()
    result = asyncio.run(flt.request(body, __event_emitter__=recorder.emit, __metadata__=metadata))
    return result, recorder


def test_guard_ok() -> None:
    body = body_with_chars(1000)
    result, recorder = run_guard(make_filter(cg, context_tokens=10000, answer_reserve_tokens=2000), body)
    assert result is body and result == body_with_chars(1000)
    assert recorder.events == []


@pytest.mark.parametrize(
    ("language", "text"), [("en", "very long (about 80%"), ("de", "sehr lang (ca. 80 %")]
)
def test_guard_warns(language: str, text: str) -> None:
    # usable = 8000 tokens; (22384 + 16) chars / 3.5 = 6400 tokens = 80 %
    flt = make_filter(cg, context_tokens=10000, answer_reserve_tokens=2000, language=language)
    _, recorder = run_guard(flt, body_with_chars(22384))
    assert len(recorder.events) == 1 and recorder.events[0]["type"] == "status"
    assert text in recorder.events[0]["data"]["description"] and recorder.events[0]["data"]["done"] is True


@pytest.mark.parametrize(
    ("language", "text"), [("en", "too large for the AI model (about 101%"), ("de", "ca. 101 %")]
)
def test_guard_blocks(language: str, text: str) -> None:
    flt = make_filter(cg, context_tokens=10000, answer_reserve_tokens=2000, language=language)
    with pytest.raises(cg.ContextLimitError) as error:
        run_guard(flt, body_with_chars(8001 * 35 // 10))
    assert text in str(error.value)


def test_guard_counts_parts_images_and_tools() -> None:
    body = {
        "model": "m",
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "x" * 349},
                    {"type": "image_url", "image_url": {"url": "data:..."}},
                    "not-a-dict",
                ],
            },
            {"role": "assistant", "content": None, "tool_calls": [{"id": "1", "function": {"name": "f"}}]},
        ],
        "tools": [{"type": "function", "function": {"name": "f"}}],
    }
    tool_chars = len(json.dumps(body["tools"])) + len(json.dumps(body["messages"][1]["tool_calls"][0]))
    expected = -(-(349 + 2 * cg.MESSAGE_OVERHEAD_CHARS + tool_chars) * 10 // 35) + 300
    assert cg.estimate_tokens(body, 3.5, 300) == expected


def test_guard_per_model_override_and_misconfiguration(caplog: pytest.LogCaptureFixture) -> None:
    flt = make_filter(
        cg,
        context_tokens=100000,
        answer_reserve_tokens=100,
        model_context_tokens="small:latest=1000, broken, x=abc",
    )
    with pytest.raises(cg.ContextLimitError):
        run_guard(flt, body_with_chars(10000, model="small:latest"))
    run_guard(flt, body_with_chars(10000, model="other"))
    assert "broken" in caplog.text and "x=abc" in caplog.text
    body = body_with_chars(10**6)
    assert run_guard(make_filter(cg, context_tokens=1000, answer_reserve_tokens=1000), body)[0] is body


def stub_compaction_config(monkeypatch: pytest.MonkeyPatch, values: dict[str, Any]) -> None:
    """Stands in for open_webui.models.config.Config, which holds the Context Compaction settings."""

    async def get_many(*keys: str) -> dict[str, Any]:
        return {key: values[key] for key in keys if key in values}

    modules = {
        "open_webui": ModuleType("open_webui"),
        "open_webui.models": ModuleType("open_webui.models"),
        "open_webui.models.config": SimpleNamespace(Config=SimpleNamespace(get_many=get_many)),
    }
    for name, module in modules.items():
        monkeypatch.setitem(sys.modules, name, module)


def test_guard_compaction_hides_hint_for_normal_growth(monkeypatch: pytest.MonkeyPatch) -> None:
    # usable = 8000 tokens, compaction from 7000: at 80 % the chat is summarized before the limit
    stub_compaction_config(
        monkeypatch,
        {"chat.context_compaction.enable": True, "chat.context_compaction.token_threshold": 7000},
    )
    flt = make_filter(cg, context_tokens=10000, answer_reserve_tokens=2000)
    assert run_guard(flt, body_with_chars(22384))[1].events == []
    # still above the compaction threshold after compaction (e.g. large documents): hint is shown
    _, recorder = run_guard(flt, body_with_chars(7200 * 35 // 10))
    assert len(recorder.events) == 1 and "very long" in recorder.events[0]["data"]["description"]
    with pytest.raises(cg.ContextLimitError):
        run_guard(flt, body_with_chars(8001 * 35 // 10))


def test_guard_compaction_threshold_too_high(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    # default threshold 80000 is above the usable 8000 tokens: compaction never runs in time
    stub_compaction_config(monkeypatch, {"chat.context_compaction.enable": True})
    flt = make_filter(cg, context_tokens=10000, answer_reserve_tokens=2000)
    run_guard(flt, body_with_chars(1000))
    assert "Context Compaction threshold" not in caplog.text  # only checked once a chat gets long
    assert len(run_guard(flt, body_with_chars(22384))[1].events) == 1
    with pytest.raises(cg.ContextLimitError):
        run_guard(flt, body_with_chars(8001 * 35 // 10))
    assert caplog.text.count("Context Compaction threshold for model 'm' (80000 tokens)") == 1


def test_guard_compaction_threshold_resolution(monkeypatch: pytest.MonkeyPatch) -> None:
    per_model = {"params": {"compact_token_threshold": 5000}}
    assert asyncio.run(cg.compaction_threshold(per_model)) is None  # outside Open WebUI
    stub_compaction_config(monkeypatch, {"chat.context_compaction.enable": False})
    assert asyncio.run(cg.compaction_threshold(per_model)) is None
    settings = {"chat.context_compaction.enable": True, "chat.context_compaction.token_threshold": "20000"}
    stub_compaction_config(monkeypatch, settings)
    assert asyncio.run(cg.compaction_threshold(None)) == 20000
    assert asyncio.run(cg.compaction_threshold(per_model)) == 5000
    assert asyncio.run(cg.compaction_threshold({"params": {"compact_token_threshold": 30000}})) == 20000
    stub_compaction_config(monkeypatch, {**settings, "chat.context_compaction.token_cap": 40000})
    assert asyncio.run(cg.compaction_threshold({"params": {"compact_token_threshold": 30000}})) == 30000


# ------------------------------------------------------------------------------- company context --

MONDAY = datetime(2026, 10, 12, 14, 5, tzinfo=ZoneInfo("Europe/Berlin"))


def test_note_english() -> None:
    flt = make_filter(cc, company_name="Example Ltd", timezone="Europe/Berlin")
    note = flt.build_note({"name": "Ada\nLovelace"}, MONDAY)
    assert note.splitlines() == [
        "[Company context]",
        "You are the internal AI assistant of Example Ltd.",
        "Today is Monday, 12 October 2026.",
        (
            "Always answer in English unless the user writes in another language "
            "or explicitly asks for a different one."
        ),
        cc.TRANSLATIONS["en"]["formal"],
        "The user's name is Ada Lovelace.",
    ]


def test_note_german_informal_with_rules() -> None:
    flt = make_filter(
        cc,
        language="de",
        company_name="Musterfirma GmbH",
        form_of_address="informal",
        include_user_name=False,
        extra_rules="Wir schreiben Kunden gegenüber in der Wir-Form.",
    )
    note = flt.build_note({"name": "Max"}, MONDAY)
    assert note.splitlines()[:3] == [
        "[Firmen-Kontext]",
        "Du bist der interne KI-Assistent der Firma Musterfirma GmbH.",
        "Heute ist Montag, 12. Oktober 2026.",
    ]
    assert "Antworte immer auf Deutsch, außer" in note
    assert "„du“" in note and "Du-Form" in note
    assert "Max" not in note
    assert note.endswith("Wir schreiben Kunden gegenüber in der Wir-Form.")


def test_note_without_company_and_custom_response_language() -> None:
    note = make_filter(cc, response_language="French").build_note(None, MONDAY)
    assert "internal AI assistant" not in note
    assert "Always answer in French unless" in note


def run_inlet(flt: Any, body: dict[str, Any]) -> dict[str, Any]:
    return asyncio.run(flt.inlet(body, __user__={"name": "Ada"}))


def test_inlet_adds_system_message_and_is_idempotent() -> None:
    flt = make_filter(cc, timezone="Europe/Berlin")
    body = run_inlet(flt, {"messages": [{"role": "user", "content": "Hi"}]})
    body = run_inlet(flt, body)
    assert [m["role"] for m in body["messages"]] == ["system", "user"]
    assert body["messages"][0]["content"].count("[Company context]") == 1
    assert re.search(r"^Today is \w+, \d{1,2} \w+ \d{4}\.$", body["messages"][0]["content"], re.MULTILINE)


def test_inlet_prepends_to_existing_system_prompt() -> None:
    flt = make_filter(cc, language="de")
    body = {
        "messages": [{"role": "system", "content": "You are a pirate."}, {"role": "user", "content": "Hi"}]
    }
    body = run_inlet(flt, run_inlet(flt, body))
    system = body["messages"][0]["content"]
    assert system.startswith("[Firmen-Kontext]\n") and system.endswith("\n\nYou are a pirate.")
    assert system.count("[Firmen-Kontext]") == 1 and len(body["messages"]) == 2


def test_inlet_handles_list_system_content_and_missing_messages() -> None:
    flt = make_filter(cc)
    body = run_inlet(flt, {"messages": [{"role": "system", "content": [{"type": "text", "text": "Base"}]}]})
    body = run_inlet(flt, body)
    parts = body["messages"][0]["content"]
    assert len(parts) == 2 and parts[0]["text"].startswith("[Company context]") and parts[1]["text"] == "Base"
    assert run_inlet(flt, {})["messages"][0]["role"] == "system"


def test_invalid_timezone_falls_back(caplog: pytest.LogCaptureFixture) -> None:
    body = run_inlet(make_filter(cc, timezone="Mars/Olympus"), {"messages": []})
    assert "Today is" in body["messages"][0]["content"]
    assert "Mars/Olympus" in caplog.text


# ------------------------------------------------------------------------------ reasoning toggle --


def run_toggle(flt: Any, body: dict[str, Any], model: dict[str, Any] | None = None) -> dict[str, Any]:
    return asyncio.run(flt.inlet(body, __model__=model))


def test_toggle_button() -> None:
    flt = rt.Filter()
    assert flt.toggle is True
    prefix = "data:image/svg+xml;base64,"
    assert flt.icon.startswith(prefix)
    assert base64.b64decode(flt.icon[len(prefix) :]).decode() == rt.ICON_SVG


def test_toggle_enables_thinking_and_keeps_other_kwargs() -> None:
    body = {"model": "m", "messages": [], "chat_template_kwargs": {"enable_thinking": False, "custom": 1}}
    result = run_toggle(make_filter(rt), body)
    assert result == {
        "model": "m",
        "messages": [],
        "chat_template_kwargs": {"enable_thinking": True, "custom": 1},
    }
    assert run_toggle(make_filter(rt), {"model": "m", "chat_template_kwargs": "broken"})[
        "chat_template_kwargs"
    ] == {"enable_thinking": True}


@pytest.mark.parametrize(("default", "expected"), [("none", None), ("low", "low")])
def test_toggle_clears_only_a_none_default(default: str, expected: str | None) -> None:
    body = run_toggle(make_filter(rt), {"model": "m", "reasoning_effort": default})
    assert "reasoning_effort" in body and body["reasoning_effort"] == expected


def test_toggle_per_model_effort_and_budget(caplog: pytest.LogCaptureFixture) -> None:
    flt = make_filter(
        rt,
        model_reasoning_effort="qwen=XHigh, broken, gemma=none, x=ultra",
        thinking_budget=2000,
        model_thinking_budget="qwen=8000, gemma=0, y=abc",
    )
    body = run_toggle(flt, {"model": "qwen", "reasoning_effort": "none"})
    assert body["reasoning_effort"] == "xhigh"
    assert body["thinking_budget_tokens"] == body["thinking_token_budget"] == 8000
    # workspace model: settings of its base model apply
    workspace = {"id": "emails", "info": {"base_model_id": "qwen"}}
    assert run_toggle(flt, {"model": "emails"}, workspace)["reasoning_effort"] == "xhigh"
    # per-model budget 0 = no limit, despite the global budget
    body = run_toggle(flt, {"model": "gemma"})
    assert not {"reasoning_effort", "thinking_budget_tokens", "thinking_token_budget"} & body.keys()
    body = run_toggle(flt, {"model": "other"})
    assert body["thinking_budget_tokens"] == 2000 and "reasoning_effort" not in body
    for entry in ("broken", "gemma=none", "x=ultra", "y=abc"):
        assert repr(entry) in caplog.text


def test_toggle_model_ids() -> None:
    assert rt.model_ids({"model": "a"}, {"id": "a", "info": {"base_model_id": "b"}}) == ["a", "b"]
    assert rt.model_ids({}, None) == []
    assert rt.model_ids({"model": "a"}, {"info": None}) == ["a"]
