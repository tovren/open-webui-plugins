"""Tests for the Office Documents tool (office-documents/office_documents.py).

Run from anywhere (Python 3.11+):
    pip install python-docx openpyxl python-pptx pydantic pytest
    python -I -m pytest -q path/to/tests/test_office_documents.py

The optional LibreOffice check (renders all three formats to PDF) runs when `soffice` and
`pdftotext` are installed.
"""

import ast
import asyncio
import base64
import importlib.util
import inspect
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from types import ModuleType, SimpleNamespace
from typing import Any

import pytest
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

PLUGIN = Path(__file__).resolve().parent.parent / "office-documents" / "office_documents.py"
ALLOWED_THIRD_PARTY = {"docx", "openpyxl", "pptx", "pydantic"}
LAZY_IMPORTS = {"open_webui", "starlette"}  # only imported inside attach_file, when running in Open WebUI
FRONTMATTER_KEYS = {"title", "author", "author_url", "version", "license", "required_open_webui_version", "description"}
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def load_plugin() -> ModuleType:
    spec = importlib.util.spec_from_file_location("plugin_office_documents", PLUGIN)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


od = load_plugin()


class Recorder:
    """Collects the events the tool sends to Open WebUI."""

    def __init__(self) -> None:
        self.events: list[dict[str, Any]] = []
        self.calls: list[dict[str, Any]] = []

    async def emit(self, event: dict[str, Any]) -> None:
        self.events.append(event)

    async def call(self, event: dict[str, Any]) -> bool:
        self.calls.append(event)
        return True

    def download(self) -> tuple[bytes, str, str]:
        assert len(self.calls) == 1, self.calls
        code = self.calls[0]["data"]["code"]
        data = base64.b64decode(js_value(code, "atob("))
        return data, js_value(code, "link.download = "), js_value(code, "{ type: ")

    def statuses(self) -> list[tuple[str, bool]]:
        return [(e["data"]["description"], e["data"]["done"]) for e in self.events if e["type"] == "status"]


def js_value(code: str, prefix: str) -> Any:
    start = code.index(prefix) + len(prefix)
    return json.JSONDecoder().raw_decode(code, start)[0]


def close_brackets(text: str) -> str:
    return od._close_brackets(text)[0]


def make_tools(**valves: Any) -> Any:
    tools = od.Tools()
    tools.valves = tools.Valves(**{"delivery": "download", **valves})
    return tools


def call(method: str, valves: dict[str, Any] | None = None, **kwargs: Any) -> tuple[str, Recorder]:
    tools = make_tools(**(valves or {}))
    recorder = Recorder()
    result = asyncio.run(
        getattr(tools, method)(**kwargs, __event_emitter__=recorder.emit, __event_call__=recorder.call)
    )
    return result, recorder


def word(**kwargs: Any) -> tuple[Any, str, str]:
    result, recorder = call("create_word_file", **kwargs)
    assert result.startswith("Done"), result
    data, filename, mime = recorder.download()
    assert mime == od.DOCX_MIME
    return Document(io.BytesIO(data)), filename, result


def excel(sheets: list[dict[str, Any]], **kwargs: Any) -> tuple[Any, str]:
    result, recorder = call("create_excel_file", sheets=sheets, **kwargs)
    assert result.startswith("Done"), result
    data, filename, mime = recorder.download()
    assert mime == od.XLSX_MIME
    return load_workbook(io.BytesIO(data)), filename


def powerpoint(slides: list[dict[str, Any]], **kwargs: Any) -> tuple[Any, str]:
    result, recorder = call("create_powerpoint_file", slides=slides, **kwargs)
    assert result.startswith("Done"), result
    data, filename, mime = recorder.download()
    assert mime == od.PPTX_MIME
    return Presentation(io.BytesIO(data)), filename


def invalid(method: str, **kwargs: Any) -> str:
    result, recorder = call(method, **kwargs)
    assert result.startswith("The file was NOT created"), result
    assert not recorder.calls
    return result


# ------------------------------------------------------------------------------------------- plugin --


def test_frontmatter_and_bundled_libraries() -> None:
    source = (PLUGIN).read_text(encoding="utf-8")
    docstring = ast.get_docstring(ast.parse(source)) or ""
    meta = dict(line.split(": ", 1) for line in docstring.split("\n\n")[0].splitlines())
    assert set(meta) == FRONTMATTER_KEYS
    top_level = set()
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.Import):
            top_level |= {alias.name.split(".")[0] for alias in node.names}
        elif isinstance(node, ast.ImportFrom) and node.module:
            top_level.add(node.module.split(".")[0])
    assert top_level <= set(sys.stdlib_module_names) | ALLOWED_THIRD_PARTY | LAZY_IMPORTS


def test_tool_signatures() -> None:
    """Open WebUI exposes every public method; injected (__x__) parameters need defaults."""
    tools = od.Tools()
    public = sorted(
        name
        for name in dir(tools)
        if not name.startswith("_") and callable(getattr(tools, name)) and not inspect.isclass(getattr(tools, name))
    )
    assert public == ["create_excel_file", "create_powerpoint_file", "create_word_file"]
    for name in public:
        method = getattr(tools, name)
        params = inspect.signature(method).parameters
        documented = set(re.findall(r":param (\w+):", method.__doc__))
        assert documented == {p for p in params if not p.startswith("__")}
        assert all(p.default is not inspect.Parameter.empty for n, p in params.items() if n.startswith("__"))
        assert "only exists once this tool has answered" in method.__doc__


def test_schema_has_no_size_limits() -> None:
    """maxItems/maxLength would blow up the grammar of llama.cpp/vLLM; limits are checked in code."""
    for adapter in (od.WORD_BLOCKS, od.EXCEL_SHEETS, od.SLIDES):
        schema = json.dumps(adapter.json_schema())
        assert not re.search(r'"(maxItems|maxLength|minItems|minLength|maximum|minimum)"', schema)


# --------------------------------------------------------------------------------------------- Word --

BLOCKS = [
    {"type": "heading", "text": "Introduction", "level": 1},
    {"type": "paragraph", "text": "Some **bold** and *italic* text with a [link](https://example.org) and [bad](javascript:alert(1)).\n\nSecond paragraph."},
    {"type": "bullets", "items": ["One", "- One a", "Two"]},
    {"type": "numbered", "items": ["First", "Second"]},
    {"type": "heading", "text": "Numbers", "level": 5},
    {"type": "table", "rows": [["Item", "Amount"], ["Rent", 1200], ["short"]]},
    {"type": "quote", "text": "A quote"},
    {"type": "numbered", "items": ["Again"]},
    {"type": "page_break"},
    {"type": "paragraph", "text": "End"},
]  # fmt: skip


def test_word_document_structure() -> None:
    doc, filename, _ = word(blocks=BLOCKS, title="Quarterly report", subtitle="Q3", header_text="Example Ltd")
    assert filename == "Quarterly report.docx"
    paragraphs = [(p.style.name, p.text) for p in doc.paragraphs]
    assert paragraphs[0] == ("Title", "Quarterly report")
    assert paragraphs[1] == ("Subtitle", "Q3")
    assert ("Heading 1", "Introduction") in paragraphs
    assert ("Heading 3", "Numbers") in paragraphs  # levels above 3 are capped
    assert ("List Bullet 2", "One a") in paragraphs
    assert ("Quote", "A quote") in paragraphs
    assert ("Normal", "Second paragraph.") in paragraphs
    text = next(p for p in doc.paragraphs if p.text.startswith("Some bold"))
    assert [r.text for r in text.runs if r.bold] == ["bold"]
    assert [r.text for r in text.runs if r.italic] == ["italic"]
    links = [rel.target_ref for rel in doc.part.rels.values() if rel.reltype.endswith("/hyperlink")]
    assert links == ["https://example.org"]  # javascript: stays plain text
    assert "bad" in text.text and "javascript" not in text.text
    table = doc.tables[0]
    assert [[c.text for c in row.cells] for row in table.rows] == [["Item", "Amount"], ["Rent", "1200"], ["short", ""]]
    assert table.rows[0].cells[0].paragraphs[0].runs[0].bold
    assert table.rows[0]._tr.trPr.find(f"{W_NS}tblHeader") is not None
    assert doc.sections[0].header.paragraphs[0].text == "Example Ltd"
    footer_xml = doc.sections[0].footer._element.xml
    assert "PAGE" in footer_xml and "NUMPAGES" in footer_xml and "Page " in footer_xml
    num_ids = [p._p.pPr.numPr.numId.val for p in doc.paragraphs if p.style.name == "List Number"]
    assert len(set(num_ids)) == 2  # the second numbered list restarts


def test_word_table_column_widths() -> None:
    rows = [["#", "Risk", "P", "Measure"], ["R10", "Supplier fails", "3", "Second supplier under contract " * 3, ""]]
    doc, _, _ = word(blocks=[{"type": "table", "rows": rows}], title="T")
    table = doc.tables[0]
    assert len(table.columns) == 4  # the trailing empty cell adds no column
    widths = [column.width for column in table.columns]
    section = doc.sections[0]
    assert sum(widths) == section.page_width - section.left_margin - section.right_margin
    assert widths[2] < widths[0] < widths[1] < widths[3]
    assert widths[0] >= 3 * od.CHAR_WIDTH + od.CELL_PADDING  # "R10" stays on one line
    assert table.rows[1].cells[3].width == widths[3]
    assert od.column_widths([["a", "b"]], 2, 1000) == [500, 500]
    assert sum(od.column_widths([["x" * 40] * 12], 12, 1000)) == 1000  # words that cannot fit are shrunk


def test_word_lists_inside_paragraphs() -> None:
    text = "Problems:\n* **Lost** requests\n  - in mail\n* No overview\nSo we need:\n1. A tool\n2. Rules\n\nEnd."
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": text}], title="T")
    assert [(p.style.name, p.text) for p in doc.paragraphs[1:]] == [
        ("Normal", "Problems:"),
        ("List Bullet", "Lost requests"),
        ("List Bullet 2", "in mail"),
        ("List Bullet", "No overview"),
        ("Normal", "So we need:"),
        ("List Number", "A tool"),
        ("List Number", "Rules"),
        ("Normal", "End."),
    ]


def test_word_toc_landscape_and_german() -> None:
    result, recorder = call(
        "create_word_file",
        valves={"language": "de", "page_numbers": False},
        blocks=BLOCKS,
        title="Bericht",
        table_of_contents=True,
        landscape=True,
    )
    doc = Document(io.BytesIO(recorder.download()[0]))
    assert "TOC" in doc.element.body.xml and "Inhalt" in [p.text for p in doc.paragraphs]
    assert doc.settings.element.find(f"{W_NS}updateFields") is not None
    section = doc.sections[0]
    assert section.page_width > section.page_height
    assert "PAGE" not in section.footer._element.xml
    assert recorder.statuses() == [("Word-Dokument wird erstellt …", False), ("Bericht.docx erstellt", True)]


@pytest.mark.parametrize(
    ("title", "first_heading", "expected_title", "heading_kept"),
    [
        ("", "Minutes of the meeting", "Minutes of the meeting", False),
        ("Minutes of the meeting", "Minutes of the meeting", "Minutes of the meeting", False),
        ("Minutes_2026_10_12.docx", "Minutes of the meeting", "Minutes of the meeting", False),
        ("Report_Q3", None, "Report Q3", None),
        ("Report", "Introduction", "Report", True),
    ],
)
def test_word_title(title: str, first_heading: str | None, expected_title: str, heading_kept: bool | None) -> None:
    blocks = ([{"type": "heading", "text": first_heading, "level": 1}] if first_heading else []) + [
        {"type": "paragraph", "text": "Body"}
    ]
    doc, filename, _ = word(blocks=blocks, title=title)
    texts = [p.text for p in doc.paragraphs]
    assert texts[0] == expected_title and filename == f"{expected_title}.docx"
    if first_heading:
        assert (first_heading in texts[1:]) == heading_kept


def test_shapes_outside_the_schema_are_rejected() -> None:
    """Other shapes are not guessed: the error names the field, so the model sends it again in the right shape."""
    doc, _, _ = word(blocks=json.dumps([{"type": "paragraph", "text": "x"}]), title="T")  # sent as a JSON string
    assert doc.paragraphs[1].text == "x"
    result = invalid("create_word_file", blocks=[{"heading": "Intro"}, {"type": "list", "items": ["a"]}])
    assert "blocks[0]: 'heading' is not a field here, its content would be lost (fields: type, text" in result
    assert "blocks[1].type: Input should be 'heading'" in result
    result = invalid("create_excel_file", sheets=[{"name": "Data", "headers": ["Month"], "rows": [["Jan"]]}])
    assert "sheets[0]: 'headers' is not a field here" in result and "(fields: name, columns, rows, chart)" in result
    slides = [{"layout": "two_columns", "left_heading": "Pros", "right_heading": "Cons", "left": ["a"], "right": ["b"]}]
    assert "slides[0].title: Field required" in invalid("create_powerpoint_file", slides=slides)
    slides = [{"layout": "bullets", "title": "Next steps", "bullets": []}]
    assert "slides[0] (bullets): 'bullets' is empty" in invalid("create_powerpoint_file", slides=slides)


def test_word_invalid_input() -> None:
    result = invalid(
        "create_word_file", blocks=[{"type": "paragraph", "text": "x"}, {"type": "table", "rows": [["x"] * 13]}]
    )
    assert "blocks[1] (table): at most 12 columns" in result
    assert "call create_word_file again" in result
    assert "blocks[1].type: Input should be" in invalid("create_word_file", blocks=[{}, {"type": "unknown"}])
    assert "the document has no content" in invalid("create_word_file", blocks=[])
    assert "the document has no content" in invalid("create_word_file", blocks=[{"type": "paragraph", "text": " "}])


def test_word_empty_blocks_and_toc_block() -> None:
    blocks = [
        {"type": "paragraph", "text": "Intro"},
        {"type": "paragraph", "text": ""},  # spacing, left out
        {"type": "heading", "text": " "},
        {"type": "table_of_contents"},
        {"type": "bullets", "items": ["a"]},
        {"type": "numbered", "items": ["b"]},
        {"type": "heading", "text": "Part", "level": 2},
        {"type": "paragraph", "text": "Body " * 25},
    ]
    doc, _, _ = word(blocks=blocks, title="T", table_of_contents=True)
    paragraphs = [(p.style.name, p.text) for p in doc.paragraphs if p.text]
    texts = [p.text for p in doc.paragraphs]
    assert texts[:3] == ["T", "Intro", "Contents"]  # empty paragraph and heading left out, contents where the block was
    assert (
        ("List Bullet", "a") in paragraphs
        and ("List Number", "b") in paragraphs
        and ("Heading 2", "Part") in paragraphs
    )
    assert doc.element.body.xml.count("TOC \\o") == 1  # not twice, although table_of_contents=True was also set


def test_word_invalid_json() -> None:
    assert "the JSON is invalid at character 2" in invalid("create_word_file", blocks="[{oops")


def test_json_text_is_repaired() -> None:
    """Long arguments arrive as JSON text; German quotes and a forgotten '}' are repaired."""
    german = '[{"type": "paragraph", "text": "Priorität „Hoch" heißt: sofort."}]'
    assert od.loads_lenient(german)[0]["text"] == "Priorität „Hoch“ heißt: sofort."
    missing_brace = (
        '[{"layout": "chart", "title": "C", "chart": {"categories": ["a"], "series": [{"name": "s", "values": [1]}],'
        ' "notes": "n"}, {"layout": "bullets", "title": "B", "bullets": ["x"]}]'
    )
    slides = od.parse_argument(od.SLIDES, missing_brace, "slides")
    assert [s.layout for s in slides] == ["chart", "bullets"] and slides[0].notes == "n"
    assert od.loads_lenient('{"a": "„x\\" ok"}') == {"a": '„x" ok'}  # valid JSON is never changed
    # a forgotten "}" before "]", followed by other tool parameters written into the same text
    forgotten = '[{"type": "bullets", "items": ["a", "b"]], "title": "T", "landscape": false}'
    extras: dict[str, Any] = {}
    assert od.loads_lenient(forgotten, extras) == [{"type": "bullets", "items": ["a", "b"]}]
    assert extras == {"title": "T", "landscape": False}
    assert od.loads_lenient('[{"rows": [["a", "]"]]') == [
        {"rows": [["a", "]"]]}
    ]  # last "]" forgotten; "]" in text kept
    assert close_brackets('{"a": ["x", {"b": 1]}') == '{"a": ["x", {"b": 1}]}'
    # table rows and list items written as several arrays without the outer brackets
    run = '[{"layout": "table", "title": "T", "table": ["A", "B"], ["1", "2"], ["3", "4"], "notes": "n"}]'
    assert od.loads_lenient(run)[0]["table"] == [["A", "B"], ["1", "2"], ["3", "4"]]
    assert od.wrap_array_run('{"a": [1], "b": 2}', 9) is None  # nothing to wrap
    result = invalid("create_word_file", blocks='[{"type": "paragraph", "text": "a" "b"}]')
    assert "at character 35 (Expecting ',' delimiter): [{\"type\"" in result
    assert '"text": "a" <<<HERE>>> "b"}]' in result
    assert 'the text before the mark ends with a " that closes it too early' in result


def test_content_is_never_dropped_silently() -> None:
    """Unknown keys are ignored by the schema, so content under another name is sent back instead of being lost."""
    assert "blocks[0]: 'content' is not a field here, its content would be lost" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "content": "Important"}, {"type": "paragraph", "text": "x"}]
    )
    # the retry after "type: Field required": the type is added, the old key kept
    assert "blocks[0]: 'heading' is not a field here" in invalid(
        "create_word_file", blocks=[{"type": "heading", "heading": "Einleitung"}, {"type": "paragraph", "text": "x"}]
    )
    assert "a paragraph block shows only 'text'; put the 'items' into a block of its own type" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "items": ["a", "b"]}]
    )
    assert "slides[0]: 'text' is not a field here" in invalid(
        "create_powerpoint_file", slides=[{"layout": "title", "title": "Deck", "text": "Subtitle"}]
    )
    result = invalid("create_powerpoint_file", slides=[{"layout": "section", "title": "Agenda", "bullets": ["a"]}])
    assert "slides[0] (section): 'bullets' is not shown on a section slide; use layout bullets for it" in result
    sheet = {"name": "S", "columns": [{"header": "A"}, {"header": "B", "type": "number"}], "rows": [["x", 1]],
             "total": {"B": "sum"}}  # fmt: skip
    assert "sheets[0]: 'total' is not a field here" in invalid("create_excel_file", sheets=[sheet])
    column = {"header": "B", "type": "number", "formulas": "=A{row}*2"}
    sheet = {"name": "S", "columns": [{"header": "A", "type": "number"}, column], "rows": [[1, None]]}
    assert "sheets[0].columns[1]: 'formulas' is not a field here" in invalid("create_excel_file", sheets=[sheet])
    assert "blocks[0]: 'data' is not a field here" in invalid(
        "create_word_file",
        blocks=[{"type": "table", "data": [[2024, 1], [2025, 2]]}],  # numbers in a list are content
    )
    assert "a page_break block shows no 'text'" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "text": "x"}, {"type": "page_break", "text": "Kapitel 2"}]
    )
    # harmless: the slide title is shown above the chart, numbers and empty values hold no text
    chart = {"title": "Umsatz", "categories": ["a"], "series": [{"name": "s", "values": [1]}]}
    powerpoint([{"layout": "chart", "title": "T", "chart": chart, "number": 3, "style": ""}], title="D")


def test_fields_written_into_the_chart_are_taken_back() -> None:
    chart = {"categories": ["a"], "series": [{"name": "s", "values": [1]}], "notes": "n"}
    (slide,) = od.parse_argument(od.SLIDES, [{"layout": "chart", "title": "T", "notes": None, "chart": chart}], "s")
    assert slide.notes == "n" and slide.chart is not None
    sheet = {
        "name": "S",
        "columns": [{"header": "A"}, {"header": "B", "type": "number"}],
        "chart": {"type": "line", "category_column": "A", "value_columns": ["B"], "rows": [["x", 1], ["y", 2]]},
    }
    ws = excel([sheet])[0]["S"]  # fmt: skip
    assert ws["B3"].value == 2 and len(ws._charts) == 1
    empty = {**MONTHS_AS_COLUMNS, "chart": {"category_column": None, "value_columns": []}}
    assert not excel([empty])[0]["Budget"]._charts  # an empty chart object means "no chart"
    slide = {
        "layout": "bullets",
        "title": "B",
        "bullets": ["a"],
        "chart": {"type": "column", "categories": [], "series": []},
    }
    powerpoint([slide], title="D")  # a slide chart without data is no chart, so nothing is lost


def test_total_rows_written_into_the_rows() -> None:
    columns = [{"header": "Pos"}, {"header": "Menge", "type": "number"}, {"header": "Preis", "type": "currency"},
               {"header": "Gesamt", "type": "currency", "formula": "=B{row}*C{row}"}]  # fmt: skip
    own_numbers = {"name": "A", "columns": columns, "rows": [["a", 1, 2], ["b", 3, 4], ["Summe", 4, 6, None]]}
    assert "rows[2] ('Summe') is a total row" in invalid("create_excel_file", sheets=[own_numbers])  # D4 = B4*C4
    sums = {"name": "B", "columns": columns[:3], "rows": [["a", 1, 2], ["Summe netto:", "=SUM(B2:B2)", "=SUM(C2:C2)"]]}
    assert excel([sums])[0]["B"]["B3"].value == "=SUM(B2:B2)"  # the model's own sums, written as formulas
    # labels that only start like a total, and sheets without numbers, are left alone
    titles = {
        "name": "C",
        "columns": [{"header": "Kurs"}],
        "rows": [["Excel Grundlagen"], ["Total Quality Management"]],
    }
    tasks = {"name": "D", "columns": [{"header": "Task"}, {"header": "Owner"}, {"header": "Done"}],
             "rows": [["Draft", "Anna", "yes"], ["Sum up results", None, None]]}  # fmt: skip
    workbook = excel([titles, tasks])[0]
    assert workbook["C"]["A3"].value == "Total Quality Management" and workbook["D"]["A3"].value == "Sum up results"
    total_too = {"name": "E", "columns": [{"header": "Posten"}, {"header": "Betrag", "type": "currency", "total": "sum"}],
                 "rows": [["Miete", 900], ["Strom", 100], ["Gesamt netto", 1000]]}  # fmt: skip
    assert "rows[2] ('Gesamt netto') is a total row" in invalid("create_excel_file", sheets=[total_too])  # twice
    no_columns = invalid("create_excel_file", sheets=[{"name": "S", "columns": [], "rows": [[], []]}])
    assert "the sheet has no columns" in no_columns


def test_null_written_as_text_in_number_columns() -> None:
    sheet = {"name": "S", "columns": [{"header": "A"}, {"header": "B"}, {"header": "C", "type": "number"}],
             "rows": [["x", "1.200", "null"], ["y", "null", 2], ["z", "3", None]]}  # fmt: skip
    ws = excel([sheet], valves={"language": "de"})[0]["S"]
    assert [ws[f"B{r}"].value for r in (2, 3, 4)] == [1200, None, 3] and ws["C2"].value is None
    # all 'null': the checks for empty columns see empty cells, not text
    offer = {"name": "O", "columns": [{"header": "Pos"}, {"header": "Menge", "type": "number"},
             {"header": "Preis", "type": "currency"}, {"header": "Gesamt", "type": "currency", "formula": "=B{row}*C{row}"}],
             "rows": [["a", 1, "null", None], ["b", 2, "null", None]]}  # fmt: skip
    assert "columns[2] ('Preis') is empty, but the formula of 'Gesamt' uses it" in invalid(
        "create_excel_file", sheets=[offer]
    )
    summe = {"name": "P", "columns": [{"header": "A"}, {"header": "B", "type": "number"}],
             "rows": [["x", 1], ["Summe", "null"]]}  # fmt: skip
    assert "rows[1] ('Summe') is a total row" in invalid("create_excel_file", sheets=[summe])
    text = {"name": "T", "columns": [{"header": "Note", "type": "text"}], "rows": [["None"]]}
    assert excel([text])[0]["T"]["A2"].value == "None"  # in a text column it may be a word


def test_items_that_all_start_with_a_dash_and_html_breaks() -> None:
    doc, _, _ = word(blocks=[{"type": "bullets", "items": ["- a", "- b"]}], title="T")
    assert [p.style.name for p in doc.paragraphs if p.text in ("a", "b")] == ["List Bullet", "List Bullet"]
    slide = {"layout": "two_columns", "title": "V", "left": ["- fast", "- cheap"], "right": ["- slow<BR/>loud"]}
    (parsed,) = od.parse_argument(od.SLIDES, [slide], "slides")
    assert parsed.left == ["fast", "cheap"] and parsed.right == ["slow\nloud"]


def test_json_error_messages_point_to_the_slip() -> None:
    """Shown on both sides of the place, with the likely slip: models repeat a mistake they cannot find."""

    def message(text: str) -> str:
        return invalid("create_word_file", blocks=text)

    def unrepaired(text: str) -> str:  # the message for a slip the repair would fix
        try:
            json.loads(text)
        except json.JSONDecodeError as e:
            return od.json_error_message("blocks", e)
        raise AssertionError("valid JSON")

    extra = unrepaired('[{"type": "numbered", "items": ["a", "b"]]}, {"type": "paragraph", "text": "x"}]')
    assert '"items": ["a", "b"] <<<HERE>>> ]}, {"type"' in extra
    assert "the open object must be closed with } first, or the ] is one too many" in extra
    assert "the } at the mark closes nothing that is open" in message('[{"type": "paragraph", "text": "x"}]}, 1')
    # the message on every Python version (3.13 has its own: 'Illegal trailing comma'); the repair deletes it
    assert "delete the comma before the ]" in unrepaired('[{"type": "paragraph", "text": "x"},]')
    assert "delete the comma before the }" in unrepaired('[{"type": "paragraph", "text": "x",}]')
    assert "cut off" in message('[{"type": "paragraph", "text": "x"}, {"type": "paragraph", "text": "ab')
    assert "a list stands where a key belongs" in message('[{"type": "table", ["a"]}]')
    # list items written as several arrays are not wrapped (they would become lists inside the list)
    assert "a list stands where a key belongs" in message('[{"type": "bullets", "items": ["a", "b"], ["c"]}]')
    assert "a key stands inside a list: the list before it is not closed" in message(
        '[{"type": "bullets", "items": ["a", "b", "level": 1}]'
    )
    assert "a comma is missing before the mark" in message('[{"type": "paragraph" "text": "x"}]')
    assert "a key is missing" in message('[{"type": "paragraph": "Der Arbeitsschutz prüft"}]')
    assert "a value is missing after the colon" in message('[{"type": "paragraph", "text": ]}]')
    assert "a backslash inside a text must be written as \\\\" in message('[{"type": "paragraph", "text": "C:\\x"}]')
    assert "a line break or tab inside a text" in message('[{"type": "paragraph", "text": "a\tb"}]')
    assert "the arguments are empty" in message(" ")
    assert "put all items into one list" in message('[{"type": "paragraph", "text": "x"}], [{"type": "page_break"}]')
    assert "they were cut off (output limit), or a text ends with" in message('[{"type": "paragraph", "text": "x\\"}]')
    assert "a text must start and end with a straight double quote" in message(
        '[{"type": "paragraph", "text": „Hallo“}]'
    )
    assert "write true, false and null in lower case" in message('[{"type": "heading", "text": "x", "level": None}]')
    assert "a list written with { }" in message('[{"type": "bullets", "items": {"a", "b"}}]')
    assert "delete the ``` after the JSON" in message('[{"type": "paragraph", "text": "x"}]\n```')


def test_garbled_arguments_are_rejected() -> None:
    slides = [{"layout": "bullets", "title": "T", "bullets": "a", '<|"|>b<|"|>],layout': "bullets"}]
    result = invalid("create_powerpoint_file", slides=slides)
    assert "slides[0]: the arguments are garbled" in result
    # Gemma's tool-call syntax broken inside a string: a key ends up in the text, the value is lost
    blocks = [{"type": "paragraph", "text": "a"}, {"type": 'paragraph",text:'}]
    assert "blocks[1]: the arguments are garbled" in invalid("create_word_file", blocks=blocks)
    sheets = [{"name": "S", "columns": [{"header": 'Preis (€)",type:'}], "rows": [[1]]}, {'"name': "T", "columns": []}]
    result = invalid("create_excel_file", sheets=sheets)
    assert "sheets[0]: the arguments are garbled" in result and "sheets[1]: the arguments are garbled" in result


# -------------------------------------------------------------------------------------------- Excel --

BUDGET = {
    "name": "Budget/2026",
    "columns": [
        {"header": "Item"},
        {"header": "Qty", "type": "integer"},
        {"header": "Price", "type": "currency"},
        {"header": "Total", "type": "currency", "formula": "=SUMME(B{row};C{row})*1", "total": "sum"},
        {"header": "Share", "type": "percent"},
        {"header": "Date", "type": "date"},
        {"header": "Ratio", "type": "number"},
    ],
    "rows": [
        ["Rent", 12, 1200, None, 45, "2026-01-01", 1.5],
        ["Travel", "3", "350,50 €", None, "7,5 %", "15.03.2026", 2.125],
        ['=HYPERLINK("http://x")', 1, "n/a", None, 12.5, "soon", 3],
    ],
    "chart": {"type": "column", "title": "Costs", "category_column": "Item", "value_columns": ["Price"]},
}


def test_excel_values_formats_formulas() -> None:
    workbook, filename = excel([BUDGET, {**BUDGET, "chart": None}], title="Budget_2026.xlsx")
    assert filename == "Budget 2026.xlsx"
    assert workbook.sheetnames == ["Budget 2026", "Budget 2026 (2)"]
    ws = workbook["Budget 2026"]
    assert [c.value for c in ws[1]] == ["Item", "Qty", "Price", "Total", "Share", "Date", "Ratio"]
    assert ws["B3"].value == 3 and ws["C3"].value == 350.5 and ws["C3"].number_format == od.CURRENCY_FORMATS["EUR"]
    assert ws["C4"].value == "n/a" and ws["C4"].data_type == "s"
    assert ws["D2"].value == "=SUM(B2,C2)*1" and ws["D4"].value == "=SUM(B4,C4)*1"
    assert ws["E2"].value == 0.45 and ws["E3"].value == 0.075 and ws["E2"].number_format == "0.0%"
    assert ws["F2"].value.date() == date(2026, 1, 1) and ws["F3"].value.date() == date(2026, 3, 15)
    assert ws["F2"].number_format == "YYYY-MM-DD" and ws["F4"].value == "soon"
    assert ws["G2"].number_format == "#,##0.000"
    assert ws["A4"].value == '=HYPERLINK("http://x")' and ws["A4"].data_type == "s"  # stays text
    assert ws["A5"].value == "Total" and ws["D5"].value == "=SUM(D2:D4)"
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref == "A1:G4"
    assert len(ws._charts) == 1 and not workbook["Budget 2026 (2)"]._charts
    chart = ws._charts[0]
    assert [s.val.numRef.f for s in chart.series] == ["'Budget 2026'!$C$2:$C$4"]
    assert chart.series[0].tx.v == "Price"


MONTHS_AS_COLUMNS = {
    "name": "Budget",
    "columns": [
        {"header": "Category"},
        *({"header": m, "type": "currency"} for m in ("Jan", "Feb", "Mar")),
        {"header": "Sum", "type": "currency", "formula": "=SUM(B{row}:D{row})"},
    ],
    "rows": [["Income", 3000, 3100, 2900], ["Expenses", 2200, 2400, 2100]],
}


def test_excel_chart_with_rows_as_series() -> None:
    chart = {"type": "line", "category_column": None, "series_rows": ["income", "Expenses"]}
    ws = excel([{**MONTHS_AS_COLUMNS, "chart": chart}])[0]["Budget"]
    (excel_chart,) = ws._charts
    assert [(s.tx.v, s.val.numRef.f) for s in excel_chart.series] == [
        ("Income", "'Budget'!$B$2:$D$2"),
        ("Expenses", "'Budget'!$B$3:$D$3"),
    ]  # the formula column "Sum" is not part of the x axis
    assert excel_chart.series[0].cat.numRef.f == "'Budget'!$B$1:$D$1"


def test_excel_chart_errors_guide_the_model() -> None:
    def errors(chart: dict[str, Any]) -> str:
        return invalid("create_excel_file", sheets=[{**MONTHS_AS_COLUMNS, "chart": chart}])

    assert "value_columns: is empty (or use series_rows" in errors({"category_column": "Category", "type": "line"})
    assert "if these are row labels, use series_rows" in errors(
        {"category_column": "Category", "value_columns": ["Income"]}
    )
    assert "series_rows: 'Profit' is not in the first column (Income, Expenses)" in errors({"series_rows": ["Profit"]})
    assert "must be next to each other" in errors({"series_rows": ["Income"], "value_columns": ["Jan", "Mar"]})


def test_excel_rows_without_formula_columns_and_german() -> None:
    sheet = {
        "name": "Offer",
        "columns": [
            {"header": "Pos"},
            {"header": "Total", "formula": "=C{row}*2"},
            {"header": "Qty", "type": "number"},
        ],
        "rows": [["A", 2], ["B", 3]],  # formula column left out
    }
    result, recorder = call(
        "create_excel_file", valves={"language": "de", "currency": "CHF"}, sheets=[sheet], title="Angebot"
    )
    ws = load_workbook(io.BytesIO(recorder.download()[0]))["Offer"]
    assert [[c.value for c in row] for row in ws.iter_rows(min_row=2)] == [["A", "=C2*2", 2], ["B", "=C3*2", 3]]
    assert recorder.statuses()[-1] == ("Angebot.xlsx erstellt", True)


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=B{row}*C{row}", "=B{row}*C{row}"),
        ('=wenn(b{row}>100;"hoch";"niedrig")', '=IF(B{row}>100,"hoch","niedrig")'),
        ("=ROUND(SUM($B$2:B{row}) / 3, 2)", "=ROUND(SUM($B$2:B{row})/3,2)"),
        ("=IFERROR(B2/C2, 0)", "=IFERROR(B{row}/C{row},0)"),  # cells of the first data row get {row}
        ("B{row}-C{row}", "=B{row}-C{row}"),  # "=" is added
        ("=AND(B2>=1, B2<>5, TRUE)", "=AND(B{row}>=1,B{row}<>5,TRUE)"),
        ("=B2/$C$1", "=B{row}/$C$1"),  # absolute rows stay
        ("=SUM(B2:B7)", "=SUM($B$2:$B$7)"),  # a range over several rows: the same in every row, written so
        ("=B{row}*0,19", "=B{row}*0.19"),  # decimal comma outside of a function
        ("=RUNDEN(B{row}*0,19;2)", "=ROUND(B{row}*0.19,2)"),  # with ";" every comma in a number is decimal
        ("=IF(A{row}>0,1,0)", "=IF(A{row}>0,1,0)"),  # commas inside a function separate arguments
    ],
)
def test_formula_whitelist_allows(formula: str, expected: str) -> None:
    assert od.normalize_formula(formula) == expected


@pytest.mark.parametrize(
    "formula",
    [
        '=WEBSERVICE("http://evil")',
        '=HYPERLINK("http://evil","x")',
        "=Sheet2!A1",
        "=[book.xlsx]Sheet1!A1",
        "=cmd|' /C calc'!A0",
        '=INDIRECT("A1")',
        "=SUM(B2:B5",
        "=SUM(B2))",
        "=B2@C2",
        "=foo",
        "=" + "1+" * 300 + "1",
        "=SUM(A1 B1)",  # a space would be Excel's intersection operator
        "=ZZZ1+1",  # beyond the last column XFD
        "=A1\uff13",  # full-width digit
    ],
)
def test_formula_whitelist_rejects(formula: str) -> None:
    with pytest.raises(ValueError):
        od.normalize_formula(formula)


SHEETS = {"stundensätze": "Stundensätze", "plan 2026": "Plan 2026", "it's": "It's"}


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=B2*Stundensätze!$B$3", "=B{row}*'Stundensätze'!$B$3"),  # quoted, written name; own cells get {row}
        ("=B{row}*'STUNDENSÄTZE'!B$2", "=B{row}*'Stundensätze'!B$2"),  # names without case, like Excel
        ("='Plan 2026'!B{row}-B{row}", "='Plan 2026'!B{row}-B{row}"),
        (
            "=SUMIF('Plan 2026'!$A$2:$A$9;A2;'Plan 2026'!$C$2:$C$9)",
            "=SUMIF('Plan 2026'!$A$2:$A$9,A{row},'Plan 2026'!$C$2:$C$9)",
        ),
        ("='It''s'!$A$1+1", "='It''s'!$A$1+1"),
        ("=SUM('Plan 2026'!B2:B9)", "=SUM('Plan 2026'!$B$2:$B$9)"),  # a range of several rows: fixed
    ],
)
def test_formula_cells_of_other_sheets(formula: str, expected: str) -> None:
    assert od.normalize_formula(formula, sheets=SHEETS) == expected


@pytest.mark.parametrize(
    ("formula", "message"),
    [
        ("=B2*'Stundensätze'!B3", "write 'Stundensätze'!$B$3 for this one cell in every row"),  # which one is meant?
        ("='Nope'!A1", "there is no sheet 'Nope' in this workbook (sheets: 'Stundensätze', 'Plan 2026', 'It's')"),
        ("='C:\\x\\[evil.xlsx]Sheet1'!A1", "references to other workbooks"),
        ("=Plan 2026!B{row}", "(cells of another sheet: 'Sheet name'!B3)"),
        ("=B{row}:'Plan 2026'!B3", "a range as 'Plan 2026'!B2:B9"),
        ("='Plan 2026'!SUM(A1)", "a range as 'Plan 2026'!B2:B9"),
        ("=B{row} 'Plan 2026'!$B$3", "a space between two values"),
    ],
)
def test_formula_cells_of_other_sheets_rejected(formula: str, message: str) -> None:
    with pytest.raises(ValueError, match=re.escape(message)):
        od.normalize_formula(formula, sheets=SHEETS)


def test_excel_cells_of_other_sheets() -> None:
    rates = {"name": "Stundensätze", "columns": [{"header": "Rolle"}, {"header": "Satz", "type": "currency"}],
             "rows": [["Senior", 120], ["Junior", 80]]}  # fmt: skip
    offer = {
        "name": "Angebot",
        "columns": [{"header": "Position"}, {"header": "Stunden", "type": "number"},
                    {"header": "Betrag", "type": "currency", "formula": "=B2*Stundensätze!$B$2", "total": "sum"}],
        "rows": [["Analyse", 10], ["Umsetzung", 20]],
    }  # fmt: skip
    summary = {"name": "Übersicht", "columns": [{"header": "Posten"}, {"header": "Wert", "type": "currency"}],
               "rows": [["Angebot", "=Angebot!C4"], ["Junior", "='stundensätze'!B3"]]}  # fmt: skip
    wb, _ = excel([offer, rates, summary])
    assert [wb["Angebot"][f"C{r}"].value for r in (2, 3, 4)] == [
        "=B2*'Stundensätze'!$B$2", "=B3*'Stundensätze'!$B$2", "=SUM(C2:C3)"
    ]  # fmt: skip
    assert (wb["Übersicht"]["B2"].value, wb["Übersicht"]["B3"].value) == ("='Angebot'!C4", "='Stundensätze'!B3")

    def with_formula(formula: str, rows: int = 2) -> dict[str, Any]:
        amount = {**offer["columns"][2], "formula": formula}
        return {**offer, "columns": [*offer["columns"][:2], amount], "rows": [["x", 1]] * rows}

    # A row counted wrong: an empty cell would silently count as 0
    result = invalid("create_excel_file", sheets=[with_formula("=B2*'Stundensätze'!$B$4"), rates])
    assert "sheets[0].columns[2].formula: 'Stundensätze'!B4 is an empty cell" in result
    assert "row 1 is the header and the data is in rows 2 to 3, columns A to B" in result
    result = invalid("create_excel_file", sheets=[offer, rates, {**summary, "rows": [["Angebot", "=Angebot!C5"]]}])
    assert "sheets[2].rows[0]: 'Angebot'!C5 is an empty cell" in result
    result = invalid("create_excel_file", sheets=[with_formula("=B{row}*'Stundensätze'!B{row}", rows=3), rates])
    assert "'Stundensätze'!B{row} is empty in the last rows: 'Stundensätze' has 2 data rows, this sheet 3" in result
    result = invalid("create_excel_file", sheets=[{**summary, "rows": [["Angebot", "='Kosten'!B2"]]}])
    assert "there is no sheet 'Kosten' in this workbook (sheets: 'Übersicht')" in result
    result = invalid("create_excel_file", sheets=[with_formula("=B{row}*'Stundensätze'!Z{row}"), rates])
    assert "'Stundensätze'!Z{row} is empty: 'Stundensätze' has no values in column Z" in result
    result = invalid("create_excel_file", sheets=[with_formula("=Angebot!C{row}*2"), rates])  # its own sheet by name
    assert "the formula uses its own column C" in result


def test_excel_sheet_names_in_formulas() -> None:
    sheets = [
        {"name": "Maße", "columns": [{"header": "A"}, {"header": "B"}], "rows": [["x", 80]]},
        {"name": "Masse", "columns": [{"header": "A"}, {"header": "B"}], "rows": [["x", 90]]},  # not the same in Excel
        {"name": "a;b", "columns": [{"header": "A"}, {"header": "B"}], "rows": [["x", 10]]},
        {"name": "S", "columns": [{"header": "A"}, {"header": "B"}],
         "rows": [["Maße", "='Maße'!B2"], ["Sum", "=SUM('a;b'!B2,1,5)"]]},
    ]  # fmt: skip
    ws = excel(sheets)[0]["S"]
    assert (ws["B2"].value, ws["B3"].value) == ("='Maße'!B2", "=SUM('a;b'!B2,1,5)")  # ';' in a name: not German


def test_excel_invalid_input() -> None:
    sheet = {
        "name": "S",
        "columns": [{"header": "A"}, {"header": "B", "formula": '=WEBSERVICE("x")'}],
        "rows": [["a", 1, 2]],
        "chart": {"category_column": "Nope", "value_columns": ["A", "Missing"]},
    }
    result = invalid("create_excel_file", sheets=[sheet])
    assert "sheets[0].columns[1].formula: the function WEBSERVICE is not allowed" in result
    assert "sheets[0].rows[0]: 3 values but only 2 columns" in result
    assert "'Nope' is not a column header (A, B)" in result and "'Missing' is not a column header (A, B)" in result
    assert "the workbook has no sheets" in invalid("create_excel_file", sheets=[])
    too_long = {"name": "S", "columns": [{"header": "A"}], "rows": [["x"]] * (od.MAX_SHEET_ROWS + 1)}
    assert "at most" in invalid("create_excel_file", sheets=[too_long])


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1.234,56 €", (1234.56, False)),
        ("1,234.5", (1234.5, False)),
        ("-3,2 %", (-0.032, True)),
        ("12", (12, False)),
        ("abc", None),
        ("€ 5 %", None),
    ],
)
def test_parse_number_text(text: str, expected: Any) -> None:
    result = od.parse_number_text(text)
    assert result == expected or (result and expected and result[0] == pytest.approx(expected[0]))


# ---------------------------------------------------------------------------------------- PowerPoint --

SLIDES = [
    {"layout": "title", "title": "Quarterly report", "subtitle": "Sales"},
    {"layout": "section", "title": "Part 1"},
    {"layout": "bullets", "title": "Results", "bullets": ["Revenue **+12 %**", "- online", "[Wiki](https://example.org)", "[x](file:///etc/passwd)"], "notes": "Say something."},
    {"layout": "two_columns", "title": "Pros and cons", "left_heading": "Pros", "left": ["fast"], "right": ["expensive"]},
    {"layout": "table", "title": "Regions", "table": [["Region", "Q3"], ["North", "1.4"], ["South"]]},
    {"layout": "chart", "title": "Trend", "chart": {"type": "line", "categories": ["Q1", "Q2"], "series": [{"name": "A", "values": [1, 2]}, {"name": "B", "values": [2, 3]}]}},
    {"layout": "chart", "title": "Shares", "chart": {"type": "pie", "categories": ["N", "S"], "series": [{"name": "Sales", "values": [60, 40]}, {"name": "ignored", "values": [1, 1]}]}},
]  # fmt: skip


def test_powerpoint_structure() -> None:
    prs, filename = powerpoint(SLIDES)
    assert filename == "Quarterly report.pptx"  # title taken from the first slide
    assert (prs.slide_width, prs.slide_height) == (od.SLIDE_WIDTH, od.SLIDE_HEIGHT)
    slides = list(prs.slides)
    assert len(slides) == 7
    assert slides[0].shapes.title.text == "Quarterly report"
    assert all(ph.text_frame.text.strip() for slide in slides for ph in slide.placeholders)  # no empty placeholders
    body = slides[2].placeholders[1].text_frame
    assert [(p.level, "".join(r.text for r in p.runs)) for p in body.paragraphs] == [
        (0, "Revenue +12 %"), (1, "online"), (0, "Wiki"), (0, "x")
    ]  # fmt: skip
    links = [r.hyperlink.address for p in body.paragraphs for r in p.runs if r.hyperlink.address]
    assert links == ["https://example.org"]
    assert slides[2].notes_slide.notes_text_frame.text == "Say something."
    columns = [ph.text_frame.text for ph in slides[3].placeholders if ph.placeholder_format.idx in (1, 2)]
    assert columns == ["Pros\nfast", "expensive"]
    table = next(s for s in slides[4].shapes if s.has_table).table
    assert [[table.cell(r, c).text for c in range(2)] for r in range(3)] == [
        ["Region", "Q3"],
        ["North", "1.4"],
        ["South", ""],
    ]
    line_chart = next(s for s in slides[5].shapes if s.has_chart).chart
    assert [s.name for s in line_chart.plots[0].series] == ["A", "B"] and line_chart.has_legend
    pie = next(s for s in slides[6].shapes if s.has_chart).chart
    assert [s.name for s in pie.plots[0].series] == ["Sales"]



def test_powerpoint_chart_xml_opens_in_powerpoint() -> None:
    # a <c:numFmt> without formatCode makes PowerPoint repair the file and drop the chart
    chart = {"categories": ["Q1", "Q2"], "series": [{"name": "s", "values": [1, 2]}]}
    slides = [{"layout": "chart", "title": t, "chart": {**chart, "type": t}} for t in ("column", "bar", "line", "pie")]
    data, _ = od.build_pptx("D", od.parse_argument(od.SLIDES, slides, "slides"))
    with zipfile.ZipFile(io.BytesIO(data)) as deck:
        charts = [deck.read(n).decode() for n in deck.namelist() if n.startswith("ppt/charts/chart")]
    assert len(charts) == 4 and all("<c:dLbls>" in xml for xml in charts)
    assert not [f for xml in charts for f in re.findall(r"<c:numFmt [^>]*>", xml) if "formatCode=" not in f]

def test_powerpoint_invalid_input() -> None:
    slides = [
        {"layout": "bullets", "title": "", "bullets": []},
        {"layout": "bullets", "title": "Long", "bullets": ["x"] * 13},
        {
            "layout": "chart",
            "title": "C",
            "chart": {"categories": ["a", "b"], "series": [{"name": "s", "values": [1]}]},
        },
        {"layout": "table", "title": "T", "table": [["x"]] * 17},
        {"layout": "chart", "title": "No chart"},
    ]
    result = invalid("create_powerpoint_file", slides=slides)
    assert "slides[0] (bullets): 'title' is empty" in result and "slides[0] (bullets): 'bullets' is empty" in result
    assert "at most 12 items" in result
    assert "1 values for 2 categories" in result
    assert "at most 16 rows" in result
    assert "slides[4] (chart): 'chart' is missing" in result


def test_design_follows_the_accent_color() -> None:
    """One accent colour drives all three formats; the shades are derived from it."""
    accent, valves = "8B1E3F", {"accent_color": "#8b1e3f"}
    design = od.Design(accent)
    blocks = [{"type": "heading", "text": "Teil"}, {"type": "table", "rows": [["A", "B"], ["1", "2"], ["3", "4"]]}]
    result, recorder = call("create_word_file", valves=valves, blocks=blocks, title="Bericht")
    doc = Document(io.BytesIO(recorder.download()[0]))
    assert doc.styles["Heading 1"].font.color.rgb == od.RGBColor.from_string(accent)
    fills = [cell._tc.tcPr.find(od.qn("w:shd")) for cell in doc.tables[0].column_cells(0)]
    assert [f.get(od.qn("w:fill")) if f is not None else None for f in fills] == [accent, None, design.band]
    footer = doc.sections[0].footer.paragraphs[0]
    assert footer.text.startswith("Bericht\t") and footer.paragraph_format.tab_stops[0].alignment == 2  # right
    sheet = {"name": "S", "columns": [{"header": "Monat"}, {"header": "Wert", "type": "number", "total": "sum"}],
             "rows": [["Jan", 1], ["Feb", 2]], "chart": {"type": "column", "category_column": "Monat", "value_columns": ["Wert"]}}  # fmt: skip
    result, recorder = call("create_excel_file", valves=valves, sheets=[sheet])
    ws = load_workbook(io.BytesIO(recorder.download()[0]))["S"]
    assert ws["A1"].fill.fgColor.rgb.endswith(accent) and ws["A1"].font.color.rgb.endswith("FFFFFF")
    (banding,) = ws.conditional_formatting  # a rule, so the stripes stay in place when the user sorts
    assert str(banding.sqref) == "A2:B3" and banding.rules[0].formula == ["MOD(ROW(),2)=1"]
    assert banding.rules[0].dxf.fill.bgColor.rgb.endswith(design.band) and ws["A3"].fill.fill_type is None
    assert ws["A4"].value == "Total"
    assert not ws.sheet_view.showGridLines and ws.page_setup.fitToWidth == 1 and ws.print_title_rows == "$1:$1"
    workbook = zipfile.ZipFile(io.BytesIO(recorder.download()[0]))
    chart_xml = workbook.read(next(n for n in workbook.namelist() if n.startswith("xl/charts/chart"))).decode()
    assert f'srgbClr val="{accent}"' in chart_xml and chart_xml.count('<delete val="0"/>') == 2  # axes shown in Excel
    prs, _ = powerpoint([{"layout": "title", "title": "Deck"}, {"layout": "bullets", "title": "B", "bullets": ["a"]}],
                        valves=valves)  # fmt: skip
    title_slide, content = prs.slides
    assert title_slide.background.fill.fore_color.rgb == od.PptxRGBColor.from_string(design.dark)
    assert content.shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb == od.PptxRGBColor.from_string(accent)
    assert "2" in [shape.text_frame.text for shape in content.shapes if shape.has_text_frame]  # slide number


def test_house_font_in_all_formats() -> None:
    valves = {"font": "  Source   Sans 3 "}
    _, recorder = call("create_word_file", valves=valves, blocks=[{"type": "paragraph", "text": "x"}], title="T")
    doc = Document(io.BytesIO(recorder.download()[0]))
    assert {doc.styles[name].font.name for name in ("Normal", "Title", "Heading 1", "Footer")} == {"Source Sans 3"}
    sheet = {"name": "S", "columns": [{"header": "A"}], "rows": [["x"]]}
    _, recorder = call("create_excel_file", valves=valves, sheets=[sheet])
    data = recorder.download()[0]
    assert load_workbook(io.BytesIO(data))["S"]["A1"].font.name == "Source Sans 3"
    assert zipfile.ZipFile(io.BytesIO(data)).read("xl/theme/theme1.xml").decode().count('typeface="Source Sans 3"') == 2
    _, recorder = call("create_powerpoint_file", valves=valves, slides=[{"layout": "title", "title": "Deck"}])
    theme = zipfile.ZipFile(io.BytesIO(recorder.download()[0])).read("ppt/theme/theme1.xml").decode()
    assert theme.count('<a:latin typeface="Source Sans 3"/>') == 2  # headings and body text, also in charts
    _, recorder = call("create_word_file", valves={"font": 'Arial"/><w:x'}, blocks=[{"type": "paragraph", "text": "x"}])
    assert Document(io.BytesIO(recorder.download()[0])).styles["Normal"].font.name == "Calibri"  # not a font name


def test_markdown_text_of_the_files() -> None:
    """The text Open WebUI shows in the "Content" tab of its file view."""
    options = od.WordOptions(title="Bericht", subtitle="Q3")
    blocks = od.parse_argument(od.WORD_BLOCKS, [
        {"type": "heading", "text": "Teil", "level": 2}, {"type": "paragraph", "text": "Text **fett**"},
        {"type": "numbered", "items": ["a", "- a1"]}, {"type": "table", "rows": [["A", "B"], ["1|2", "x"]]},
        {"type": "quote", "text": "Zitat"}, {"type": "paragraph", "text": " "},
    ], "blocks")  # fmt: skip
    data, text = od.build_docx(options, blocks, od.TRANSLATIONS["de"])
    assert text == ("# Bericht\n\n*Q3*\n\n### Teil\n\nText **fett**\n\n1. a\n   1. a1\n\n"
                    "| A | B |\n| --- | --- |\n| 1\\|2 | x |\n\n> Zitat")  # fmt: skip
    sheet = {"name": "Budget", "columns": [{"header": "Posten"}, {"header": "Betrag", "type": "currency", "total": "sum"},
             {"header": "Anteil", "type": "percent"}, {"header": "Datum", "type": "date"}],
             "rows": [["Miete", 1200.5, 45, "2026-01-01"]], "chart": {"type": "pie", "title": "Kosten",
             "category_column": "Posten", "value_columns": ["Betrag"]}}  # fmt: skip
    sheets = od.parse_argument(od.EXCEL_SHEETS, [sheet], "sheets")
    data, text = od.build_xlsx(sheets, "EUR", od.TRANSLATIONS["de"], decimal_comma=True)
    assert "| Miete | 1.200,50 € | 45 % | 01.01.2026 |" in text and "| Summe | =SUM(B2:B2) |" in text
    assert text.startswith("## Budget") and text.endswith("*Diagramm: Kosten (pie)*")
    slides = od.parse_argument(od.SLIDES, [
        {"layout": "title", "title": "Deck", "subtitle": "Sub"},
        {"layout": "chart", "title": "Umsatz", "notes": "Kurz\nhalten", "chart": {"categories": ["Q1", "Q2"],
         "series": [{"name": "2026", "values": [1.5, 2]}], "unit": "T€"}},
    ], "slides")  # fmt: skip
    data, text = od.build_pptx("Deck", slides, texts=od.TRANSLATIONS["de"])
    assert text == ("# Deck\n\n## 1. Deck\n\nSub\n\n## 2. Umsatz\n\n*Diagramm (column, T€)*\n\n"
                    "|  | Q1 | Q2 |\n| --- | --- | --- |\n| 2026 | 1.5 | 2 |\n\n> Notizen: Kurz halten")  # fmt: skip


def test_word_files_render_in_browser_previews() -> None:
    """docx-preview (Open WebUI's preview) reads the document defaults, the theme, the linked character styles and
    the bullets' own font, which Word itself does not use."""
    _, recorder = call(
        "create_word_file", valves={"accent_color": "8B1E3F"}, blocks=[{"type": "bullets", "items": ["a"]}]
    )
    package = zipfile.ZipFile(io.BytesIO(recorder.download()[0]))
    styles, numbering = package.read("word/styles.xml").decode(), package.read("word/numbering.xml").decode()
    defaults = re.search(r"<w:docDefaults>.*?</w:docDefaults>", styles, re.S)[0]
    assert 'w:ascii="Calibri"' in defaults and "asciiTheme" not in defaults  # not the theme's Cambria
    heading_char = re.search(r'w:styleId="Heading1Char".*?</w:style>', styles, re.S)[0]
    assert '<w:color w:val="8B1E3F"/>' in heading_char and "themeColor" not in heading_char
    assert "Symbol" not in numbering and 'w:lvlText w:val="\u2022"' in numbering  # a real bullet, not a Symbol glyph
    assert 'typeface="Cambria"' not in package.read("word/theme/theme1.xml").decode()


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
)


def rewrite_package(data: bytes, drop: str = "", types: tuple[str, str] = ("", "")) -> bytes:
    """Remove a part (and its relationship and content type) or change a content type of a Word package."""
    source, buffer = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as target:
        for item in source.infolist():
            content = source.read(item)
            if drop and item.filename == drop:
                continue
            if item.filename in ("[Content_Types].xml", "word/_rels/document.xml.rels") and drop:
                content = re.sub(rf"<(Override|Relationship)[^>]*{Path(drop).name}\"[^>]*/>".encode(), b"", content)
            if item.filename == "[Content_Types].xml" and types[0]:
                content = content.replace(types[0].encode(), types[1].encode())
            target.writestr(item, content)
    return buffer.getvalue()


def letterhead(path: Path, header: bool = True, numbering: bool = True) -> str:
    """A company template: its own Normal font, a header with a logo, a footer, wide margins, sample text, a mail
    merge and an author; saved as .dotx."""
    template = Document()
    section = template.sections[0]
    section.top_margin = od.Cm(4.5)
    template.styles["Normal"].font.name = "Georgia"
    if header:
        section.header.paragraphs[0].text = "ACME GmbH "
        section.header.paragraphs[0].add_run().add_picture(io.BytesIO(PNG_1X1))
        section.footer.paragraphs[0].text = "Industriestr. 1, Stuttgart"
    template.add_paragraph("MUSTERTEXT", style="List Bullet")
    merge = od.OxmlElement("w:mailMerge")
    merge.append(od.OxmlElement("w:mainDocumentType"))
    template.settings.element.append(merge)
    template.core_properties.author = "Max Mustermann"
    buffer = io.BytesIO()
    template.save(buffer)
    data = rewrite_package(buffer.getvalue(), types=("document.main+xml", "template.main+xml"))
    if not numbering:
        data = rewrite_package(data, drop="word/numbering.xml")
    path.write_bytes(data)
    return str(path)


def test_word_template(tmp_path: Path) -> None:
    blocks = [{"type": "numbered", "items": ["a", "b"]}, {"type": "table", "rows": [["A", "B"], ["1", "2"]]},
              {"type": "bullets", "items": ["x", "- y"]}, {"type": "paragraph", "text": "Text"}]  # fmt: skip
    _, recorder = call(
        "create_word_file", valves={"word_template": letterhead(tmp_path / "letterhead.dotx")},
        blocks=blocks, title="Angebot", header_text="Vertraulich",
    )  # fmt: skip
    data = recorder.download()[0]
    doc = Document(io.BytesIO(data))
    section = doc.sections[0]
    assert round(section.top_margin.cm, 2) == 4.5  # the template's page
    assert section.header.paragraphs[0].text == "ACME GmbH "  # not added to: no "Vertraulich", no page numbers
    assert section.footer.paragraphs[0].text == "Industriestr. 1, Stuttgart"
    assert any("image" in rel.reltype for rel in section.header.part.rels.values())  # the logo
    texts = [p.text for p in doc.paragraphs]
    assert "MUSTERTEXT" not in texts and texts[0] == "Angebot" and "Text" in texts
    assert doc.styles["Normal"].font.name == "Georgia"  # the header and footer build on it
    assert {p.style.name for p in doc.paragraphs if p.text in ("Text", "")} == {"Body Text"}
    assert doc.styles["Body Text"].font.name == "Calibri" and doc.styles["Body Text"].base_style is None
    assert doc.tables[0].rows[1].cells[0].paragraphs[0].style.name == "Body Text"
    numbering = doc.part.numbering_part.element
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith("List"):  # every list points to an existing numbering
            num_id = paragraph._p.pPr.numPr.numId.val if paragraph._p.pPr.numPr is not None else None
            num_id = num_id or paragraph.style.element.pPr.numPr.numId.val
            assert numbering.num_having_numId(num_id) is not None
    package = zipfile.ZipFile(io.BytesIO(data))
    assert "mailMerge" not in package.read("word/settings.xml").decode()
    assert doc.core_properties.author == "" and doc.core_properties.title == "Angebot"


def test_word_template_without_header_or_numbering(tmp_path: Path) -> None:
    path = letterhead(tmp_path / "plain.dotx", header=False, numbering=False)
    doc, _, _ = word(valves={"word_template": path}, blocks=[{"type": "numbered", "items": ["a"]}], header_text="H")
    assert round(doc.sections[0].top_margin.cm, 2) == 4.5
    assert doc.sections[0].header.paragraphs[0].text == "H"  # the template has none: added as usual
    assert doc.sections[0].footer.paragraphs[0].text.endswith("Page 1 of 1")  # and the page numbers
    style_num_id = doc.styles["List Number"].element.pPr.numPr.numId.val
    assert doc.part.numbering_part.element.num_having_numId(style_num_id) is not None


def test_word_template_shapes(tmp_path: Path) -> None:
    sections = Document()  # the letterhead is in the first section; the last one takes it over (Word's default)
    sections.sections[0].header.paragraphs[0].text = "LETTERHEAD"
    sections.add_section()
    sections.save(tmp_path / "sections.docx")
    doc, _, _ = word(valves={"word_template": str(tmp_path / "sections.docx")}, blocks=[{"type": "paragraph", "text": "x"}],
                     header_text="H")  # fmt: skip
    assert [p.text for p in doc.sections[0].header.paragraphs] == ["LETTERHEAD"] and len(doc.sections) == 1

    plain = Path(letterhead(tmp_path / "plain.dotx", header=False))  # has a Footer style with tab stops, no footer
    doc, _, _ = word(valves={"word_template": str(plain)}, blocks=[{"type": "paragraph", "text": "x"}])
    tabs = doc.sections[0].footer.paragraphs[0].paragraph_format.tab_stops
    assert [tab.alignment for tab in tabs].count(od.WD_TAB_ALIGNMENT.RIGHT) == 1  # the page number on the right
    assert od.WD_TAB_ALIGNMENT.CLEAR in [tab.alignment for tab in tabs]  # not at the style's centre tab

    no_settings = tmp_path / "no-settings.dotx"  # settings.xml is optional
    no_settings.write_bytes(rewrite_package(plain.read_bytes(), drop="word/settings.xml"))
    doc, _, _ = word(valves={"word_template": str(no_settings)}, blocks=[{"type": "paragraph", "text": "x"}])
    assert round(doc.sections[0].top_margin.cm, 2) == 4.5

    no_page = Document()  # no page size and margins: the usual ones
    for element in no_page.sections[0]._sectPr.findall(od.qn("w:pgMar")) + no_page.sections[0]._sectPr.findall(
        od.qn("w:pgSz")
    ):
        no_page.sections[0]._sectPr.remove(element)
    no_page.sections[0].header.paragraphs[0].text = "ACME"
    no_page.save(tmp_path / "no-page.docx")
    doc, _, _ = word(
        valves={"word_template": str(tmp_path / "no-page.docx")}, blocks=[{"type": "paragraph", "text": "x"}]
    )
    assert doc.sections[0].header.paragraphs[0].text == "ACME" and round(doc.sections[0].left_margin.cm, 2) == 2.5


@pytest.mark.parametrize("content", [None, b"not a zip", b"PK\x03\x04broken"])
def test_word_template_that_cannot_be_used(tmp_path: Path, caplog: pytest.LogCaptureFixture, content: bytes) -> None:
    path = tmp_path / "template.docx"
    if content is not None:
        path.write_bytes(content)
    with caplog.at_level("WARNING"):
        doc, _, _ = word(valves={"word_template": str(path)}, blocks=[{"type": "paragraph", "text": "x"}])
    assert round(doc.sections[0].top_margin.cm, 2) == 2.5  # made without it
    assert "documents are made without it" in caplog.text


def test_bad_accent_color_falls_back() -> None:
    prs, _ = powerpoint(SLIDES[:3], valves={"accent_color": "red; drop"})
    assert prs.slides[2].shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb == od.PptxRGBColor.from_string(
        "1F4E79"
    )


def test_body_font_size() -> None:
    assert od.body_font_size(["a"] * 7) == 24  # short items keep the full size
    assert od.body_font_size(["a"] * 12) == 18
    assert od.body_font_size(["x" * 200] * 6) == 16  # long items wrap
    assert od.body_font_size(["x" * 200] * 12) is None  # does not fit at all: the model has to split the slide
    narrow = od.body_font_size(["x" * 60] * 6, base=22, width=(od.CONTENT_WIDTH - od.COLUMN_GAP) // 2)
    assert narrow < od.body_font_size(["x" * 60] * 6, base=22)


# ----------------------------------------------------------------------------------------- delivery --


class FakeOpenWebUI:
    """Stands in for the Open WebUI modules attach_file imports."""

    def __init__(self, monkeypatch: pytest.MonkeyPatch, fail_upload: bool = False) -> None:
        self.uploads: list[dict[str, Any]] = []
        self.linked: list[dict[str, Any]] = []
        self.message_files: list[dict[str, Any]] = [{"type": "file", "id": "earlier"}]
        self.file_data: dict[str, dict[str, Any]] = {}
        fake = self

        async def upload_file_handler(request: Any, file: Any, metadata: Any, process: bool, user: Any) -> Any:
            if fail_upload:
                raise RuntimeError("storage down")
            fake.uploads.append(
                {"name": file.filename, "type": file.content_type, "data": file.file.read(), "process": process,
                 "metadata": metadata, "user": user.id}
            )  # fmt: skip
            return SimpleNamespace(id="file-1")

        class Chats:
            @staticmethod
            async def insert_chat_files(**kwargs: Any) -> None:
                fake.linked.append(kwargs)

            @staticmethod
            async def add_message_files_by_id_and_message_id(chat_id: str, message_id: str, files: list) -> list:
                fake.message_files += files
                return fake.message_files

        class UserModel:
            def __init__(self, **data: Any) -> None:
                self.id = data["id"]

        class Files:
            @staticmethod
            async def update_file_data_by_id(id: str, data: dict) -> None:
                fake.file_data.setdefault(id, {}).update(data)

        modules = {
            "open_webui": ModuleType("open_webui"),
            "open_webui.models": ModuleType("open_webui.models"),
            "open_webui.models.chats": SimpleNamespace(Chats=Chats),
            "open_webui.models.users": SimpleNamespace(UserModel=UserModel),
            "open_webui.models.files": SimpleNamespace(Files=Files),
            "open_webui.routers": ModuleType("open_webui.routers"),
            "open_webui.routers.files": SimpleNamespace(upload_file_handler=upload_file_handler),
            "open_webui.utils": ModuleType("open_webui.utils"),
            "open_webui.utils.chat_id": SimpleNamespace(
                is_saved_chat_id=lambda c: bool(c) and not c.startswith("local:")
            ),
        }
        for name, module in modules.items():
            monkeypatch.setitem(sys.modules, name, module)


def attach(chat_id: str, with_request: bool = True) -> tuple[str, Recorder]:
    pytest.importorskip("starlette")  # ships with Open WebUI (FastAPI)
    tools = od.Tools()
    recorder = Recorder()
    request = SimpleNamespace(app=SimpleNamespace(url_path_for=lambda name, id: f"/api/v1/files/{id}/content"))
    result = asyncio.run(
        tools.create_excel_file(
            sheets=[{"name": "S", "columns": [{"header": "A"}], "rows": [["x"]]}],
            title="Book",
            __user__={"id": "u1", "name": "Ada"},
            __request__=request if with_request else None,
            __chat_id__=chat_id,
            __message_id__="m1",
            __event_emitter__=recorder.emit,
            __event_call__=recorder.call,
        )
    )
    return result, recorder


def test_attach_to_saved_chat(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOpenWebUI(monkeypatch)
    result, recorder = attach("chat-1")
    assert result.startswith("Done") and "attached to your answer" in result
    (upload,) = fake.uploads
    assert upload["name"] == "Book.xlsx" and upload["type"] == od.XLSX_MIME and not upload["process"]
    assert upload["user"] == "u1" and upload["metadata"] == {"chat_id": "chat-1", "message_id": "m1"}
    load_workbook(io.BytesIO(upload["data"]))
    assert fake.linked == [{"chat_id": "chat-1", "message_id": "m1", "file_ids": ["file-1"], "user_id": "u1"}]
    assert fake.file_data["file-1"]["content"] == "## S\n\n| A |\n| --- |\n| x |"  # the "Content" tab
    (files_event,) = [e for e in recorder.events if e["type"] == "chat:message:files"]
    assert [f["id"] for f in files_event["data"]["files"]] == ["earlier", "file-1"]  # earlier files stay visible
    entry = files_event["data"]["files"][1]
    assert entry == {
        "type": "file", "id": "file-1", "url": "/api/v1/files/file-1/content", "name": "Book.xlsx",
        "size": len(upload["data"]), "content_type": od.XLSX_MIME,
    }  # fmt: skip
    assert not recorder.calls


def test_attach_in_temporary_chat(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOpenWebUI(monkeypatch)
    result, recorder = attach("local:abc")
    assert result.startswith("Done") and not fake.linked
    (files_event,) = [e for e in recorder.events if e["type"] == "chat:message:files"]
    assert [f["id"] for f in files_event["data"]["files"]] == ["file-1"]


@pytest.mark.parametrize("fail_upload", [True, False])
def test_attach_falls_back_to_download(monkeypatch: pytest.MonkeyPatch, fail_upload: bool) -> None:
    FakeOpenWebUI(monkeypatch, fail_upload=fail_upload)
    result, recorder = attach("chat-1", with_request=fail_upload)  # without a request, attaching fails early
    assert "downloaded in the user's browser" in result
    assert recorder.download()[1] == "Book.xlsx"


def test_no_way_to_deliver() -> None:
    tools = make_tools()
    result = asyncio.run(
        tools.create_word_file(blocks=[{"type": "paragraph", "text": "x"}], title="T", __event_call__=None)
    )
    assert result == od.RESULT_FAILED


def test_internal_error_is_not_leaked(monkeypatch: pytest.MonkeyPatch) -> None:
    def boom(*args: Any) -> bytes:
        raise RuntimeError("secret path /root/x")

    monkeypatch.setattr(od, "build_docx", boom)
    result, recorder = call("create_word_file", blocks=[{"type": "paragraph", "text": "x"}])
    assert result == od.RESULT_FAILED and "secret" not in result
    assert recorder.statuses()[-1] == ("The file could not be created", True)


# ------------------------------------------------------------------------------------ review fixes --


def test_repairs_never_drop_content() -> None:
    """A repair that would lose text is not applied; the model gets the position of the problem instead."""
    stray = '[{"type": "heading", "text": "Intro"}}, {"type": "paragraph", "text": "Body"}]'  # only a brace too many
    assert od.loads_lenient(stray) == [{"type": "heading", "text": "Intro"}, {"type": "paragraph", "text": "Body"}]
    cut_off = '[{"type": "paragraph", "text": "A"}, {"type": "table", "rows": [["H1", "H2"], ["c"'
    assert "the JSON is invalid" in invalid("create_word_file", blocks=cut_off)
    split_list = '[{"type": "paragraph", "text": "A"}], [{"type": "paragraph", "text": "B"}]'
    assert "the JSON is invalid" in invalid("create_word_file", blocks=split_list)
    lost_key = (
        '[{"layout": "chart", "title": "C", {"categories": ["a"], "series": []}}, {"layout": "section", "title": "S"}]'
    )
    assert "the JSON is invalid" in invalid("create_powerpoint_file", slides=lost_key)
    # a correctly escaped quote after „ does not stop a bracket repair
    escaped = '[{"type": "paragraph", "text": "Er sagte „\\"Hallo\\" und ging."}'
    assert od.loads_lenient(escaped) == [{"type": "paragraph", "text": 'Er sagte „"Hallo" und ging.'}]


def test_bracket_and_quote_slips_are_repaired() -> None:
    """Slips seen in real runs; only brackets and quotes change, and a repair that would merge two objects is not
    taken."""
    swapped = '[{"name": "A", "rows": [["a", 1], ["b", 2]}], {"name": "B", "rows": [["c", 3]]}]'  # '}]' for ']}'
    assert od.loads_lenient(swapped) == [{"name": "A", "rows": [["a", 1], ["b", 2]]}, {"name": "B", "rows": [["c", 3]]}]
    extra = '[{"type": "bullets", "items": ["a", "b"]]}, {"type": "paragraph", "text": "x"}]'
    assert od.loads_lenient(extra) == [{"type": "bullets", "items": ["a", "b"]}, {"type": "paragraph", "text": "x"}]
    no_brace = '[{"name": "A", "rows": [["a", 1]}], "name": "B", "rows": [["c", 3]]}]'  # the '{' of B is missing
    assert od.loads_lenient(no_brace) == [{"name": "A", "rows": [["a", 1]]}, {"name": "B", "rows": [["c", 3]]}]
    stray_quote = '[{"type": "heading", "text": "Ziel", "level": 1"}, {"type": "paragraph", "text": "x"}]'
    assert od.loads_lenient(stray_quote)[0] == {"type": "heading", "text": "Ziel", "level": 1}
    open_quote = '[{"type": "heading", "text": "Ziel", level": 2}]'
    assert od.loads_lenient(open_quote) == [{"type": "heading", "text": "Ziel", "level": 2}]
    trailing = '[{"type": "table", "rows": [["a", "b", ], ["c", "d"]], }, {"type": "bullets", "items": ["x",]},]'
    assert od.loads_lenient(trailing) == [{"type": "table", "rows": [["a", "b"], ["c", "d"]]},
                                          {"type": "bullets", "items": ["x"]}]  # fmt: skip
    assert od.loads_lenient('[{"type": "table", "type": "table", "rows": [["a"]}]')[0]["rows"] == [["a"]]
    assert od._repair_json('[{"type": "paragraph", "text": "A", "text": "B"]') is None  # "A" would be lost
    split = '[{"name": "S", "chart": {"type": "pie", "value_columns": ["E"}, "unit": "EUR"}}]'
    assert od.loads_lenient(split)[0]["chart"] == {"type": "pie", "value_columns": ["E"], "unit": "EUR"}  # one chart
    after = '[{"type": "table", "rows": [["a"]]], "title": "Bericht", "text": "Wichtig"}, {"type": "paragraph"}]'
    assert "the JSON is invalid" in invalid("create_word_file", blocks=after)  # "Wichtig" would be lost
    sheets = '[{"name": "S", "columns": [{"header": "A"}, {"header": "B", "type": "number"}], "rows": [["x", 1]}]'
    assert excel(sheets)[0]["S"]["B2"].value == 1  # type: ignore[arg-type]


def test_parameters_written_into_the_json_text_are_used() -> None:
    blocks = '[{"type": "paragraph", "text": "Body"}], "title": "Quartalsbericht", "landscape": true}'
    doc, filename, _ = word(blocks=blocks)
    assert filename == "Quartalsbericht.docx"
    assert doc.sections[0].page_width > doc.sections[0].page_height


def test_bad_numbers_and_nesting_are_reported() -> None:
    assert "NaN is not a valid number" in invalid(
        "create_powerpoint_file",
        slides='[{"layout": "chart", "title": "C", "chart": {"categories": ["a"], '
        '"series": [{"name": "s", "values": [NaN]}]}}]',
    )
    assert "finite number" in invalid(
        "create_powerpoint_file",
        slides=[
            {
                "layout": "chart",
                "title": "C",
                "chart": {"categories": ["a"], "series": [{"name": "s", "values": [float("inf")]}]},
            }
        ],
    )
    assert "nested too deeply" in invalid("create_word_file", blocks="[" * 100_000 + "]" * 100_000)
    assert "could not be read" in invalid(
        "create_word_file", blocks='[{"type": "paragraph", "text": "x", "level": ' + "9" * 5000 + "}]"
    )
    encoded_twice = json.dumps(json.dumps([{"type": "paragraph", "text": "x"}]))
    assert word(blocks=encoded_twice, title="T")[0].paragraphs[1].text == "x"


def test_inline_markdown_has_no_catastrophic_backtracking() -> None:
    start = time.perf_counter()
    for text in ("**" + "*a" * 5000 + "*", " __a" * 20_000, "**a " * 20_000):
        od.parse_inline(text)
    assert time.perf_counter() - start < 2
    assert [(r.text, r.bold, r.italic) for r in od.parse_inline("**a *b***")] == [
        ("a ", True, False),
        ("b", True, True),
    ]


def test_chart_labels_stay_text_in_the_embedded_workbook() -> None:
    chart = {"categories": ['=WEBSERVICE("https://evil.example/")', "{=1+1}", "https://x.example"],
             "series": [{"name": "=1+1", "values": [1, 2, 3]}]}  # fmt: skip
    result, recorder = call("create_powerpoint_file", slides=[{"layout": "chart", "title": "C", "chart": chart}])
    assert result.startswith("Done"), result
    package = zipfile.ZipFile(io.BytesIO(recorder.download()[0]))
    (embedded,) = [n for n in package.namelist() if n.startswith("ppt/embeddings/")]
    workbook = zipfile.ZipFile(io.BytesIO(package.read(embedded)))
    sheet = workbook.read("xl/worksheets/sheet1.xml").decode()
    assert "<f" not in sheet and "hyperlink" not in sheet.lower()
    assert "WEBSERVICE" in workbook.read("xl/sharedStrings.xml").decode()  # kept, as text


def test_excel_numbers_totals_and_text_safety() -> None:
    sheet = {
        "name": "History",
        "columns": [
            {"header": "**Item**"},
            {"header": "Qty", "total": "sum"},
            {"header": "Code"},
            {"header": "Note", "type": "text", "total": "count"},
        ],
        "rows": [["A", 2, "01234", "=1+1"], ["B", 3.5, "00042", None], ["C", "4", "07743", "ok"]],
    }  # fmt: skip
    ws = excel([sheet])[0]["History_"]
    assert ws["A1"].value == "Item"  # Markdown markers removed, Excel cells have no inline formatting
    assert [ws[f"B{r}"].value for r in (2, 3, 4)] == [2, 3.5, 4]  # an untyped column of numbers holds numbers
    assert ws["C2"].value == "01234"  # leading zeros: a code, stays text
    assert ws["D2"].value == "=1+1" and ws["D2"].data_type == "s" and ws["D2"].quotePrefix  # a 'text' column
    assert ws["B5"].value == "=SUM(B2:B4)" and ws["D5"].value == "=COUNTA(D2:D4)" and ws["D5"].number_format == "0"
    assert ws["A5"].value == "Total"
    long_name = {"name": "A" * 30 + "'b", "columns": [{"header": "x"}], "rows": [["y"]]}
    assert excel([long_name])[0].sheetnames == ["A" * 30]  # no trailing apostrophe after cutting
    assert excel([{"columns": [{"header": "x"}], "rows": [["y"]]}])[0].sheetnames == ["Sheet1"]  # name left out


def test_excel_rejects_content_that_cannot_be_meant() -> None:
    shifted = {
        "name": "Budget",
        "columns": [{"header": "Einnahmen", "type": "currency"}, {"header": "Ausgaben", "type": "currency"},
                    {"header": "Saldo", "type": "currency", "formula": "=A{row}-B{row}"}],
        "rows": [["Miete", 3500, 3500], ["Essen", 500, 480], ["Auto", 800, 750]],
    }  # fmt: skip
    assert "('Einnahmen'): 3 of 3 values are not numbers (e.g. 'Miete')" in invalid(
        "create_excel_file", sheets=[shifted]
    )
    no_prices = {
        "name": "Angebot",
        "columns": [{"header": "Position"}, {"header": "Menge"}, {"header": "Einzelpreis"},
                    {"header": "Gesamt", "type": "currency", "formula": "=C{row}*B{row}"},
                    {"header": "Netto", "type": "currency", "total": "sum"}],
        "rows": [["Server", 2, None, None], ["Lizenzen", 10, None, None]],
    }  # fmt: skip
    result = invalid("create_excel_file", sheets=[no_prices])
    assert "columns[2] ('Einzelpreis') is empty, but the formula of 'Gesamt' uses it" in result
    assert "columns[4] ('Netto'): the total needs values" in result
    text_chart = {"name": "S", "columns": [{"header": "A"}, {"header": "B"}], "rows": [["x", "y"]],
                  "chart": {"category_column": "A", "value_columns": ["B"]}}  # fmt: skip
    assert "'B' has no numbers to plot" in invalid("create_excel_file", sheets=[text_chart])
    too_big = {"name": "S", "columns": [{"header": str(i)} for i in range(50)], "rows": [[]] * 6000}
    assert "cells in all sheets together" in invalid("create_excel_file", sheets=[too_big])


def test_excel_formulas_written_into_cells() -> None:
    sheet = {
        "name": "Budget",
        "columns": [{"header": "Monat"}, {"header": "Ein", "type": "currency"}, {"header": "Aus", "type": "currency"},
                    {"header": "Saldo", "type": "currency"}],
        "rows": [["Jan", 4000, 2500, "=B2-C2"], ["Feb", 4200, 2600, "=B3-C3"],
                 ["Summe", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]],
    }  # fmt: skip
    ws = excel([sheet])[0]["Budget"]
    assert [ws[f"D{r}"].value for r in (2, 3)] == ["=B2-C2", "=B3-C3"] and ws["D2"].data_type == "f"
    assert [c.value for c in ws[4]] == ["Summe", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]  # sums in the cells
    mixed = {**sheet, "rows": [["Jan", 4000, 2500, "=B2-C2"], ["Feb", 4200, 2600, "=SUM(B2:C3)"]]}
    ws = excel([mixed])[0]["Budget"]  # no common pattern: each cell keeps its own formula
    assert (ws["D2"].value, ws["D3"].value) == ("=B2-C2", "=SUM(B2:C3)") and ws["D3"].data_type == "f"
    blocked = {**sheet, "rows": [["Jan", 4000, 2500, '=WEBSERVICE("http://x")'], ["Feb", 4200, 2600, "=B3-C3"]]}
    assert "the formula '=WEBSERVICE" in invalid("create_excel_file", sheets=[blocked])
    months = {"name": "M", "columns": [{"header": "Posten"}, {"header": "Jan"}, {"header": "Feb"}],
              "rows": [["Einnahmen", 3000, 3100], ["Ausgaben", 2500, 2600], ["Saldo", "=B2-B3", "=C2-C3"],
                       ['=HYPERLINK("http://x")', 1, 2]]}  # fmt: skip
    ws = excel([months])[0]["M"]
    assert (ws["B4"].value, ws["C4"].value, ws["B4"].data_type) == ("=B2-B3", "=C2-C3", "f")
    assert ws["A5"].value == '=HYPERLINK("http://x")' and ws["A5"].data_type == "s"  # not allowed: stays text


def test_percent_cells_hold_the_fraction() -> None:
    """Percent numbers go in (5 means 5 %), the cell holds 0.05: fractions sent in, '/100' on a percent cell and
    '*100' in a percent column are rejected with a hint."""
    offer = {
        "name": "Angebot",
        "columns": [{"header": "Artikel"}, {"header": "Menge", "type": "integer"},
                    {"header": "Preis", "type": "currency"}, {"header": "Rabatt", "type": "percent"},
                    {"header": "Netto", "type": "currency", "formula": "=B{row}*C{row}*(1-D{row})"}],
        "rows": [["Stuhl", 6, 240, 5], ["Tisch", 2, 890, "7,5 %"], ["Lampe", 3, 45, 0.5]],
    }  # fmt: skip
    ws = excel([offer])[0]["Angebot"]
    assert (ws["D2"].value, ws["D3"].value, ws["D4"].value) == (0.05, 0.075, 0.005)

    def errors(**changes: Any) -> str:
        return invalid("create_excel_file", sheets=[{**offer, **changes}])

    message = errors(rows=[["Stuhl", 6, 240, 0.05], ["Tisch", 2, 890, "0,1"], ["Montage", 1, 80, 0]])
    assert "('Rabatt'): all values are below 1 (e.g. 0.05)" in message and "'0.5 %'" in message
    assert excel([{**offer, "rows": [["Stuhl", 6, 240, "0,5 %"], ["Tisch", 2, 890, "0.25 %"]]}])[0]["Angebot"]
    divided = [*offer["columns"][:4], {**offer["columns"][4], "formula": "=B{row}*C{row}*(1-D{row}/100)"}]
    assert "D{row} is a percent cell ('Rabatt'), which already holds the fraction (5 % is 0.05): write D{row} " \
           "instead of D{row}/100" in errors(columns=divided)
    conditions = {"name": "Konditionen", "columns": [{"header": "Angabe"}, {"header": "Wert", "type": "percent"}],
                  "rows": [["Rabatt", 5], ["MwSt", 19]]}  # fmt: skip
    total = {"name": "Summe", "columns": [{"header": "Posten"}, {"header": "Betrag", "type": "currency"}],
             "rows": [["Netto", 1000], ["Rabatt", "=-B2*Konditionen!B2/100"]]}  # fmt: skip
    message = invalid("create_excel_file", sheets=[total, conditions])
    assert "sheets[0].rows[1]: 'Konditionen'!B2 is a percent cell ('Wert')" in message
    shares = {"name": "Umlage", "columns": [{"header": "Abteilung"}, {"header": "Fläche", "type": "number"},
              {"header": "Anteil", "type": "percent", "formula": "=B{row}/SUM($B$2:$B$3)*100"}],
              "rows": [["Vertrieb", 120], ["Lager", 380]]}  # fmt: skip
    message = invalid("create_excel_file", sheets=[shares])
    assert "'Anteil' is a percent column" in message and "drop the *100" in message
    shares["columns"][2]["formula"] = "=B{row}/SUM($B$2:$B$3)"
    assert excel([shares])[0]["Umlage"]["C2"].value == "=B2/SUM($B$2:$B$3)"
    per_hundred = {**shares, "columns": [*shares["columns"][:2], {"header": "Je 100 m²", "type": "number",
                   "formula": "=B{row}/100"}]}  # fmt: skip
    assert excel([per_hundred])[0]["Umlage"]["C2"].value == "=B2/100"  # /100 of a number column is fine


def test_html_breaks_and_header_only_tables() -> None:
    doc, _, _ = word(blocks=[{"type": "table", "rows": [["A", "B"], ["x<br>y", "z"]]}], title="T")
    assert doc.tables[0].rows[1].cells[0].text == "x\ny"
    header_only = [{"layout": "table", "title": "Zeitplan", "table": [["Phase", "Zeitraum"]]}]
    assert "only a header row" in invalid("create_powerpoint_file", slides=header_only)


def test_findings_of_the_quality_review() -> None:
    # Markdown table written into a paragraph becomes a real table
    text = "Kosten:\n| Posten | Betrag |\n| :--- | ---: |\n| Lizenz | 4.800 € |\n| Einrichtung | 2.000 € |\nEnde."
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": text}], title="T")
    assert [[c.text for c in row.cells] for row in doc.tables[0].rows] == [
        ["Posten", "Betrag"], ["Lizenz", "4.800 €"], ["Einrichtung", "2.000 €"]
    ]  # fmt: skip
    # a heading that repeats the title in short is left out; no empty page at the end
    blocks = [{"type": "heading", "text": "Protokoll der Teambesprechung"}, {"type": "paragraph", "text": "x"},
              {"type": "page_break"}]  # fmt: skip
    doc, _, _ = word(blocks=blocks, title="Protokoll der Teambesprechung vom 12.10.2026")
    assert [p.text for p in doc.paragraphs] == ["Protokoll der Teambesprechung vom 12.10.2026", "x"]
    assert "w:br" not in doc.element.body.xml or 'w:type="page"' not in doc.element.body.xml
    # a short document gets no page break after the table of contents
    doc, _, _ = word(blocks=[{"type": "heading", "text": "A"}, {"type": "paragraph", "text": "y" * 150}],
                     title="T", table_of_contents=True)  # fmt: skip
    assert 'w:type="page"' not in doc.element.body.xml
    # a table of contents without chapters is sent back
    assert "table of contents but no chapters" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "text": "Autor: [Name]"}], table_of_contents=True
    )
    # all items written with "- " are top-level items, not sub items
    prs, _ = powerpoint([{"layout": "bullets", "title": "Ziele", "bullets": ["- Effizienz", "- Qualität"]}], title="D")
    assert [p.level for p in prs.slides[0].placeholders[1].text_frame.paragraphs] == [0, 0]
    # a formula that uses its own column; an empty "Netto-Summe" row is sent back: "total" builds the row
    own = {"name": "S", "columns": [{"header": "Menge", "type": "number"}, {"header": "Preis", "type": "currency"},
           {"header": "Gesamt", "type": "currency", "formula": "=B{row}*C{row}"}],
           "rows": [[1, 2], [3, 4]]}  # fmt: skip
    assert "uses its own column C" in invalid("create_excel_file", sheets=[own])
    total_column = {"header": "Gesamt", "type": "currency", "formula": "=A{row}*B{row}"}
    offer = {**own, "columns": [*own["columns"][:2], total_column],
             "rows": [[1, 2], [3, 4], ["Netto-Summe", None, None]]}  # fmt: skip
    assert "rows[2] ('Netto-Summe') is a total row: remove it and set \"total\"" in invalid(
        "create_excel_file", sheets=[offer]
    )
    fixed = {**offer, "columns": [*own["columns"][:2], {**total_column, "total": "sum"}], "rows": [[1, 2], [3, 4]]}
    ws = excel([fixed], valves={"language": "de"})[0]["S"]
    assert [c.value for c in ws[4]] == ["Summe", None, "=SUM(C2:C3)"] and ws.max_row == 4
    # unnamed sheets follow the language; chart unit at the value axis
    assert excel([{"columns": [{"header": "x"}], "rows": [["y"]]}], valves={"language": "de"})[0].sheetnames == [
        "Tabelle1"
    ]
    chart = {"categories": ["IT", "HR"], "series": [{"name": "h", "values": [4, 3]}], "unit": "Stunden pro Woche"}
    prs, _ = powerpoint([{"layout": "chart", "title": "Ersparnis", "chart": chart}], title="D")
    (frame,) = [shape for shape in prs.slides[0].shapes if shape.has_chart]
    assert frame.chart.value_axis.axis_title.text_frame.text == "Stunden pro Woche"


def test_escaped_unicode_and_garbled_text() -> None:
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": "x"}], title="T", header_text="Team \\u2013 Protokoll")
    assert doc.sections[0].header.paragraphs[0].text == "Team \N{EN DASH} Protokoll"
    garbled = [{"type": "paragraph", "text": "Phase 1 -0.2 },{text: Phase 2"}]
    assert "the arguments are garbled" in invalid("create_word_file", blocks=garbled)
    assert "the arguments are garbled" in invalid("create_word_file", blocks=[{"{type": "heading", "text": "x"}])


def test_slide_table_columns_fit_long_words() -> None:
    table = [["Phase", "Aktivität", "Zeitraum", "Verantwortlich"],
             ["Optimierung", "Unternehmensweite Einführung", "Monat 4-5", "Projektleitung"]]  # fmt: skip
    prs, _ = powerpoint([{"layout": "table", "title": "Zeitplan", "table": table}], title="D")
    (shape,) = [s for s in prs.slides[0].shapes if s.has_table]
    size = shape.table.cell(1, 0).text_frame.paragraphs[0].runs[0].font.size.pt
    needed = [max(len(w) for row in table for w in row[c].split()) * od.CHAR_WIDTH * size / 11 for c in range(4)]
    assert all(column.width >= need for column, need in zip(shape.table.columns, needed, strict=True))


def test_escaped_line_breaks() -> None:
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": "**Datum:** 12.10.\\n**Ort:** Raum 2"}], title="T")
    assert doc.paragraphs[1].text == "Datum: 12.10.\nOrt: Raum 2"
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": "Pfad C:\\neu\\Ordner"}], title="T")
    assert doc.paragraphs[1].text == "Pfad C:\\neu\\Ordner"  # other backslashes: left as it is


def test_excel_chart_with_only_type_and_title() -> None:
    def plotted(columns: list[Any], rows: list[Any]) -> tuple[list[str], str]:
        (chart,) = excel([sheet(columns, rows, chart={"type": "column", "title": "Umsatz"})])[0]["S"]._charts
        labels = chart.series[0].cat
        return [s.val.numRef.f for s in chart.series], (labels.strRef or labels.numRef).f

    revenue = {"header": "Umsatz", "type": "currency"}
    assert plotted(["Monat", revenue], [["Jan", 10], ["Feb", 12]]) == (["'S'!$B$2:$B$3"], "'S'!$A$2:$A$3")
    # the labels are the first text column, not the position numbers; years and numbers are no values
    products = plotted(["Pos.", "Produkt", revenue], [[1, "Tisch", 500], [2, "Stuhl", 300]])
    assert products == (["'S'!$C$2:$C$3"], "'S'!$B$2:$B$3")
    regions = plotted(["Region", "Jahr", revenue], [["Nord", 2025, 500], ["Süd", 2025, 300]])
    assert regions == (["'S'!$C$2:$C$3"], "'S'!$A$2:$A$3")
    dates = [{"header": "Datum", "type": "date"}, "Kategorie", revenue]
    assert plotted(dates, [["2026-01-01", "a", 1], ["2026-02-01", "b", 2]])[1] == "'S'!$A$2:$A$3"
    assert plotted(["Jahr", revenue], [[2025, 500], [2026, 300]]) == (["'S'!$B$2:$B$3"], "'S'!$A$2:$A$3")


@pytest.mark.parametrize(
    ("text", "decimal_comma", "expected"),
    [
        ("0,125", None, 0.125),  # thousands never follow a leading 0
        ("-0.750", None, -0.75),
        ("1,799 €", True, 1.799),  # German: a single comma is the decimal mark
        ("1.500", True, 1500),
        ("1.500", False, 1.5),  # English: a single dot is the decimal mark
        ("1,799", False, 1799),
        ("1.234.567", True, 1234567),
    ],
)
def test_parse_number_text_decimal_mark(text: str, decimal_comma: bool | None, expected: float) -> None:
    result = od.parse_number_text(text, decimal_comma)
    assert result is not None and result[0] == pytest.approx(expected)


def test_word_numbering_and_text_lines() -> None:
    text = "Termine:\n3. Oktober ist Feiertag.\n24. Dezember ist frei."
    items = ["One", "- a", "- b", "Two", "- c"]
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": text}, {"type": "numbered", "items": items}], title="T")
    assert [(p.style.name, p.text) for p in doc.paragraphs[1:2]] == [
        ("Normal", "Termine:\n3. Oktober ist Feiertag.\n24. Dezember ist frei.")
    ]
    sub_lists = [p._p.pPr.numPr.numId.val for p in doc.paragraphs if p.style.name == "List Number 2"]
    assert len(sub_lists) == 3 and sub_lists[0] == sub_lists[1] != sub_lists[2]  # each sub list restarts at 1


def test_word_title_after_spacing_and_error_positions() -> None:
    blocks = [
        {"type": "paragraph", "text": ""},
        {"type": "heading", "text": "Report"},
        {"type": "paragraph", "text": "x"},
    ]
    doc, _, _ = word(blocks=blocks)
    assert [p.text for p in doc.paragraphs] == ["Report", "x"]  # the title is not repeated
    blocks = [
        {"type": "heading", "text": "Report"},
        {"type": "paragraph", "text": ""},
        {"type": "table", "rows": [["x"] * 13]},
    ]
    assert "blocks[2] (table): at most 12 columns" in invalid("create_word_file", blocks=blocks)


def test_word_settings_order() -> None:
    blocks = [
        {"type": "table_of_contents"},
        {"type": "heading", "text": "A"},
        {"type": "paragraph", "text": "x" * 120},
        {"type": "table_of_contents"},
    ]
    doc, _, _ = word(blocks=blocks, title="T")
    tags = [child.tag.split("}")[1] for child in doc.settings.element]
    assert tags.count("updateFields") == 1 and tags.index("updateFields") < tags.index("compat")


def test_slide_overflow() -> None:
    assert "too much text for one slide" in invalid(
        "create_powerpoint_file", slides=[{"layout": "bullets", "title": "T", "bullets": ["x" * 200] * 12}]
    )
    table = [["Header", "Value"]] + [[f"Row {i}", str(i)] for i in range(15)]
    prs, _ = powerpoint([{"layout": "table", "title": "T", "table": table}], title="Deck")
    (shape,) = [shape for shape in prs.slides[0].shapes if shape.has_table]
    assert shape.top + sum(row.height for row in shape.table.rows) <= od.SLIDE_HEIGHT
    xml = shape._element.xml
    assert not re.search(r'(cy|h)="[0-9]+\.[0-9]', xml)  # whole numbers only, as the file format requires


def test_safe_file_names_and_links() -> None:
    assert od.safe_filename("x‮fdp.docx", "D") == "xfdp.docx"
    assert od.safe_filename("CON", "D") == "CON_"
    assert od.is_safe_url("mailto:a@b.example?subject=Hi")
    assert not od.is_safe_url("mailto:a@b.example?attach=/etc/passwd")


def test_download_that_the_browser_does_not_confirm() -> None:
    async def gone(event: dict[str, Any]) -> Any:
        return {"error": "Client session disconnected."}

    tools = make_tools()
    result = asyncio.run(
        tools.create_word_file(blocks=[{"type": "paragraph", "text": "x"}], title="T", __event_call__=gone)
    )
    assert result == od.RESULT_FAILED


# ------------------------------------------------------------------------------- second review (0.3.0) --


def test_escaped_line_breaks_take_linear_time() -> None:
    """'\\n' * n + '\\x' took seconds with a regex over the whole text (holding the GIL for all of Open WebUI)."""
    for text in ("\\n" * 30_000 + "\\x", "\\u2013" * 20_000 + "\\x", "\\n" * 30_000):
        start = time.perf_counter()
        od._unescape_newlines(text)
        assert time.perf_counter() - start < 0.2
    assert od._unescape_newlines("a\\nb") == "a\nb" and od._unescape_newlines("C:\\neu\\x") == "C:\\neu\\x"


def sheet(columns: list[Any], rows: list[Any], **extra: Any) -> dict[str, Any]:
    return {"name": "S", "columns": [c if isinstance(c, dict) else {"header": c} for c in columns], "rows": rows,
            **extra}  # fmt: skip


ARTICLES = [{"header": "Artikel"}, {"header": "Menge", "type": "integer"}, {"header": "Preis", "type": "currency"}]


def with_total(formula: str, rows: list[Any] | None = None, **column: Any) -> dict[str, Any]:
    total = {"header": "Summe", "type": "currency", "formula": formula, **column}
    return sheet([*ARTICLES, total], rows or [["Stuhl", 6, 240], ["Tisch", 2, 890], ["Lampe", 3, 45]])


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=B2/SUM(B2:B4)", "=B{row}/SUM($B$2:$B$4)"),  # the share of each row, not of B2 in every row
        ("=SUM(B2:B{row})", "=SUM($B$2:B{row})"),  # a running total from the first data row
        ("=SUM(D2*F2,D3*F3)", "=SUM(D2*F2,D3*F3)"),  # cells of several rows: fixed cells, as written
        ("=SUMME(B2:D2)", "=SUM(B{row}:D{row})"),  # a range within the first data row
        ("=SUM($B$2:B9)", "=SUM($B$2:$B$9)"),  # one end fixed, the other a fixed number: a fixed range
        ("=SUM(B2:$B$9)", "=SUM($B$2:$B$9)"),
        ("=SUM($B$2:B{row})", "=SUM($B$2:B{row})"),  # a running total stays
    ],
)
def test_column_formula_rows(formula: str, expected: str) -> None:
    assert od.normalize_formula(formula) == expected


@pytest.mark.parametrize(
    ("formula", "message"),
    [
        ("=B3*C3", "B3, C3 in a column formula"),  # one row off
        ("=B1*C1", "(row 1 is the header row)"),
        ("=B{row}*C2", "write C{row} for the cell in the same row (e.g. =B{row}*C{row}), $C$2 for one fixed cell"),
    ],
)
def test_column_formula_without_row_is_sent_back(formula: str, message: str) -> None:
    with pytest.raises(ValueError, match=re.escape(message)):
        od.normalize_formula(formula)


def test_formula_on_a_text_column() -> None:
    """Miscounted letters: '=A{row}*B{row}' with the article names in A gives #VALUE! in every row."""
    result = invalid("create_excel_file", sheets=[with_total("=A{row}*B{row}")])
    assert "the formula computes with A{row}, but 'Artikel' holds text (e.g. 'Stuhl'), which gives #VALUE!" in result
    assert "(columns of 'S': A=Artikel, B=Menge, C=Preis, D=Summe)" in result
    ws = excel([with_total('=IF(A{row}="Stuhl",B{row}*C{row}*0.9,B{row}*C{row})')])[0]["S"]  # comparisons are fine
    assert ws["D2"].value == '=IF(A2="Stuhl",B2*C2*0.9,B2*C2)'
    times = sheet(["Tag", "Beginn", "Ende", {"header": "Dauer", "formula": "=C{row}-B{row}"}],
                  [["Mo", "08:00", "16:30"]])  # fmt: skip
    assert excel([times])[0]["S"]["D2"].value == "=C2-B2"  # Excel computes with times written as text too
    header = sheet(["A", {"header": "B", "type": "number"}], [["x", 2], ["y", "=B2*$B$1"]])
    assert "$B$1 is in the header row ('B')" in invalid("create_excel_file", sheets=[header])
    single = with_total("=B{row}*C{row}", [["Stuhl", 6, 240], ["Tisch", "-", 890]])
    assert "but sheets[0].rows[1] has '-' there, which gives #VALUE!: send a number, or null" in invalid(
        "create_excel_file", sheets=[single]
    )


def test_single_texts_in_number_columns() -> None:
    """A text in a number column that a total, formula or chart uses would silently count as 0."""
    costs = sheet(["Posten", {"header": "Betrag", "type": "currency", "total": "sum"}], [["a", 150], ["b", "ca. 500"]])
    result = invalid("create_excel_file", sheets=[costs])
    assert "sheets[0].rows[1]: 'ca. 500' in 'Betrag' is not a number, so the total cannot count it" in result
    assert "send the number (500)" in result
    costs["rows"][1][1] = "3 h"
    assert "send the number (3) and put the unit into the header, e.g. 'Betrag (h)'" in invalid(
        "create_excel_file", sheets=[costs]
    )
    hours = sheet(["Posten", {"header": "Dauer", "type": "number"}], [["a", "3 h"], ["b", "2 h"], ["c", 1]])
    assert "If they are numbers with a unit: send the number (3)" in invalid("create_excel_file", sheets=[hours])
    costs["rows"][1][1] = "offen"  # placeholders stay allowed
    assert excel([costs])[0]["S"]["B3"].value == "offen"
    unused = sheet(["Posten", {"header": "Betrag", "type": "currency"}], [["a", 150], ["b", "ca. 500"]])
    assert excel([unused])[0]["S"]["B3"].value == "ca. 500"  # nothing uses the column: stays as sent


@pytest.mark.parametrize(
    ("formula", "written"),
    [
        ("=B{row}*(1-D{row}%)", "D{row}%"),
        ("=B{row}*C{row}*(D{row})/100", "D{row}/100"),
        ("=B{row}*C{row}*D{row}*0.01", "D{row}*0.01"),
    ],
)
def test_percent_cells_divided_again(formula: str, written: str) -> None:
    offer = sheet([*ARTICLES, {"header": "Rabatt", "type": "percent"}, {"header": "Netto", "formula": formula}],
                  [["Stuhl", 6, 240, 5]])  # fmt: skip
    result = invalid("create_excel_file", sheets=[offer])
    assert "D{row} is a percent cell ('Rabatt'), which already holds the fraction" in result
    assert f"instead of {written}" in result


def test_list_marks_written_by_the_model() -> None:
    blocks = [
        {"type": "bullets", "items": ["• Punkt", "• Zweiter", "- Unterpunkt"]},
        {"type": "numbered", "items": ["1. Erstens", "2) Zweitens", "3. Drittens"]},
        {"type": "bullets", "items": ["* Stern", "Ohne"]},  # not every item: left as it is
        {"type": "numbered", "items": ["3. Oktober", "5. November"]},  # no counting from 1: left as it is
    ]
    doc, _, _ = word(blocks=blocks, title="T")
    assert [p.text for p in doc.paragraphs[1:]] == [
        "Punkt", "Zweiter", "Unterpunkt", "Erstens", "Zweitens", "Drittens", "* Stern", "Ohne", "3. Oktober",
        "5. November",
    ]  # fmt: skip
    prs, _ = powerpoint([{"layout": "bullets", "title": "Ziele", "bullets": ["– Effizienz", "– Qualität"]}], title="D")
    assert [p.text for p in prs.slides[0].placeholders[1].text_frame.paragraphs] == ["Effizienz", "Qualität"]


def test_markdown_headings_and_tables_in_paragraphs() -> None:
    text = "## Abschnitt\nText darunter.\n\n---\n\n### Unterpunkt\n#hashtag bleibt"
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": text}], title="T")
    assert [(p.style.name, p.text) for p in doc.paragraphs[1:]] == [
        ("Heading 2", "Abschnitt"), ("Normal", "Text darunter."), ("Heading 3", "Unterpunkt"),
        ("Normal", "#hashtag bleibt"),
    ]  # fmt: skip
    wide = (
        "| " + " | ".join(f"S{i}" for i in range(15)) + " |\n|" + " --- |" * 15 + "\n| " + " | ".join("x" * 15) + " |"
    )
    assert "blocks[0] (paragraph): at most 12 columns (the Markdown table in the text)" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "text": wide}], title="T"
    )
    long = "| A | B |\n| --- | --- |\n" + "| x | y |\n" * 600
    assert "at most 500 rows (the Markdown table in the text)" in invalid(
        "create_word_file", blocks=[{"type": "paragraph", "text": long}], title="T"
    )


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=SVERWEIS(A{row};$E$2:$F$9;2;FALSCH)", "=VLOOKUP(A{row},$E$2:$F$9,2,FALSE)"),
        ("=SUMMEWENNS($C$2:$C$9;$A$2:$A$9;A{row})", "=SUMIFS($C$2:$C$9,$A$2:$A$9,A{row})"),
        ("=NETTOARBEITSTAGE(B{row};C{row})", "=NETWORKDAYS(B{row},C{row})"),
        ("=EDATUM(B{row};3)+MONATSENDE(B{row};0)-HEUTE()", "=EDATE(B{row},3)+EOMONTH(B{row},0)-TODAY()"),
        ("=WENN(B{row}>5;„ja“;“nein”)", '=IF(B{row}>5,"ja","nein")'),  # German quotes in a formula
        ("=XVERWEIS(A{row};$E$2:$E$9;$F$2:$F$9)", "=_xlfn.XLOOKUP(A{row},$E$2:$E$9,$F$2:$F$9)"),  # Excel 2019+
        ("=_xlfn.MAXIFS($C$2:$C$9,$A$2:$A$9,A{row})", "=_xlfn.MAXIFS($C$2:$C$9,$A$2:$A$9,A{row})"),
        ('=TEXTVERKETTEN(", ";WAHR;A{row};B{row})', '=_xlfn.TEXTJOIN(", ",TRUE,A{row},B{row})'),
    ],
)
def test_formula_whitelist_new_functions(formula: str, expected: str) -> None:
    assert od.normalize_formula(formula) == expected


@pytest.mark.parametrize(
    ("formula", "message"),
    [
        ("=B2*", "the formula ends with '*': a value is missing after it"),
        ("=B2*/2", "'/' follows '*' without a value between them"),
        ("=SUM()", "SUM has no arguments here; write it as SUM(B2:B9)"),
        ("==B2", "the formula starts with '=='"),
        ("=ROUND(B2)", "ROUND has 1 argument here; write it as ROUND(value, digits)"),
        ("=IF(,1,2)", "an argument is missing before ','"),  # only arguments after the first may be left out
        ("=(B2,C2)", "',' separates the arguments of a function"),
        ("=SUMIFS(C2:C9,A2:A9)", "SUMIFS has 2 arguments here"),
        ("=2B2", "'B2' follows '2' without an operator"),
        ('=IF(B2>0;";1)', "a text in the formula is not closed"),
        ("=SUMIF(A:A,1)", "whole columns (A:A) are not allowed: write the rows, e.g. $A$2:$A$50"),  # no sheet known
        ("=C{row-1}+B{row}", "{row-1} (another row than the current one) is not possible in a column formula"),
        ("=TIMEVALUE(C{row})", "give the column with the times the type time"),
        ('=INDIRECT("A1")', "the function INDIRECT is not allowed"),
        ("=" + "(" * 40 + "1" + ")" * 40, "nested too deeply"),
    ],
)
def test_formula_syntax_is_checked(formula: str, message: str) -> None:
    with pytest.raises(ValueError, match=re.escape(message)):
        od.normalize_formula(formula)


def test_whole_columns_become_the_data_rows() -> None:
    data = sheet(["Abteilung", {"header": "Kosten", "type": "currency"}], [["IT", 10], ["HR", 20], ["IT", 5]])
    summary = {"name": "Summe", "columns": [{"header": "Abteilung"}, {"header": "Kosten", "type": "currency",
               "formula": "=SUMIF(S!A:A,A{row},S!$B:$B)"}, {"header": "Anteil", "type": "percent",
               "formula": "=B{row}/SUM(B:B)"}], "rows": [["IT"], ["HR"]]}  # fmt: skip
    ws = excel([data, summary])[0]["Summe"]
    assert ws["B2"].value == "=SUMIF('S'!$A$2:$A$4,A2,'S'!$B$2:$B$4)" and ws["C3"].value == "=B3/SUM($B$2:$B$3)"


def test_whole_columns_in_a_cell_end_above_it() -> None:
    """'=SUM(B:B)' in the total row would add up the cell itself (a circular reference)."""
    total = sheet(["A", {"header": "B", "type": "number"}], [["Stuhl", 240], ["Tisch", 500], ["Summe", "=SUM(B:B)"]])
    assert excel([total])[0]["S"]["B4"].value == "=SUM($B$2:$B$3)"
    share = sheet(["A", {"header": "B", "type": "number"}, {"header": "C", "type": "percent"}],
                  [["x", 1, "=B2/SUM(B:B)"], ["y", 3, None]])  # fmt: skip
    assert excel([share])[0]["S"]["C2"].value == "=B2/SUM($B$2:$B$3)"  # another column: all data rows
    between = sheet(["A", {"header": "B", "type": "number"}], [["x", 1], ["Summe", "=SUM(B:B)"], ["y", 3]])
    assert "B:B in B3 includes B3 itself (a circular reference): write the rows to compute with, e.g. $B$2:$B$2" in (
        invalid("create_excel_file", sheets=[between])
    )
    top = sheet(["A", {"header": "B", "type": "number"}], [["Summe", "=SUMME(B:B)"], ["x", 1], ["y", 3]])
    assert "e.g. $B$3:$B$4" in invalid("create_excel_file", sheets=[top])


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=A2&" ""x"""', '=A2&" ""x"""'),  # "" is a " inside a text
        ('=WENN(A2="x";„[neu]“;"")', '=IF(A2="x","[neu]","")'),  # [ in a text with German quotes
        ("=VLOOKUP(A2,$A$2:$C$9,2,)", "=VLOOKUP(A2,$A$2:$C$9,2,)"),  # arguments left out, as Excel allows
        ("=IF(A2,,1)", "=IF(A2,,1)"),
        ("=WENN(A2;;)", "=IF(A2,,)"),
    ],
)
def test_formula_texts_and_left_out_arguments(formula: str, expected: str) -> None:
    assert od.normalize_formula(formula, False) == expected
    for unsafe in ("=[Book]S!A1", "='[Book]S'!A1", "=cmd|' /C calc'!A0", '=„x“&[Book]S!A1'):
        with pytest.raises(od.UnsafeFormula):
            od.normalize_formula(unsafe, False)
    with pytest.raises(ValueError, match=re.escape("ROUND has 3 arguments here")):  # left out ones count
        od.normalize_formula("=ROUND(A2,,)", False)


def test_rows_as_objects_written_into_the_chart() -> None:
    chart = {"type": "column", "category_column": "Artikel", "value_columns": ["Preis"],
             "rows": [{"Artikel": "Stuhl", "Preis": 240}, {"Artikel": "Tisch", "Preis": 500}]}  # fmt: skip
    columns = [{"header": "Artikel"}, {"header": "Preis", "type": "number"}]
    ws = excel([{"name": "S", "columns": columns, "chart": chart}])[0]["S"]
    assert (ws["A3"].value, ws["B3"].value, len(ws._charts)) == ("Tisch", 500, 1)


def test_rejected_formulas_in_cells() -> None:
    """A formula in a cell that is merely not allowed or broken comes back (it would be written as text); one that
    reaches outside (HYPERLINK ...) stays text, a column of type text keeps everything as text."""
    broken = sheet(["A", "B"], [["x", "=SUMMENPRODUKT(B2:B3;C2:C3)"]])
    assert (
        "sheets[0].rows[0]: the formula '=SUMMENPRODUKT(B2:B3;C2:C3)' in 'B' is not allowed: the function "
        "SUMMENPRODUKT is not allowed: add a column with the product" in invalid("create_excel_file", sheets=[broken])
    )
    safe = sheet(["A", "B", {"header": "C", "type": "text"}],
                 [["x", '=HYPERLINK("http://x")', "=SUMME(B2:"], ["y", "=> siehe unten", "=foo("]])  # fmt: skip
    ws = excel([safe])[0]["S"]
    assert ws["B2"].value == '=HYPERLINK("http://x")' and ws["B2"].data_type == "s"
    assert (ws["B3"].value, ws["C2"].value, ws["C2"].data_type) == ("=> siehe unten", "=SUMME(B2:", "s")


def test_dates_and_times() -> None:
    dates = sheet(["A", {"header": "Datum", "type": "date"}], [["x", "22.10.26"], ["y", "2026-10-22 14:30"]])
    ws = excel([dates], valves={"language": "de"})[0]["S"]
    assert ws["B2"].value.date() == date(2026, 10, 22) and ws["B3"].value.hour == 14
    assert ws["B2"].number_format == "DD.MM.YYYY hh:mm"
    for value, message in [("05/11/2026", "could be day/month or month/day"), ("2026-11", "has no day"),
                           ("31.11.2026", "is not a valid date"), ("14:30", "is a time")]:  # fmt: skip
        result = invalid(
            "create_excel_file", sheets=[sheet(["A", {"header": "Datum", "type": "date"}], [["x", value]])]
        )
        assert f"sheets[0].rows[0]: '{value}' {message}" in result
    untyped = sheet(["A", "Datum", "Beginn"], [["x", "2026-11-02", "08:00"], ["y", "2026-11-03", "8:15 Uhr"]])
    ws = excel([untyped])[0]["S"]
    assert ws["B2"].value.date() == date(2026, 11, 2) and ws["C3"].number_format == "hh:mm"
    work = sheet([{"header": "Beginn", "type": "time"}, {"header": "Ende", "type": "time"},
                  {"header": "Dauer", "type": "time", "formula": "=B{row}-A{row}", "total": "sum"},
                  {"header": "Stunden", "type": "number", "formula": "=(B{row}-A{row})*24", "total": "sum"}],
                 [["08:00", "16:30"], ["08:15", "17:00"]])  # fmt: skip
    wb = excel([work])[0]
    ws = wb["S"]
    assert ws["A2"].value.hour == 8 and ws["A2"].number_format == "hh:mm"
    assert (ws["C2"].number_format, ws["C4"].value, ws["C4"].number_format) == ("[h]:mm", "=SUM(C2:C3)", "[h]:mm")
    trap = {**work, "columns": [*work["columns"][:2], {"header": "Stunden", "type": "number",
            "formula": "=B{row}-A{row}"}]}  # fmt: skip
    assert "which Excel holds as fractions of a day (8 hours are 0.33), so this number column would show 0.35" in (
        invalid("create_excel_file", sheets=[trap])
    )
    hours = {**work, "columns": [*work["columns"][:2], {"header": "Dauer", "type": "time",
             "formula": "=(B{row}-A{row})*24"}]}  # fmt: skip
    assert "multiplies the times by 24, so 8:30 would show as 204:00" in invalid("create_excel_file", sheets=[hours])
    assert "write times as hh:mm" in invalid("create_excel_file", sheets=[{**work, "rows": [[8, "16:30"]]}])


def test_currencies_of_the_values() -> None:
    prices = sheet(["A", {"header": "Preis", "type": "currency", "total": "sum"},
                    {"header": "Brutto", "type": "currency", "formula": "=B{row}*1.19"}],
                   [["x", "$1,200.00"], ["y", "$5.50"]])  # fmt: skip
    ws = excel([prices])[0]["S"]
    assert ws["B2"].value == 1200 and ws["B2"].number_format == ws["C2"].number_format == od.CURRENCY_FORMATS["USD"]
    assert ws["B4"].number_format == od.CURRENCY_FORMATS["USD"]  # the total row too
    prices["rows"][1][1] = "5,50 €"
    assert "the values have different currencies (USD, EUR)" in invalid("create_excel_file", sheets=[prices])
    assert od.parse_number_text("1’234.50") == (1234.5, False) and od.parse_number_text("CHF 1'000.00") == (1000, False)


def test_unambiguous_shapes_are_accepted() -> None:
    rows = [{"Artikel": "Stuhl", "Preis (€)": 240}, {"Preis (€)": 890, "Artikel": "Tisch"}]
    objects = sheet(["Artikel", {"header": "Preis (€)", "type": "currency"}], rows)
    ws = excel([objects])[0]["S"]
    assert [[c.value for c in row] for row in ws.iter_rows(min_row=2)] == [["Stuhl", 240], ["Tisch", 890]]
    objects["rows"] = [{"Artikel": "Stuhl", "Preis": 240}]
    assert "rows[0] is an object whose key 'Preis' is not a column header (Artikel, Preis (€))" in invalid(
        "create_excel_file", sheets=[objects]
    )
    doc, _, _ = word(blocks=[{"type": "table", "rows": [["A", "B"], [None, "x"]]}], title="T")
    assert [c.text for c in doc.tables[0].rows[1].cells] == ["", "x"]
    prs, _ = powerpoint([{"layout": "table", "title": "T", "table": [["A", "B"], ["x", None]]}], title="D")
    (shape,) = [s for s in prs.slides[0].shapes if s.has_table]
    assert shape.table.cell(1, 1).text == ""
    wrapped = json.dumps({"sheets": [sheet(["A"], [["x"]])], "title": "Inventar 2026"})
    result, recorder = call("create_excel_file", sheets=wrapped)
    assert result.startswith("Done") and recorder.download()[1] == "Inventar 2026.xlsx"
    lost = json.dumps({"sheets": [sheet(["A"], [["x"]])], "notes": "Bitte prüfen"})
    assert "sheets: 'notes' is not a parameter of this tool, its content would be lost" in invalid(
        "create_excel_file", sheets=lost
    )


def test_formats_of_formula_columns_without_type() -> None:
    plan = sheet(["Aufgabe", {"header": "Start", "type": "date"}, {"header": "Frist", "formula": "=B{row}+14"},
                  {"header": "Menge", "type": "integer"}, {"header": "Preis", "type": "currency"},
                  {"header": "Summe", "formula": "=D{row}*E{row}"}, {"header": "Tage", "formula": "=C{row}-B{row}"}],
                 [["A", "2026-10-22", None, 3, 411.5]])  # fmt: skip
    ws = excel([plan], valves={"language": "de"})[0]["S"]
    assert [ws[f"{c}2"].number_format for c in "CFG"] == ["DD.MM.YYYY", od.CURRENCY_FORMATS["EUR"], "#,##0"]
    wrong = {
        **plan,
        "columns": [*plan["columns"][:2], {"header": "Frist", "type": "currency", "formula": "=B{row}+14"}],
    }
    assert "the formula gives a date, which this currency column would show as a number" in invalid(
        "create_excel_file", sheets=[{**wrong, "rows": [["A", "2026-10-22"]]}]
    )


def test_word_letter_without_title() -> None:
    doc, filename, _ = word(blocks=[{"type": "paragraph", "text": "Sehr geehrte Frau Keller,"}])
    assert [p.text for p in doc.paragraphs] == ["Sehr geehrte Frau Keller,"] and filename == "Document.docx"
    assert not any(p.style.name == "Title" for p in doc.paragraphs)


def test_excel_several_charts_with_units() -> None:
    charts = [{"type": "line", "category_column": "Monat", "value_columns": ["Umsatz"], "unit": "EUR"},
              {"type": "column", "category_column": "Monat", "value_columns": ["Kosten"]}]  # fmt: skip
    data = sheet(["Monat", {"header": "Umsatz", "type": "currency"}, {"header": "Kosten", "type": "currency"}],
                 [["Jan", 10, 4], ["Feb", 12, 5]], chart=charts)  # fmt: skip
    ws = excel([data])[0]["S"]
    assert len(ws._charts) == 2 and [c.anchor._from.row for c in ws._charts] == [1, 1 + od.CHART_ROWS]
    assert ws._charts[0].y_axis.title is not None and ws._charts[1].y_axis.title is None
    result = invalid("create_excel_file", sheets=[{**data, "chart": [*charts, {"value_columns": ["Gewinn"]}]}])
    assert "sheets[0].chart[2].value_columns: 'Gewinn' is not a column header" in result
    assert "at most 4 charts" in invalid("create_excel_file", sheets=[{**data, "chart": charts * 3}])
    for chart, message in [({"type": "pi"}, "sheets[0].chart.type: Input should be"),
                           ([{"type": "pi"}], "sheets[0].chart[0].type: Input should be"),
                           ("Umsatz", "sheets[0].chart: Input should be an object (or a list of objects)")]:  # fmt: skip
        result = invalid("create_excel_file", sheets=[{**data, "chart": chart}])
        assert message in result and "valid list" not in result and "instance of" not in result


def test_values_sent_in_formula_columns() -> None:
    """Values that differ from what the formula computes: wrong letters or wrong values; 0 or the same value in
    every row are placeholders."""
    result = invalid("create_excel_file", sheets=[with_total("=B{row}*C{row}", [["Stuhl", 6, 240, 1440],
                                                                                ["Tisch", 2, 890, 999]])])  # fmt: skip
    assert "sheets[0].rows[1]: the formula =B{row}*C{row} of 'Summe' gives 1780 in this row, but the row sends 999" \
           in result and "A=Artikel, B=Menge" in result  # fmt: skip
    for rows in ([["Stuhl", 6, 240, 0], ["Tisch", 2, 890, 0]], [["Stuhl", 6, 240, 20], ["Tisch", 2, 890, 20]],
                 [["Stuhl", 6, 240, 1440.4], ["Tisch", 2, 890, 1780]]):  # fmt: skip
        assert excel([with_total("=B{row}*C{row}", rows)])[0]["S"]["D2"].value == "=B2*C2"


def test_sum_rows_below_the_items() -> None:
    """Net, VAT and gross below the line items: a formula sent in a row of a formula column replaces the column
    formula there, and the line item formula is left out of such rows instead of showing 0."""
    rows = [["Stuhl", 6, 240, None], ["Tisch", 2, 890, None], ["Netto", None, None, "=SUM(D2:D3)"],
            ["MwSt 19 %", None, None, "=ROUND(D4*0.19,2)"], ["Brutto", None, None, "=D4+D5"]]  # fmt: skip
    ws = excel([with_total("=B{row}*C{row}", rows)])[0]["S"]
    assert [ws[f"D{r}"].value for r in range(2, 7)] == ["=B2*C2", "=B3*C3", "=SUM(D2:D3)", "=ROUND(D4*0.19,2)",
                                                        "=D4+D5"]  # fmt: skip
    discount = sheet([*ARTICLES, {"header": "Rabatt", "type": "percent"},
                      {"header": "Netto", "type": "currency", "formula": "=B{row}*C{row}*(1-D{row})"},
                      {"header": "Hinweis", "formula": '=IF(D{row}>0,"rabattiert","")'}],
                     [["Stuhl", 6, 240, 5, None, None], ["Summe", None, None, None, "=SUM(E2:E2)", None]])  # fmt: skip
    ws = excel([discount])[0]["S"]
    assert ws["E3"].value == "=SUM(E2:E2)" and ws["F3"].value is None  # no IF over the empty row either
    # a last row 'Summe' with its own formulas is the model's own total, not one the tool must build
    assert excel([with_total("=B{row}*C{row}", [["Stuhl", 6, 240], ["Summe", None, None, "=SUM(D2:D2)"]])])


def test_sum_rows_next_to_charts_copies_and_templates() -> None:
    rows = [["Stuhl", 6, 240, "=B2*C2"], ["Tisch", 2, 890, "=B2*C2"], ["Netto", None, None, "=SUM(D2:D3)"]]
    ws = excel([with_total("=B{row}*C{row}", rows, total="none")])[0]["S"]
    assert ws["D3"].value == "=B3*C3"  # a copy of the column formula, even with the wrong row, is the column formula
    chart = {"type": "pie", "category_column": "Artikel", "value_columns": ["Summe"]}
    ws = excel([{**with_total("=B{row}*C{row}", rows[:2] + [rows[2], ["MwSt", None, None, "=D4*0.19"]]),
                 "chart": chart}])[0]["S"]  # fmt: skip
    assert ws.auto_filter.ref == "A1:D3" and "$D$2:$D$3" in str(ws._charts[0].series[0].val.numRef.f)
    discount = [*rows[:2], ["Rabatt 5 %", None, None, "=-ROUND(SUM(D2:D3)*0.05,2)"]]
    assert excel([with_total("=B{row}*C{row}", discount, total="sum")])  # meant to be in the total
    template = sheet(["Angabe", {"header": "Wert", "type": "number"}],
                     [["Stunden (bitte eintragen)", None], ["Satz", 80], ["Betrag", "=B2*B3"]])  # fmt: skip
    assert excel([template])[0]["S"]["B4"].value == "=B2*B3"  # an input left for the user
    loop = sheet(["A", {"header": "B", "formula": "=C{row}+1"}, {"header": "C", "formula": "=B{row}+1"}],
                 [["x", None, None], ["Summe", "=SUM(B2:B2)", "=SUM(C2:C2)"]])  # fmt: skip
    assert "circular" not in str(call("create_excel_file", sheets=[loop])[0])  # formula columns using each other


def test_workbook_without_rows() -> None:
    """A parallel test call or an empty copy of the file would be attached as an empty workbook."""
    empty = invalid("create_excel_file", sheets=[{"name": "Test", "columns": [{"header": "x"}], "rows": []}])
    assert "sheets: no sheet has rows, so the file would be empty" in empty
    template = sheet(["Datum", {"header": "Stunden", "type": "number"}], [[None, None], [None, None]])
    assert excel([template])[0]["S"].max_row == 3  # a template with empty rows is fine
    info = sheet(["Hinweis"], [])
    assert excel([sheet(["A", {"header": "B", "type": "number"}], [["x", 1]]), {**info, "name": "T"}])


def test_sum_rows_that_go_wrong() -> None:
    def problem(rows: list[Any], **column: Any) -> str:
        return invalid("create_excel_file", sheets=[with_total("=B{row}*C{row}", rows, **column)])

    items = [["Stuhl", 6, 240, None], ["Tisch", 2, 890, None]]
    assert "rows[2]: the formula =SUM(D2:D4) in 'Summe' uses its own cell D4 (a circular reference)" in problem(
        [*items, ["Netto", None, None, "=SUM(D2:D4)"]]
    )
    twice = problem([*items, ["davon Stühle", None, None, '=SUMIF(A2:A3,"Stuhl",D2:D3)']], total="sum")
    assert "computes from 'Summe' itself, and the column's total (\"sum\") adds up every row" in twice
    # rows counted wrong: D5 is the empty row between the items and the sums, the net is in D6
    shifted = problem([*items, [None, None, None, None], ["Netto", None, None, "=SUM(D2:D3)"],
                       ["MwSt", None, None, "=D4*0.19"]])  # fmt: skip
    assert "rows[4]: the formula =D4*0.19 uses D4, which is empty, so it computes with 0" in shifted
    assert "(rows: 2=Stuhl, 3=Tisch, 5=Netto, 6=MwSt); is the row number right?" in shifted
    own_column = invalid("create_excel_file", sheets=[with_total('=IF(A{row}="Netto",SUM(D2:D3),B{row}*C{row})')])
    assert "A sum row below the items gets its own formula in its row instead, e.g. '=SUM(D2:D6)'" in own_column


def test_numbers_right_aligned_in_word_and_slide_tables() -> None:
    rows = [["Posten", "Betrag", "Notiz"], ["Lizenz", "4.800 €", "jährlich"], ["Einrichtung", "-", "einmalig"]]
    doc, _, _ = word(blocks=[{"type": "table", "rows": rows}], title="T")
    assert [c.paragraphs[0].alignment for c in doc.tables[0].rows[1].cells] == [None, od.WD_ALIGN_PARAGRAPH.RIGHT,
                                                                               None]  # fmt: skip
    prs, _ = powerpoint([{"layout": "table", "title": "T", "table": rows}], title="D")
    (shape,) = [s for s in prs.slides[0].shapes if s.has_table]
    assert shape.table.cell(0, 1).text_frame.paragraphs[0].alignment == od.PP_ALIGN.RIGHT
    assert shape.table.cell(1, 2).text_frame.paragraphs[0].alignment is None


def test_one_decimal_mark_per_column() -> None:
    """The values that show their decimal mark ('2.350,50') decide it for the whole column, also against the
    language; '1.500' next to them is not read the other way without a word."""
    for language, amounts in (("en", ["1.500 €", "2.350,50 €", "980,00 €"]), ("de", ["$1,500", "$2,350.50", "$980"])):
        rows = [[f"P{i}", amount] for i, amount in enumerate(amounts)]
        ws = excel([sheet(["Posten", {"header": "Betrag", "total": "sum"}], rows)], valves={"language": language})[0]
        assert [ws["S"].cell(r, 2).value for r in (2, 3, 4)] == [1500, 2350.5, 980]
    for language, expected in (("de", [1500, 12000]), ("en", [1.5, 12])):  # nothing shows the mark: the language
        ws = excel([sheet(["A", "B"], [["x", "1.500"], ["y", "12.000"]])], valves={"language": language})[0]["S"]
        assert [ws["B2"].value, ws["B3"].value] == expected
    ws = excel([sheet(["A", "B"], [["x", "1.234.567"], ["y", "2.000"]])], valves={"language": "en"})[0]["S"]
    assert [ws["B2"].value, ws["B3"].value] == [1234567, 2000]  # a separator that comes twice is no decimal mark
    ws = excel([sheet(["A", "B"], [["x", "2.350,50"], ["y", "1,200.50"]])])[0]["S"]
    assert [ws["B2"].value, ws["B3"].value] == [2350.5, 1200.5]  # both marks, but nothing they decide
    mixed = sheet(["A", "Betrag"], [["x", "2.350,50"], ["y", "1,200.50"], ["z", "1.500"]])
    assert (
        "sheets[0].columns[1] ('Betrag'): the values use both ',' and '.' as decimal mark ('2.350,50', '1,200.50'), "
        "so '1.500' could be 1.5 or 1500: send the numbers as JSON numbers"
    ) in invalid("create_excel_file", sheets=[mixed])


@pytest.mark.parametrize(
    ("header", "expected"),
    [("Jahr", True), ("Geschäftsjahr", True), ("Jahr (JJJJ)", True), ("Year", True), ("Nr.", True), ("Lfd. Nr.", True),
     ("Kunden-Nr.", True), ("Artikelnr", True), ("Rechnungsnummer", True), ("Order no.", True), ("Customer ID", True),
     ("Invoice number", True), ("Pos.", True), ("PLZ", True), ("Postleitzahl", True), ("ZIP code", True),
     ("Vorjahr", False), ("Jahresumsatz", False), ("Umsatz (Jahr)", False), ("Casino", False), ("Paid", False),
     ("Total number", False), ("Number", False), ("Anzahl", False), ("Average number", False),
     ("Number of employees", False), ("Durchschnittliche Nummer", False), ("Nummer", True), ("No.", True), ("#", True),
     ("Laufende Nummer", True), ("Pos.-Nr.", True), ("Employee no.", True), ("Cost center number", True)],
)  # fmt: skip
def test_columns_of_years_and_numbers(header: str, expected: bool) -> None:
    assert od.id_column(od.ExcelColumn(header=header, type="integer"), [2024, 2025]) is expected


def test_years_and_numbers_without_thousands_separators() -> None:
    columns = ["Region", "Jahr", {"header": "Kundennr.", "type": "integer"}, {"header": "PLZ", "type": "integer"},
               {"header": "Nr.", "type": "integer", "total": "sum"}, {"header": "Umsatz", "type": "number"},
               {"header": "Stück", "type": "integer"}]  # fmt: skip
    rows = [["Nord", 2024, 100234, 70173, 1001, 1500, 1200], ["Süd", 2025, 100235, 80331, 1002, 2500, 1300]]
    ws = excel([sheet(columns, rows)], valves={"language": "de"})[0]["S"]
    assert [ws.cell(2, c).number_format for c in range(2, 8)] == ["0", "0", "0", "#,##0", "#,##0", "#,##0"]
    sheets = od.parse_argument(od.EXCEL_SHEETS, [sheet(columns, rows)], "sheets")
    text = od.build_xlsx(sheets, "EUR", od.TRANSLATIONS["de"], decimal_comma=True)[1]
    assert "| Nord | 2024 | 100234 | 70173 | 1.001 | 1.500 | 1.200 |" in text
    years = sheet(["Datum", {"header": "Jahr", "formula": "=YEAR(A{row})"}], [["2026-01-05"]])
    assert excel([years])[0]["S"]["B2"].number_format == "0"
    amounts = sheet(["A", {"header": "Jahr", "type": "number"}], [["x", 1.5]])  # not whole numbers: amounts
    assert excel([amounts])[0]["S"]["B2"].number_format == "#,##0.0"
    counts = sheet(["Region", {"header": "Average number", "type": "integer"}], [["Nord", 12000], ["Süd", 15000]])
    ws = excel([{**counts, "chart": {"type": "bar", "title": "Average"}}])[0]["S"]  # an amount: plotted
    assert ws["B2"].number_format == "#,##0" and ws._charts[0].series[0].val.numRef.f == "'S'!$B$2:$B$3"


def test_document_properties() -> None:
    """Title, the user as author and the date of today, nothing of python-docx, python-pptx and openpyxl."""
    user = {"id": "u1", "name": "Erika Muster", "email": "erika@example.com", "role": "user"}
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": "Hallo"}], title="Bericht Q3", __user__=user)
    prs, _ = powerpoint([{"layout": "title", "title": "Vortrag"}], title="Vortrag", __user__=user)
    for core, title in ((doc.core_properties, "Bericht Q3"), (prs.core_properties, "Vortrag")):
        assert (core.title, core.author, core.last_modified_by, core.comments) == (title, *[user["name"]] * 2, "")
        assert core.created.date() == date.today() and core.revision == 1
    wb, _ = excel([sheet(["A"], [["x"]])], title="Liste", __user__=user)
    assert (wb.properties.title, wb.properties.creator, wb.properties.lastModifiedBy) == ("Liste", *[user["name"]] * 2)
    _, recorder = call("create_powerpoint_file", slides=[{"layout": "title", "title": "Vortrag"}])
    core = Presentation(io.BytesIO(recorder.download()[0])).core_properties
    assert (core.author, core.last_modified_by, core.comments) == ("", "", "")  # not "Steve Canny"
    _, recorder = call("create_excel_file", sheets=[sheet(["A"], [["x"]])])
    assert b"openpyxl" not in zipfile.ZipFile(io.BytesIO(recorder.download()[0])).read("docProps/core.xml")


def test_windows_paths_keep_their_backslashes() -> None:
    for text, expected in [
        (r"Ablage: \\fs01\Vertrieb\_Archiv\2026", r"Ablage: \\fs01\Vertrieb\_Archiv\2026"),
        (r"(G:\, H:\) und C:\Daten\*.xlsx, \*kein\* Kursiv", r"(G:\, H:\) und C:\Daten\*.xlsx, *kein* Kursiv"),
        (r"**C:\Daten** und „\\nas\x“", r"C:\Daten und „\\nas\x“"),
        # every backslash escaped as in Markdown: the escapes apply, as before
        (r"auf \\\\srv01\\backup\\<JJJJ-MM> oder C:\\Daten", r"auf \\srv01\backup\<JJJJ-MM> oder C:\Daten"),
    ]:
        assert od.to_plain(text) == expected
    ws = excel([sheet(["Pfad"], [[r"\\fs01\Vertrieb\_Archiv\2026"]])])[0]["S"]
    assert ws["A2"].value == r"\\fs01\Vertrieb\_Archiv\2026"


def test_links_in_excel_cells() -> None:
    rows = [
        ["[Intranet](https://intranet.acme.de/hr)", "siehe [Wiki](https://wiki.acme.de) oder [IT](mailto:it@acme.de)"],
        ["<https://acme.de>", "[Böse](javascript:alert(1))"],
    ]
    ws = excel([sheet(["Seite", "Notiz"], rows)])[0]["S"]
    assert (ws["A2"].value, ws["A2"].hyperlink.target, ws["A2"].font.u) == ("Intranet", rows[0][0][11:-1], "single")
    assert (ws["A3"].value, ws["A3"].hyperlink.target, ws["A3"].style) == ("https://acme.de",) * 2 + ("Hyperlink",)
    assert ws["B2"].value == "siehe Wiki (https://wiki.acme.de) oder IT (mailto:it@acme.de)"
    assert ws["B2"].hyperlink is None
    assert ws["B3"].value == "Böse" and ws["B3"].hyperlink is None  # only http, https and mailto
    sheets = od.parse_argument(od.EXCEL_SHEETS, [sheet(["Seite", "Notiz"], rows)], "sheets")
    assert "| [Intranet](https://intranet.acme.de/hr) |" in od.build_xlsx(sheets, "EUR", od.TRANSLATIONS["en"])[1]


def test_code_blocks_in_word_paragraphs() -> None:
    text = ("So startet man den Dienst:\n\n```bash\n# Dienst neu starten\nsudo systemctl restart nginx\nif true; then\n"
            "    echo **ok** | tee log\nfi\n```\n\nFertig.")  # fmt: skip
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": text}], title="Anleitung")
    paragraphs = [(p.style.name, p.text) for p in doc.paragraphs[1:]]
    code = "# Dienst neu starten\nsudo systemctl restart nginx\nif true; then\n    echo **ok** | tee log\nfi"
    assert paragraphs == [("Normal", "So startet man den Dienst:"), ("Normal", code), ("Normal", "Fertig.")]
    assert doc.paragraphs[2].runs[0].font.name == od.CODE_FONT
    assert doc.paragraphs[2]._p.pPr.find(f"{W_NS}shd") is not None
    # without its closing line the block ends with the text; ```inline``` is no block
    unclosed = od.split_paragraph("Text\n```\n| a | b |\n|---|---|")
    assert unclosed == [("text", ["Text"]), ("code", ["| a | b |", "|---|---|"])]
    assert od.split_paragraph("```x``` bleibt Text") == [("text", ["```x``` bleibt Text"])]
    # a fence line holds at most one word (the language): the text after ~~~durchgestrichen~~~ stays
    assert od.split_paragraph("~~~durchgestrichen~~~ Text\nmehr") == [("text", ["~~~durchgestrichen~~~ Text", "mehr"])]
    assert od.split_paragraph("```Achtung wichtig\nmehr") == [("text", ["```Achtung wichtig", "mehr"])]
    for opening, closing in (("```python", "```"), ("~~~", "~~~"), ("~~~ c++", "~~~~"), ("````shell-session", "````")):
        assert od.split_paragraph(f"Text\n{opening}\nx = 1\n{closing}\nmehr") == [
            ("text", ["Text"]), ("code", ["x = 1"]), ("text", ["mehr"])
        ]


def test_iso_timestamps() -> None:
    rows = [["a", "2026-10-23T08:30:00Z"], ["b", "2026-10-23T09:00:00+02:00"], ["c", "2026-10-23 08:30:15.250"]]
    ws = excel([sheet(["A", {"header": "Zeit", "type": "date"}], rows)])[0]["S"]
    expected = [datetime(2026, 10, 23, 8, 30), datetime(2026, 10, 23, 9), datetime(2026, 10, 23, 8, 30, 15, 250000)]
    for r, value in enumerate(expected, start=2):  # the time zone is dropped, the time stays as written
        assert abs(ws.cell(r, 2).value - value) < timedelta(milliseconds=1)
    assert ws["B2"].number_format == "YYYY-MM-DD hh:mm"
    ws = excel([sheet(["A", "Zeit"], rows)])[0]["S"]  # a column without a type becomes a date column
    assert isinstance(ws["B2"].value, datetime)
    hours = sheet(["A", {"header": "Zeit", "type": "date"}], [["a", "2026-10-23 8:30 Uhr"]])
    message = "is not a valid date: write dates as YYYY-MM-DD, e.g. 2026-11-02, with a time as YYYY-MM-DD hh:mm"
    assert message in invalid("create_excel_file", sheets=[hours])


def test_booleans_in_tables() -> None:
    ws = excel([sheet(["Aufgabe", "Erledigt"], [["Angebot", True], ["Rechnung", False]])])[0]["S"]
    assert (ws["B2"].value, ws["B3"].value, ws["B2"].data_type, ws["B2"].number_format) == (True, False, "b", "General")
    counted = sheet(["A", {"header": "Menge", "type": "integer", "total": "sum"}], [["x", 3], ["y", True]])
    message = "'True' in 'Menge' is not a number, so the total cannot count it"
    assert message in invalid("create_excel_file", sheets=[counted])
    sheets = od.parse_argument(od.EXCEL_SHEETS, [sheet(["A", "B"], [["x", True]])], "sheets")
    assert "| x | WAHR |" in od.build_xlsx(sheets, "EUR", od.TRANSLATIONS["de"], decimal_comma=True)[1]
    rows = [["Punkt", "Erledigt"], ["Angebot", True], ["Rechnung", False]]
    blocks = [{"type": "table", "rows": rows}]
    _, recorder = call("create_word_file", valves={"language": "de"}, blocks=blocks, title="T")
    table = Document(io.BytesIO(recorder.download()[0])).tables[0]
    assert [row.cells[1].text for row in table.rows] == ["Erledigt", "Ja", "Nein"]
    prs, _ = powerpoint([{"layout": "table", "title": "T", "table": rows}], title="T")
    (shape,) = [s for s in prs.slides[0].shapes if s.has_table]
    assert [shape.table.cell(r, 1).text for r in range(3)] == ["Erledigt", "Yes", "No"]


def test_dates_before_1900_stay_text() -> None:
    """Excel has no dates before 1900; 1900-01-01 to 1900-02-28 are a day apart in Excel and LibreOffice."""
    rows = [["Reichsgründung", "1871-01-18"], ["x", "1900-02-28"], ["y", "1900-03-01"]]
    for language, early in (("de", ["18.01.1871", "28.02.1900"]), ("en", ["1871-01-18", "1900-02-28"])):
        ws = excel([sheet(["A", {"header": "Datum", "type": "date"}], rows)], valves={"language": language})[0]["S"]
        assert [ws["B2"].value, ws["B3"].value, ws["B4"].value] == [*early, datetime(1900, 3, 1)]
        assert ws["B2"].data_type == "s" and ws["B2"].alignment.horizontal == "center"
    # the checks see the text too: computing with it gives #VALUE! in Excel
    deadline = {"header": "Frist", "type": "date", "formula": "=B{row}+14"}
    later = sheet(["A", {"header": "Datum", "type": "date"}, deadline], rows[:1])
    assert (
        "the formula computes with 'Datum' (B{row}), but sheets[0].rows[0] has '1871-01-18' there: Excel has no dates "
        "before 1900-03-01"
    ) in invalid("create_excel_file", sheets=[later])
    in_text = sheet(["A", "B", "C"], [["x", "18.01.1871", "=B2+1"]])  # a text column: the same text
    assert "Excel has no dates before 1900-03-01" in invalid("create_excel_file", sheets=[in_text])
    ws = excel([sheet(["A", {"header": "Datum", "type": "date"}, deadline], rows[2:])])[0]["S"]
    assert ws["C2"].value == "=B2+14"


def test_word_header_text_is_plain() -> None:
    doc, _, _ = word(blocks=[{"type": "paragraph", "text": "x"}], title="T", header_text="**ACME** GmbH · *intern*")
    assert doc.sections[0].header.paragraphs[0].text == "ACME GmbH · intern"


def test_type_errors_of_a_cell_in_one_line() -> None:
    nested = sheet(["A", "B"], [["x", 1], [["y", 2]]])
    result = invalid("create_excel_file", sheets=[nested])
    assert "- sheets[0].rows[1][0]: Input should be a valid string, number or boolean\n" in result


@pytest.mark.skipif(not shutil.which("soffice"), reason="needs LibreOffice")
def test_new_formulas_compute_in_libreoffice() -> None:
    """The functions written with _xlfn., times and whole columns give values, not #NAME? or Err:5xx."""
    work = sheet([{"header": "Name"}, {"header": "Beginn", "type": "time"}, {"header": "Ende", "type": "time"},
                  {"header": "Stunden", "type": "number", "formula": "=(C{row}-B{row})*24"},
                  {"header": "Stufe", "formula": '=IFS(D{row}>8,"lang",TRUE,"kurz")'},
                  {"header": "Max", "formula": "=MAXIFS(D:D,A:A,A{row})"},
                  {"header": "Text", "formula": '=TEXTJOIN("-",TRUE,A{row},E{row})'},
                  {"header": "Suche", "formula": "=XLOOKUP(A{row},A:A,D:D)"}],
                 [["Ana", "08:00", "16:30"], ["Ben", "08:00", "16:00"]])  # fmt: skip
    with tempfile.TemporaryDirectory() as tmp:
        excel([work])[0].save(Path(tmp) / "work.xlsx")
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "xlsx",
             "--outdir", f"{tmp}/out", "work.xlsx"],
            cwd=tmp, check=True, capture_output=True, timeout=180,
        )  # fmt: skip
        ws = load_workbook(Path(tmp) / "out" / "work.xlsx", data_only=True).active  # with the computed values
        assert [c.value for c in ws[2]][3:] == [8.5, "lang", 8.5, "Ana-lang", 8.5]
        assert [c.value for c in ws[3]][3:] == [8, "kurz", 8, "Ben-kurz", 8]


def excel_data(sheets: list[dict[str, Any]], **kwargs: Any) -> bytes:
    result, recorder = call("create_excel_file", sheets=sheets, **kwargs)
    assert result.startswith("Done"), result
    return recorder.download()[0]


VALUES_BUDGET = sheet([{"header": "Position"}, {"header": "Plan", "type": "currency", "total": "sum"},
                       {"header": "Ist", "type": "currency", "total": "sum"},
                       {"header": "Abweichung", "type": "currency", "formula": "=C{row}-B{row}", "total": "sum"},
                       {"header": "Status", "formula": '=IF(C{row}>B{row},"über","ok")'},
                       {"header": "MwSt", "formula": "=ROUND(C{row}*0.19,2)"},
                       {"header": "Prüfen", "formula": "=AND(C{row}>0,NOT(D{row}>400))"}],
                      [["Miete", 1500, 1480], ["Werbung", 3000, 3450.5], ["Sonstiges", 500, None]])  # fmt: skip
VALUES_BILL = [
    {"name": "Sätze", "columns": [{"header": "Art"}, {"header": "Satz", "type": "percent"}], "rows": [["MwSt", "19%"]]},
    {"name": "Rechnung", "columns": [{"header": "Pos"}, {"header": "Netto", "type": "currency"},
                                     {"header": "Saldo", "type": "currency"}, {"header": "Frist", "type": "date"}],
     "rows": [["A", 100, "=B2*(1+'Sätze'!$B$2)", "2026-10-01"], ["B", 50, "=C2+B3", "=D2+14"],
              ["C", 2.5, "=C3+B4", "=MAX(D2:D3)"]]},
]  # fmt: skip


def test_formula_values_are_stored() -> None:
    """openpyxl writes formulas without their values; viewers that do not compute (Open WebUI's preview) show the
    values stored next to them, Excel computes them again."""
    ws = load_workbook(io.BytesIO(excel_data([VALUES_BUDGET])), data_only=True).active
    assert [c.value for c in ws[2]][3:] == [-20, "ok", 281.2, True]
    assert [c.value for c in ws[3]][3:] == [450.5, "über", 655.6, False]  # ROUND of 655.595, not of 655.59499...
    assert [c.value for c in ws[4]][3:] == [-500, "ok", 0, False]  # an empty cell is 0
    assert [c.value for c in ws[5]][:4] == ["Total", 5000, 4930.5, -69.5]
    workbook = load_workbook(io.BytesIO(excel_data(VALUES_BILL)), data_only=True)
    ws = workbook["Rechnung"]
    assert [ws[f"C{r}"].value for r in (2, 3, 4)] == pytest.approx([119, 169, 171.5])
    assert ws["D3"].value == ws["D4"].value == datetime(2026, 10, 15)
    formulas = load_workbook(io.BytesIO(excel_data([VALUES_BUDGET])))
    assert formulas.active["D2"].value == "=C2-B2"  # the formulas stay


def test_formula_values_left_to_excel() -> None:
    """Formulas whose value is not certain get none: other functions and the formulas that use them, errors, what is
    left of 0.3 - 0.1 - 0.2, booleans in COUNT (LibreOffice counts them), texts by order, numbers joined to text with
    a decimal mark."""
    work = sheet([{"header": "Name"}, {"header": "X", "type": "number"}, {"header": "Y", "type": "number"},
                  {"header": "Flag"}, {"header": "Suche", "formula": "=VLOOKUP(A{row},$A$2:$B$3,2,FALSE)"},
                  {"header": "Doppelt", "formula": "=E{row}*2"}, {"header": "Quote", "formula": "=B{row}/C{row}"},
                  {"header": "Rest", "formula": "=B{row}-C{row}-0.2"}, {"header": "Zahl", "formula": "=COUNT(B$2:D$3)"},
                  {"header": "Ordnung", "formula": '=A{row}<"m"'}, {"header": "Text", "formula": '=B{row}&" Stück"'}],
                 [["a", 1, 0, True], ["b", 0.3, 0.1, False]])  # fmt: skip
    ws = load_workbook(io.BytesIO(excel_data([work])), data_only=True).active
    assert [c.value for c in ws[2]][4:] == [None, None, None, 0.8, None, None, "1 Stück"]
    assert [c.value for c in ws[3]][4:] == [None, None, pytest.approx(3), None, None, None, None]


def test_formula_values_stop_after_many_range_cells(monkeypatch: pytest.MonkeyPatch) -> None:
    """A running total reads the rows above it again in every row: after MAX_RANGE_CELLS the rest is left to Excel."""
    monkeypatch.setattr(od, "MAX_RANGE_CELLS", 10)
    work = sheet([{"header": "X", "type": "number"}, {"header": "Summe", "formula": "=SUM($A$2:A{row})"}],
                 [[1], [2], [3], [4], [5]])  # fmt: skip
    ws = load_workbook(io.BytesIO(excel_data([work])), data_only=True).active
    assert [ws[f"B{r}"].value for r in range(2, 7)] == [1, 3, 6, 10, None]

    monkeypatch.setattr(od, "MAX_RANGE_CELLS", 3)  # one function at other places of the formulas: read once
    shared = sheet([{"header": "X", "type": "number"}, {"header": "Anteil", "formula": "=A{row}/SUM($A$2:$A$4)"},
                    {"header": "Rest", "formula": "=SUM($A$2:$A$4)-A{row}"}], [[1], [2], [5]])  # fmt: skip
    ws = load_workbook(io.BytesIO(excel_data([shared])), data_only=True).active
    assert [ws[f"B{r}"].value for r in (2, 3, 4)] == [0.125, 0.25, 0.625] and ws["C4"].value == 3


def test_formula_values_never_break_the_file(monkeypatch: pytest.MonkeyPatch) -> None:
    """An overflow gets no value (<v>inf</v> is no number), an error in FormulaValues none at all."""
    huge = sheet([{"header": "X", "type": "number", "total": "sum"}], [[1e308], [1e308]])
    data = excel_data([huge])
    assert load_workbook(io.BytesIO(data), data_only=True).active["A4"].value is None
    assert b"inf" not in zipfile.ZipFile(io.BytesIO(data)).read("xl/worksheets/sheet1.xml")

    def broken(self: Any) -> None:
        raise TypeError("unexpected")

    monkeypatch.setattr(od.FormulaValues, "all", broken)
    assert load_workbook(io.BytesIO(excel_data([VALUES_BUDGET])), data_only=True).active["D2"].value is None


@pytest.mark.skipif(not shutil.which("soffice"), reason="needs LibreOffice")
def test_formula_values_match_libreoffice() -> None:
    """The stored values are the ones LibreOffice computes: it shows stored values as they are."""
    numbers = sheet([{"header": "A", "type": "number"}, {"header": "B", "type": "number"},
                     {"header": "Rund", "formula": "=ROUND(A{row}*B{row},2)"},
                     {"header": "Pot", "formula": "=A{row}^B{row}"}, {"header": "Vgl", "formula": "=A{row}>=B{row}"},
                     {"header": "Neg", "formula": "=-A{row}%"}, {"header": "Abs", "formula": "=ABS(A{row}-B{row})"},
                     {"header": "Wenn", "formula": "=IF(A{row}>2,A{row})"},
                     {"header": "Stat", "formula": "=AVERAGE($A$2:$B$6)+MIN($A$2:$A$6)-COUNT($A$2:$B$6)"}],
                    [[2.675, 1], [1.005, 1], [0.1, 0.2], [-2.5, 2], [None, 3]], name="Zahlen")  # fmt: skip
    times = sheet([{"header": "Start", "type": "date"}, {"header": "Von", "type": "time"},
                   {"header": "Bis", "type": "time"}, {"header": "Frist", "type": "date", "formula": "=A{row}+14"},
                   {"header": "Std", "formula": "=(C{row}-B{row})*24"}, {"header": "Zeit", "type": "time", "total": "sum"}],
                  [["2026-10-01", "08:00", "16:30", None, None, "01:30"],
                   ["2026-10-24 14:30", "22:00", "06:00", None, None, "23:45"]], name="Zeiten")  # fmt: skip
    data = excel_data([VALUES_BUDGET, numbers, times, *VALUES_BILL])
    stored, formulas = load_workbook(io.BytesIO(data), data_only=True), load_workbook(io.BytesIO(data))
    with tempfile.TemporaryDirectory() as tmp:
        formulas.save(Path(tmp) / "work.xlsx")  # without the values: LibreOffice computes them
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "xlsx",
             "--outdir", f"{tmp}/out", "work.xlsx"],
            cwd=tmp, check=True, capture_output=True, timeout=180,
        )  # fmt: skip
        computed = load_workbook(Path(tmp) / "out" / "work.xlsx", data_only=True)
    compared = 0
    for ws in stored.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and formulas[ws.title][cell.coordinate].data_type == "f":
                    expected = computed[ws.title][cell.coordinate].value
                    expected = pytest.approx(expected) if isinstance(expected, float) else expected
                    assert cell.value == expected, (ws.title, cell.coordinate)
                    compared += 1
    assert compared >= 50


# ----------------------------------------------------------------------------------------- rendering --


@pytest.mark.skipif(not (shutil.which("soffice") and shutil.which("pdftotext")), reason="needs LibreOffice and poppler")
def test_files_render_in_libreoffice() -> None:
    files = {
        "doc.docx": word(blocks=BLOCKS, title="Render test")[0],
        "book.xlsx": excel([BUDGET])[0],
        "deck.pptx": powerpoint(SLIDES)[0],
    }
    with tempfile.TemporaryDirectory() as tmp:
        for name, document in files.items():
            document.save(Path(tmp) / name)
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "pdf",
             "--outdir", tmp, *files],
            cwd=tmp, check=True, capture_output=True, timeout=180,
        )  # fmt: skip
        texts = {
            name: subprocess.run(
                ["pdftotext", str(Path(tmp) / f"{Path(name).stem}.pdf"), "-"],
                capture_output=True,
                text=True,
                check=True,
            ).stdout
            for name in files
        }
    assert "Render test" in texts["doc.docx"] and "Page 1 of 2" in texts["doc.docx"]
    assert "Rent" in texts["book.xlsx"] and "Total" in texts["book.xlsx"]
    assert "Quarterly report" in texts["deck.pptx"] and "Pros and cons" in texts["deck.pptx"]
