# Office Documents

An Open WebUI **tool** that lets the model create Office files itself: "Write the minutes as a Word file", "Make
me a budget in Excel", "Turn this into a presentation". The model does not write any code. It fills in
a fixed JSON schema, the tool checks it and builds the file. The finished file is attached to the
answer, where the user can preview and download it.

| Tool function | What the model can use |
|---|---|
| `create_word_file` | Title and subtitle, table of contents (at the top or where the model places it), headings (3 levels), paragraphs, bullet and numbered lists (2 levels, numbering restarts for each list), tables with a coloured header row that repeats on every page and column widths fitted to the content, quotes, page breaks, text in the page header, "Page X of Y" footer, landscape pages |
| `create_excel_file` | Several sheets with typed columns (`text`, `number`, `integer`, `currency`, `percent`, `date`), formulas per column, a total row (`sum`, `average`, `min`, `max`, `count`) and native Excel charts (column, bar, line, pie). Header row coloured and frozen, with an auto filter and fitted column widths; printed on one page width |
| `create_powerpoint_file` | 16:9 slides with the layouts `title`, `section`, `bullets`, `two_columns`, `table` and `chart` (native PowerPoint charts: column, bar, line, pie, optionally with the unit at the value axis), plus speaker notes |

Inline `**bold**`, `*italic*` and `[links](https://…)` work in all text fields of Word and PowerPoint;
in Excel cells the markers are removed. File names come from the title, e.g. `Budget 2026.xlsx`.

![An answer with the attached presentation](images/answer.png)

![The presentation in Open WebUI's file preview](images/preview.png)

![Files made by the tool: a Word document, an Excel sheet with a chart and slides](images/files.png)

## Setup

File: [`office_documents.py`](office_documents.py).

1. Install it under **Workspace → Tools** (not under Functions), see [Installation](../README.md#installation).
2. Enable it for a model: **Workspace → Models → (model) → Edit**, tick **Office Documents** under
   *Tools*. Users can also switch it on per chat in the **Integrations** menu of the chat input.
3. In the same model settings, set **Advanced Params → Function Calling** to **Native**. The tool
   needs a model and server with tool calling (llama-server with `--jinja`, vLLM with
   `--enable-auto-tool-choice` and the model's `--tool-call-parser`).
4. Set **max_tokens** there as well (e.g. `8192`), so a model that gets stuck in a loop while
   writing the arguments stops after a while.

## What the model sends

A short example for a slide deck (the model sees the full schema with descriptions):

```json
{
  "title": "Quarterly report Q3",
  "slides": [
    {"layout": "title", "title": "Quarterly report Q3", "subtitle": "Sales"},
    {"layout": "bullets", "title": "Results", "bullets": ["Revenue **+12 %**", "- mostly online"], "notes": "…"},
    {"layout": "chart", "title": "Revenue", "chart": {"type": "column", "categories": ["Q1", "Q2", "Q3"],
      "series": [{"name": "2026", "values": [11, 13, 14]}]}}
  ]
}
```

- **Excel values:** numbers as JSON numbers, percentages as `12.5` for 12.5 %, dates as `YYYY-MM-DD`.
  Text such as `1.234,56 €` or `7,5 %` is converted too. A separator followed by exactly three digits
  is read as thousands (`1.500`), except after a leading zero (`0,125`) and for the decimal mark of the
  `language` valve when it appears only once (`de`: `1,799` is 1.799; `en`: `1.500` is 1.5). Columns sent
  without a type whose values are all numbers become number columns, so sums and charts work; values
  with leading zeros (postal codes, article numbers) stay text.
- **Percent cells hold the fraction**, as in Excel: `12.5` is stored as 0.125 and shown as 12.5 %, so a
  formula uses the cell as it is (`=B{row}*(1-C{row})` for a discount in C). The tool sends back three
  slips that models make with this (35 of 660 Excel calls in the v0.1 runs of
  [office-bench-de](https://github.com/tovren/office-bench-de)): fractions sent into a percent
  column (`0.05` for 5 %, which would show as 0.05 %; a value below 1 % can be sent as `"0.5 %"`), a
  percent cell divided by 100 in a formula, and `*100` in the formula of a percent column (25 % would
  show as 2500 %).
- **Excel formulas** belong to a column and are filled into every data row, e.g. `=B{row}*C{row}`
  (`{row}` is the current row). A formula written for the first data row (`=B2*C2`) gets `{row}`
  automatically. Allowed are cell references of the same workbook, numbers, quoted text,
  operators and these functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `COUNT`, `COUNTA`,
  `COUNTIF`, `SUMIF`, `AVERAGEIF`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `ABS`, `PRODUCT`, `POWER`, `SQRT`, `MOD`, `INT`. German names (`SUMME`, `WENN`,
  `RUNDEN` …), `;` as separator and decimal commas (`=RUNDEN(B{row}*0,19;2)`) are translated. The
  total row's `count` counts all filled cells (`COUNTA`). Formulas may also be written into single cells,
  e.g. a balance row `=B2-B3`, `=C2-C3` … when the months are columns; they are checked against the same
  whitelist. In columns the model marks as `text`, values that start with `=` stay text.
- **Cells of other sheets** are written as `'Rates'!$B$3` (also `Rates!B3`, without regard to case). In a
  column formula the model has to say which cell it means: `$B$3` for one cell in every row (e.g. an
  hourly rate) or `B{row}` for the same row; `B3` alone is sent back. A single cell of another sheet that
  is empty, usually because the rows there were counted wrong (row 1 is the header), is sent back with the
  rows and columns that sheet has, as it would silently count as 0. Ranges such as `'Times'!$A$2:$A$50`
  may reach beyond the data.
- **Excel charts** plot columns (`category_column` plus `value_columns`) or, for tables with e.g. the
  months as columns, rows (`series_rows`: the labels in the first column of the rows to plot). A chart
  with only a type and title plots the number columns against the first column.
- **Errors go back to the model.** If something is missing or too long, the tool answers with a list
  of exact problems (e.g. `slides[3].chart.series[0]: 2 values for 3 categories`), and the model
  can call it again. The positions refer to the list the model sent.
- **Content that cannot be meant is rejected**, so a technically valid but useless file is not reported
  as done: a number column that holds mostly text (usually rows shifted against the columns), an empty
  column that a formula uses (e.g. prices left out of an offer), a total on an empty column, a total row
  written into the rows (a last row labelled e.g. "Summe" or "Netto-Summe": it would get the column
  formulas too, be counted twice next to a column's `total`, or add up nothing, so the model is told to use
  `total`), a chart column without numbers,
  a formula that uses its own column, a slide table with only a header row, a table of contents without
  chapters, and more text than fits on a slide or into a slide table.
- **Content is never dropped without a word.** A field the tool does not know is sent back when it holds
  text or a list of values (`"content"` instead of `"text"`, `"formulas"`, a sheet-level `"total"`, a
  table under `"data"`), and so is content in a field that the block or layout does not show (a paragraph's
  `items`, text on a page break, bullets on a section slide). Unknown fields with a single number or an
  empty value, and a chart's own `title` on a slide (the slide title is shown), are ignored.
- **Other shapes are not guessed.** A block without `type`, a slide without `title` or a list sent as one
  text gets an error that names the field, and the model sends it again in the right shape. Only
  unambiguous things are accepted as they are: `null` for unused fields, `{"slides": [...]}` instead of
  the list, list items that all start with `- ` (top-level items, not sub items), `<br>` as a line break,
  Markdown lists and tables inside a paragraph (numbered lines only become a list when they count up from
  1, so "3. Oktober …" stays text), numbers sent as text in an Excel column without a type, `"null"` as
  text in a number or date column (an empty cell), a chart object without data (no chart), formulas
  without `=`, formulas written into single cells, file-name-style titles (`Report_Q3.xlsx`), a missing title (the
  first heading or slide title is used), and empty paragraphs used as spacing (left out). A heading that
  only repeats the title is left out.
- **Broken JSON is repaired.** Long arguments often reach the tool as JSON text rather than as
  structured data (the server could not parse them), with the same few mistakes: German quotes
  closed with a straight quote (`„Hoch"`), a forgotten `}` or `]`, table rows written as several
  arrays without the outer brackets, or the other tool parameters (e.g. the title) written into the
  same text, which are then used. Fields that end up inside a chart object because its `}` was forgotten
  (e.g. the speaker notes) are taken back to the slide or sheet. Valid JSON is never changed, and a repair
  is only used if no content is lost. Otherwise the error shows the text on both sides of the place and
  names the likely slip: a quote that ends a text too early, a bracket too many, a key inside a list that
  was not closed, a missing comma or value, a comma before `]`, a backslash or line break inside a text,
  or arguments cut off in the middle (output limit reached).
- **Tested on new requests:** with five requests that were not used while writing the tool, Qwen3.8 27B
  needed the JSON repair for 8 of 40 calls (with a little thinking) and would have needed the shape
  repairs of earlier versions for 1 of 60. Gemma 4 26B sends shapes outside the schema much more often
  (about 4 in 10 calls of the earlier tests); the tool sends them back, which costs extra rounds.

## Valves

| Valve | Default | Description |
|---|---|---|
| `language` | `en` | Language of the status messages and of texts inside the files: total row label, page footer, date format, table of contents heading (`en` or `de`); also the decimal mark for numbers sent as text in Excel |
| `delivery` | `attach` | `attach`: store the file in Open WebUI and attach it to the answer (preview, stays in the chat). `download`: hand it straight to the browser as a download (nothing is stored, but it is gone after a reload). If attaching fails, the tool falls back to a download |
| `currency` | `EUR` | Currency of `currency` columns in Excel (`EUR`, `USD`, `GBP`, `CHF`) |
| `accent_color` | `1F4E79` | Accent colour of all three formats (titles, table header rows, charts, title slide), as a hex value without `#`. Lighter and darker tones are derived from it |
| `font` | `Calibri` | Font of all three formats, e.g. a company font such as `Arial` or `Aptos`. It has to be installed where the files are opened; otherwise Office shows a similar one. Excel columns get a little more room for fonts wider than Calibri |
| `page_numbers` | `true` | Add "Page X of Y" to the footer of Word documents |
| `word_template` | _(empty)_ | Path of a Word file (`.docx` or `.dotx`) on the Open WebUI server, e.g. a letterhead: Word documents get its header, footer, page size and margins (see [Word template](#word-template)) |

## Word template

With the `word_template` valve, Word documents are built on a company file, e.g. a letterhead. Copy the
`.docx` or `.dotx` to the server where Open WebUI runs, e.g. into its data volume (Docker:
`/app/backend/data/templates/letterhead.dotx`), and enter the path.

- **Taken from the template:** header and footer with their logos, fields and fonts (also a different
  first page, e.g. a large letterhead on page 1 and a small one after it), page size and margins.
- **Not taken:** the template's text is removed, also what is not in its header or footer (e.g. an address
  window or a logo placed in the body). The content keeps the tool's design (`accent_color`, `font`): it
  uses the styles *Body Text*, *Title*, *Heading 1–3*, *Quote* and the list styles, which are replaced in
  the document. The template's *Normal* style stays as it is, since its header and footer build on it, so
  set `font` to the company font if the two should match.
- If the template has a header, the model's `header_text` is not added to it; if it has a footer, the
  "Page X of Y" footer is not added either. Without them, both are added as usual.
- Removed from the template: its author and title, a mail merge (Word would ask to run its SQL command),
  write protection, and the path of the template it was made from.
- If the file cannot be read (missing, not a Word file, larger than 20 MB), the documents are made without
  it and Open WebUI's log says why.

## Security & privacy

- **No code from the model.** It only supplies data, which is validated against the schema before
  fixed code builds the file. There is no shell and no network access, and nothing is loaded from
  URLs (images are not supported for that reason).
- **Excel:** every formula, in a column's `formula` field or in a cell, is checked against the whitelist
  above: references to other workbooks (`[…]`, or a sheet name that is not in the file), DDE (`|`) and every
  other function, e.g. `WEBSERVICE`, `HYPERLINK` or `INDIRECT`, are rejected. A cell value that does not pass,
  such as `=HYPERLINK(...)`, is written as text with Excel's quote prefix, so it stays text when the user
  edits it (in a number column the model is told instead). The data sheet behind PowerPoint charts
  ("Edit Data") gets labels and series names as text only, never as formulas or links.
- **Links** are only created for `http`, `https` and `mailto` (only with `subject`, `body`, `cc`,
  `bcc`); other link targets stay plain text. Invisible format characters such as a right-to-left
  override are removed from file names.
- **Limits** keep a runaway model from blocking the server: 200,000 characters per file, 500 Word
  blocks, 20 sheets with 10,000 rows and 50 columns each and 250,000 cells in all, 60 slides with 12
  items and 16 table rows each, 500 characters per formula. The Markdown parser runs in linear time,
  so crafted text (e.g. from a prompt injection in an uploaded document) cannot stall the server.
- **Storage:** with `delivery = attach` the file is uploaded to Open WebUI's file storage under the
  user's account and linked to the chat, so only the user (and admins) can open it. It is not added
  to any knowledge base, and Open WebUI's own text extraction and embedding are skipped; instead the tool
  stores a Markdown version of the file (headings, lists, tables, slide texts and chart data), which the
  "Content" tab of Open WebUI's file view shows. With `delivery = download` nothing is stored; like
  Response Export, this needs a CSP that allows `'unsafe-eval'`. The download only counts as done when
  the browser confirms it (at most 60 seconds); otherwise the model is told that the file could not be
  created.
- Like every Open WebUI plugin, the tool runs inside the Open WebUI server process.

## Notes and limitations

- **Previews in Open WebUI** use browser libraries that know only part of the formats. Word files are
  written so that the preview shows their fonts, colours and bullets, too. The Excel preview shows formula
  cells empty (formulas have no stored results; Excel and LibreOffice calculate them when the file is
  opened) and dates as numbers, and the slide preview leaves out some formatting such as bullets and the
  alignment of titles. In Excel, PowerPoint and LibreOffice the files look as intended.
- **Table of contents:** Word fills it in when the document is opened and asks first whether to
  update the fields. LibreOffice needs **Tools → Update → Update All**.
- **Design:** one calm, modern look for all three formats, in shades of the `accent_color` valve and in the
  font of the `font` valve (also for charts and for text the user adds later). Word documents can take the
  header and footer of a company template ([Word template](#word-template)); Excel and PowerPoint
  templates are not supported.
  - **Word:** A4 with `language = de` (Letter otherwise), a title block with a line in the accent colour,
    headings in the accent colour, tables with a coloured header row, light banding and lines between the
    rows only, quotes with a coloured bar, the title and "Page X of Y" in the footer.
  - **Excel:** a coloured header row with white text, light banding (a conditional format, so the stripes
    stay in place when the table is sorted), thin row lines and no grid lines, text, numbers and dates
    aligned with a small indent, a highlighted total row, charts in matching colours.
    Printed on one page width (landscape for wide sheets and sheets with a chart) with the header row on
    every page and the title in the page header.
  - **PowerPoint:** a dark title slide, section slides in the accent colour, content slides with the
    presentation title and slide number in the footer, coloured bullets, two columns as light cards, tables
    like in Word, charts in matching colours with values at the bars when there are few. Text is not shrunk
    to fit; the font size is chosen from the amount of text (24 down to 12 pt), slides take at most 12
    items, and if the text does not fit even at 12 pt the model is asked to split the slide.
- **The model makes the difference.** Five German test requests (minutes, budget, presentation, offer,
  concept) with Open WebUI's tool-call handling; the candidates ran 40 times each, and two files per model
  and request were reviewed blind (language, content and arithmetic, completeness, structure, usability;
  1 to 5). With the current tool almost every model ends up with a file; the difference is the content:
  - **Qwen3.8 27B with a little thinking 4.1** (the best candidate to run yourself; FP8 fits a 48 GB GPU),
    Qwen3.8 Flash 3.8, Qwen3.8 27B without thinking 3.8. Now and then a Chinese word in German text.
  - **gpt-oss-120b 3.3:** the fastest and technically most reliable, but thin texts. **Gemma 4 31B 3.2.**
  - **Gemma 4 26B-A4B 3.0 via API, 2.6 locally (Q4 on llama-server):** the weakest choice for this tool. After
    a garbled tool call, llama-server rejects the next request (HTTP 500), so the chat ends with an error.
  - **Claude Sonnet 5.5 as a reference 4.5:** the tool is not the bottleneck; the remaining mistakes come from
    the models. Its reviewers were Claude models themselves, so this number may be a little flattering.
  - In an earlier, smaller round (one file per request): Nemotron 3 Super 3.6, Mistral Small 4 3.5,
    Nemotron 3.5 Lightning 3.2, Ministral 3 14B 2.8, gpt-oss-20b 2.2, GLM-4.7-Flash 2.0.
  - Common to all: wrong weekdays and calendar weeks, placeholders such as "System X", miscalculated
    licence costs, missing totals. Every generated file needs a look before it is used.
- **Gemma 4's tool-call syntax** sometimes breaks inside long arguments: the token `<|"|>` or a key such as
  `",text:` ends up inside a text and the rest is lost. This was seen with the 26B model on llama-server and
  once with the 31B model via an API. The tool rejects such calls, so the model sends them again instead of
  a file with missing text.
- **Uses Open WebUI internals.** Attaching the file relies on the same internal functions as Open
  WebUI's built-in image generation (0.11.4). If they change in a later version, the tool falls back
  to a browser download.
