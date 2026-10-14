# Open WebUI plugins: Response Export, Context Guard, Company Context, Reasoning Toggle, Office Documents

Five self-contained plugins for [Open WebUI](https://github.com/open-webui/open-webui):

| Plugin | Type | File | What it does |
|---|---|---|---|
| [Response Export](#response-export) | Action | `response_export.py` | Adds buttons to export an answer as Word or Excel, print or save it as PDF, or open it as an e-mail |
| [Context Guard](#context-guard) | Filter | `context_guard.py` | Warns before the context window overflows and stops the request with a clear message |
| [Company Context](#company-context) | Filter | `company_context.py` | Adds company name, date, language and form of address to every request |
| [Reasoning Toggle](#reasoning-toggle) | Filter (toggle) | `reasoning_toggle.py` | Adds a button to the chat input that lets local reasoning models (vLLM, llama-server) think before they answer |
| [Office Documents](#office-documents) | Tool | `office_documents.py` | Lets the model create Word, Excel and PowerPoint files and attaches them to its answer |

Response Export, Context Guard, Company Context and Office Documents are in English by default and
include a German translation (`language` valve); Reasoning Toggle shows no texts of its own. All five
use only the Python standard library and libraries that already ship with Open WebUI, so nothing has
to be installed.

---

## Response Export

Adds four buttons below every assistant message:

| Button | Result |
|---|---|
| **Word (.docx)** | Word document with headings, bullet and numbered lists (one nesting level, numbering restarts for each list), tables with a bold header row, bold/italic/strikethrough/inline code, links shown as `text (url)`, code blocks and quotes |
| **Excel (.xlsx)** | One worksheet per Markdown table. The header row is bold, shaded and frozen, with an auto filter and fitted column widths. Numbers become real numbers, including German formats (`1.234,56`, `12,5 %`, `1.234,56 €`), `$`/`£` amounts and the codes `EUR`, `USD` and `GBP` |
| **PDF** | Opens the browser's print dialog with a clean A4 layout and a small footer. Choose "Save as PDF". The suggested file name is the first heading plus a timestamp |
| **E-Mail** | Opens the default mail program through a `mailto:` link. A `Subject:`/`Betreff:` line becomes the subject, and any preamble before it ("Here is a draft:") is dropped. If the answer is too long for a `mailto:` link, an `.eml` draft with plain-text and HTML parts is downloaded instead. Outlook opens it as a new, editable message |

Reasoning and tool-call blocks (`<details>…</details>`, `<think>…</think>`) are removed before
exporting. File names use the first heading of the answer plus a timestamp, e.g.
`Quarterly report 2026-10-12 1405.docx`.

> Screenshot: _\<placeholder: the four buttons below a message\>_
>
> Screenshot: _\<placeholder: exported Word document and Excel sheet\>_

### Valves

| Valve | Default | Description |
|---|---|---|
| `language` | `en` | Language of the buttons, messages, default file names and print footer (`en` or `de`) |
| `timezone` | _(empty)_ | IANA time zone for file name timestamps and the footer date, e.g. `Europe/Berlin`. Empty means server time, which is often UTC in Docker |
| `max_mailto_length` | `1900` | Longest `mailto:` link handed to the mail program (Outlook on Windows handles about 2000 characters). Longer answers are downloaded as `.eml` |
| `always_eml` | `false` | Always download a formatted `.eml` draft instead of a plain-text `mailto:` link |
| `print_footer` | `true` | Add the "Created with the AI assistant on …" footer to printed pages |

### Security & privacy

- Everything is generated in memory on the server and sent straight to the user's browser. There is
  no network access and no shell, and nothing is written to disk or stored.
- The browser receives the generated file (base64), the `mailto:` link or the print page, wrapped in
  a fixed JavaScript snippet delivered through Open WebUI's `execute` event. All dynamic values are
  JSON-encoded.
- **CSP requirement:** the `execute` event runs code with `new Function(...)`. If your reverse
  proxy sets a `Content-Security-Policy`, it must allow `'unsafe-eval'`, otherwise the buttons do
  nothing. Open WebUI sets no CSP by default.
- The print page runs in a hidden same-origin iframe, so model output is never trusted as HTML:
  - raw HTML in the answer is always escaped;
  - links are only created for `http`, `https`, `mailto` and `tel`;
  - images are not loaded and are shown as links instead, so there are no tracking pixels;
  - the page carries its own strict CSP (`default-src 'none'`).
- Excel cells are written as text or numbers, never as formulas, so a cell such as `=HYPERLINK(...)`
  stays plain text.
- The e-mail is never sent by the plugin. The user's own mail program opens it as a draft.

### Notes and limitations

- **Ambiguous numbers in Excel:** a dot or comma followed by exactly three digits is read as a
  thousands separator, so `2.000` becomes 2000 and `1,234` becomes 1234. Integers with a leading
  zero (postcodes, IDs) and numbers with more than 15 digits stay text.
- **`.eml` drafts:** Outlook (Windows/Mac) opens them as editable drafts thanks to the
  `X-Unsent: 1` header. Thunderbird and Apple Mail open them as received messages, so use
  "Edit as new" or forward them.
- **Button labels** are language-neutral ("Word (.docx)", "Excel (.xlsx)", "PDF", "E-Mail"): Open WebUI
  builds the button list without loading the valves, so translated labels would only switch after the
  first click. Messages, file names and the print footer follow the `language` valve.

---

## Context Guard

Estimates the prompt size right before each request is sent to the model and warns or stops
before the context window overflows.

- Runs on the `request` hook (new in Open WebUI 0.11). It runs **after** document excerpts
  (RAG/files) have been injected and again after every tool-call round, so everything the model
  will see is counted.
- Once the prompt uses `warn_percent` of the usable context, a status line appears above the
  answer. The usable context is the context size minus `answer_reserve_tokens`.
- If the prompt would not fit, the request is stopped with a friendly explanation and next steps:
  start a new chat, paste less text, attach fewer documents.

Why this helps: **llama-server** and **vLLM** reject oversized requests with a technical HTTP 400
error ("exceeds the available context size"). **Ollama** silently cuts off the oldest part of the
prompt, so the model quietly forgets earlier messages or documents.

The token count is an **estimate** (characters ÷ `chars_per_token`, plus a flat amount per image
and the JSON size of tool definitions and tool calls), not a real tokenizer. Lower
`chars_per_token` to be more cautious.

> Screenshot: _\<placeholder: warning status line\>_
>
> Screenshot: _\<placeholder: "request too large" message\>_

### Valves

| Valve | Default | Description |
|---|---|---|
| `priority` | `100` | Filter order (lower runs first). Keep it high so content added by other filters is counted |
| `language` | `en` | Language of the warning and error messages (`en` or `de`) |
| `context_tokens` | `32768` | Context size of your model: llama-server `--ctx-size`, vLLM `--max-model-len`, Ollama `num_ctx` |
| `model_context_tokens` | _(empty)_ | Optional per-model sizes as `model_id=tokens` pairs, e.g. `qwen3-32b=40960, llama3.1:8b=8192`. Other models use `context_tokens` |
| `answer_reserve_tokens` | `6000` | Tokens kept free for the answer, including reasoning |
| `chars_per_token` | `3.5` | Characters per token for the estimate (English about 4, German about 3.5 to 4.3) |
| `tokens_per_image` | `300` | Estimated tokens per attached image |
| `warn_percent` | `75` | Show the hint from this share of the usable context (percent) |

If `answer_reserve_tokens` is not smaller than the context size, the check is skipped and a
warning is logged, so a misconfiguration never blocks every chat.

### Security & privacy

Read-only: the request is never modified. No network access, nothing is stored, and no message
content is logged.

---

## Company Context

Adds a short system note to every request, the same for every model and assistant, including newly
created ones. Example with `company_name = Example Ltd` and `timezone = Europe/Berlin`:

```text
[Company context]
You are the internal AI assistant of Example Ltd.
Today is Monday, 12 October 2026.
Always answer in English unless the user writes in another language or explicitly asks for a different one.
Address the user formally. Texts you write for them (e-mails, letters) should also use a formal tone unless they ask otherwise.
The user's name is Ada Lovelace.
```

With `language = de` the note is written in German, with German dates
(`Heute ist Montag, 12. Oktober 2026.`) and *Sie*/*du* as the form
of address.

- If the request already has a system message (model or chat system prompt), the note is put in
  front of it and the existing prompt is kept. Otherwise a new system message is added.
- The note starts with a marker line, so it is never added twice to the same request.
- The note contains the date but not the time. It opens the prompt, so it has to stay the same all
  day: then the model server (llama-server, vLLM, Ollama) can reuse its prompt cache instead of
  processing the whole chat again on every turn.

> Screenshot: _\<placeholder: answer that uses the current date / user name\>_

### Valves

| Valve | Default | Description |
|---|---|---|
| `priority` | `0` | Filter order (lower runs first) |
| `language` | `en` | Language of the note and date format (`en` or `de`) |
| `company_name` | _(empty)_ | Your company name. Empty leaves out the company line |
| `response_language` | _(empty)_ | Language the assistant should answer in, written in the note's language (e.g. `English`, `Deutsch`, `French`). Empty means the note's language |
| `form_of_address` | `formal` | `formal` or `informal` (German: *Sie* / *du*), also for texts written for the user |
| `timezone` | `UTC` | IANA time zone that determines today's date, e.g. `Europe/Berlin`. Invalid values fall back to server time |
| `include_user_name` | `true` | Tell the model the user's display name, e.g. for e-mail signatures |
| `extra_rules` | _(empty)_ | Additional house rules appended to the note |

### Security & privacy

No network access, nothing is stored. Only the system message is changed and user messages are left
untouched. With `include_user_name` on, the user's display name from Open WebUI is sent to the model
(line breaks removed, at most 100 characters).

---

## Reasoning Toggle

Adds a lightbulb button to the chat input. While it is on, the model thinks before it answers, and
Open WebUI shows the reasoning as the usual collapsible block above the answer. Made for local
reasoning models on **vLLM** and **llama-server**, e.g. Qwen3.x, Gemma 4, GLM or gpt-oss.

- Users switch it on in the **Integrations** menu (the button next to **+** in the chat input). While
  it is on, a lightbulb shows in the chat input; clicking it switches it off again. Open WebUI shows
  toggle filters in the chat input only while they are on, so there is no permanently visible
  switch. **Default Filters** in the model settings can switch it on by default.
- When it is on, the request gets `chat_template_kwargs: {"enable_thinking": true}`. vLLM and
  llama-server both hand this to the model's chat template. An optional effort level and token budget
  per model come on top (valves below).
- When it is off, the filter does not run at all, and the model behaves as configured. Models that
  think by default, such as Qwen3.x, keep thinking, so switch thinking off as their default (next
  section).

> Screenshot: _\<placeholder: lightbulb button in the chat input and the reasoning block\>_

### Why not Open WebUI's Model Controls?

Since 0.12.0, Open WebUI has **Model Controls**: a Thinking control in the model editor whose options
set custom parameters, picked from a knobs button in the chat input. They only show for users with
both **Allow Chat Controls** and **Allow Chat Params** (`USER_PERMISSIONS_CHAT_CONTROLS`,
`USER_PERMISSIONS_CHAT_PARAMS`), and the server ignores the choice otherwise. Those two permissions
also open the Chat Controls side panel, where users can change temperature, max tokens, context size
and the other advanced parameters. There is no permission for Model Controls alone.

A toggle filter needs neither permission. If your users get no chat controls, this filter is the way
to let them switch thinking on. If they may use chat controls anyway, Model Controls are the built-in
alternative.

### Make "no thinking" the default

For models that think by default, open the model's settings (**Admin Panel → Settings → Models**,
or **Workspace → Models** for custom models), go to **Advanced Params** and set **Reasoning Effort**
to `none`. vLLM and llama-server then switch thinking off, and while the button is on, the filter
replaces `none` with the template's default or the configured level. Title and tag generation stay
fast as well. Gemma 4 does not think unless asked to, so it needs no default.

### Effort levels and budget

| Valve | Default | Description |
|---|---|---|
| `priority` | `0` | Filter order (lower runs first) |
| `model_reasoning_effort` | _(empty)_ | Optional effort level per model as `model_id=level` pairs, e.g. `qwen3.8-27b=medium, gpt-oss-120b=high`. Without an entry, the chat template's default applies (Qwen3.8: `xhigh`, gpt-oss: `medium`) |
| `thinking_budget` | `0` | Maximum reasoning tokens per answer; `0` means no limit |
| `model_thinking_budget` | _(empty)_ | Optional per-model budgets as `model_id=tokens` pairs, e.g. `gemma-4-26b=4000`. `0` switches the limit off for that model |

Workspace models use the entry of their base model unless they have one of their own.

- **Effort levels depend on the model.** The level is passed to the chat template, which decides what
  it means. Qwen3.8 knows `low`, `medium` and `xhigh` and rejects `high`, so the request fails; gpt-oss
  knows `low`, `medium` and `high`. Gemma 4 has no levels and ignores the setting. Check the model card
  or chat template before you add an entry.
- **The budget is a hard cap.** When it is used up, the server ends the reasoning and the model has
  to answer. Too small a budget makes answers worse: on a counting task, Gemma 4 gave a wrong answer
  with 64 tokens and the right one with 300. The budget is sent as `thinking_budget_tokens`
  (llama-server) and `thinking_token_budget` (vLLM); each server ignores the other's name.
- **vLLM needs `--reasoning-parser`** (and, for some models, `--reasoning-config`) for the budget,
  otherwise it rejects every request that carries one. Leave the budget at `0` until that is set up.
- If you use [Context Guard](#context-guard), keep the budget below its `answer_reserve_tokens`, as the
  reasoning counts towards the answer.

### Security & privacy

No network access, nothing is stored or logged. Only `chat_template_kwargs`, `reasoning_effort` and
the two budget fields are set; messages are left untouched. Enable the filter only for local models:
cloud APIs such as OpenAI's reject unknown parameters like `chat_template_kwargs`.

---

## Office Documents

A tool that lets the model create Office files itself: "Write the minutes as a Word file", "Make
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

> Screenshot: _\<placeholder: answer with an attached presentation and its preview\>_
>
> Screenshot: _\<placeholder: generated Word document, Excel sheet with chart, slides\>_

### Setup

1. Install it under **Workspace → Tools** (not under Functions), see [Installation](#installation).
2. Enable it for a model: **Workspace → Models → (model) → Edit**, tick **Office Documents** under
   *Tools*. Users can also switch it on per chat in the **Integrations** menu of the chat input.
3. In the same model settings, set **Advanced Params → Function Calling** to **Native**. The tool
   needs a model and server with tool calling (llama-server with `--jinja`, vLLM with
   `--enable-auto-tool-choice` and the model's `--tool-call-parser`).
4. Set **max_tokens** there as well (e.g. `8192`), so a model that gets stuck in a loop while
   writing the arguments stops after a while.

### What the model sends

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
- **Excel formulas** belong to a column and are filled into every data row, e.g. `=B{row}*C{row}`
  (`{row}` is the current row). A formula written for the first data row (`=B2*C2`) gets `{row}`
  automatically. Allowed are cell references of the same sheet, numbers, quoted text,
  operators and these functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `COUNT`, `COUNTA`,
  `COUNTIF`, `SUMIF`, `AVERAGEIF`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `ABS`, `PRODUCT`, `POWER`, `SQRT`, `MOD`, `INT`. German names (`SUMME`, `WENN`,
  `RUNDEN` …), `;` as separator and decimal commas (`=RUNDEN(B{row}*0,19;2)`) are translated. The
  total row's `count` counts all filled cells (`COUNTA`). Formulas may also be written into single cells,
  e.g. a balance row `=B2-B3`, `=C2-C3` … when the months are columns; they are checked against the same
  whitelist. In columns the model marks as `text`, values that start with `=` stay text.
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

### Valves

| Valve | Default | Description |
|---|---|---|
| `language` | `en` | Language of the status messages and of texts inside the files: total row label, page footer, date format, table of contents heading (`en` or `de`); also the decimal mark for numbers sent as text in Excel |
| `delivery` | `attach` | `attach`: store the file in Open WebUI and attach it to the answer (preview, stays in the chat). `download`: hand it straight to the browser as a download (nothing is stored, but it is gone after a reload). If attaching fails, the tool falls back to a download |
| `currency` | `EUR` | Currency of `currency` columns in Excel (`EUR`, `USD`, `GBP`, `CHF`) |
| `accent_color` | `1F4E79` | Accent colour of all three formats (titles, table header rows, charts, title slide), as a hex value without `#`. Lighter and darker tones are derived from it |
| `font` | `Calibri` | Font of all three formats, e.g. a company font such as `Arial` or `Aptos`. It has to be installed where the files are opened; otherwise Office shows a similar one. Excel columns get a little more room for fonts wider than Calibri |
| `page_numbers` | `true` | Add "Page X of Y" to the footer of Word documents |

### Security & privacy

- **No code from the model.** It only supplies data, which is validated against the schema before
  fixed code builds the file. There is no shell and no network access, and nothing is loaded from
  URLs (images are not supported for that reason).
- **Excel:** every formula, in a column's `formula` field or in a cell, is checked against the whitelist
  above: references to other sheets (`!`) and other workbooks (`[…]`), DDE (`|`) and every other
  function, e.g. `WEBSERVICE`, `HYPERLINK` or `INDIRECT`, are rejected. A cell value that does not pass,
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

### Notes and limitations

- **Previews in Open WebUI** use browser libraries that know only part of the formats. Word files are
  written so that the preview shows their fonts, colours and bullets, too. The Excel preview shows formula
  cells empty (formulas have no stored results; Excel and LibreOffice calculate them when the file is
  opened) and dates as numbers, and the slide preview leaves out some formatting such as bullets and the
  alignment of titles. In Excel, PowerPoint and LibreOffice the files look as intended.
- **Table of contents:** Word fills it in when the document is opened and asks first whether to
  update the fields. LibreOffice needs **Tools → Update → Update All**.
- **Design:** one calm, modern look for all three formats, in shades of the `accent_color` valve and in the
  font of the `font` valve (also for charts and for text the user adds later). Company templates are not
  supported yet.
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

---

## Compatibility

- Tested with **Open WebUI 0.11.4** (the version in `required_open_webui_version`). Context Guard
  needs at least 0.11 because it uses the `request` filter hook.
- Python 3.11 and 3.12. Response Export was tested with python-docx 1.1.2 and newer, and
  openpyxl 3.1. Office Documents was tested with python-docx 1.2.0, openpyxl 3.1.5 and
  python-pptx 1.0.2, the versions bundled with Open WebUI 0.11.4.
- No extra packages: everything used is bundled with Open WebUI. The plugins do not depend on
  `ENABLE_PIP_INSTALL_FRONTMATTER_REQUIREMENTS`.

## Installation

1. In Open WebUI, open **Admin Panel → Functions** and click **+** (new function). Alternatively,
   use **Get** on the plugin's page on openwebui.com and import it into your instance.
2. Paste the content of the plugin file and save it. The ID and name come from the file's header.
3. Enable the function with its toggle.
4. Make it active:
   - **Global:** in the function's **⋯** menu, enable **Global** to apply it to all models.
   - **Per model:** in **Workspace → Models → (model) → Edit**, tick the plugin under
     *Actions* (Response Export) or *Filters* (Context Guard, Company Context, Reasoning Toggle).
   - Reasoning Toggle is best enabled **per model**, only for models that can reason. Enabled
     globally, its button also shows up for models where it does nothing.
5. Click the gear icon of the function to set its **Valves** (language, context size, company name …).

Office Documents is a tool, not a function: install it under **Workspace → Tools → +**, set its
valves with the gear icon there and enable it per model (see [Setup](#setup)).

Recommended order when using both filters: Company Context `priority = 0`, Context Guard
`priority = 100`, so the guard also counts the added note.

## Tests

The `tests/` folder contains a pytest suite covering all four plugins (document structure, number
formats, e-mail/mailto/.eml, print HTML, both languages, warn/block thresholds, idempotence,
reasoning parameters per model …):

```bash
python -m venv /tmp/owui-test && /tmp/owui-test/bin/pip install python-docx openpyxl python-pptx pydantic pytest
/tmp/owui-test/bin/python -I -m pytest -q tests
```

Office Documents has its own file, `tests/test_office_documents.py` (all three formats, formula
whitelist, limits, repaired model input, content checks, error messages, run time of the Markdown
parser). Its tests for attaching files to a chat replace Open WebUI's internals with stubs and need
`starlette` (`pip install starlette`); without it they are skipped.

Optional checks run when the tools are installed:

- **Node.js** runs the generated browser snippets against a small DOM stub.
- **A headless Chromium** renders the print page to a PDF (`chrome-headless-shell`, or set
  `PRINT_TEST_BROWSER`).
- **LibreOffice** converts the Word file to PDF to check list numbering, and renders the Word, Excel
  and PowerPoint files of Office Documents.
- **poppler-utils** is needed for both PDF checks.

## License

[MIT](LICENSE) © \<AUTHOR NAME\>
