# Changelog

Each plugin has its own version (in the header of its file). Earlier changes are in the git history.

## 2026-10-24

Office Documents 0.4.2 and Office File UI (no version). The other plugins are unchanged.

### Office Documents 0.4.2

- The stored formula results are rounded as Excel shows them (at most 15 digits and 10 decimals): the
  preview showed `180 - 210.4` as -30.400000000000006 and a sum of such differences as
  -4.249999999999986 instead of -30.4 and -4.25.

### Office Documents 0.4.1

- Excel files store the results of their formulas, so Open WebUI's preview, which does not compute,
  shows them instead of empty cells (differences, totals, `IF` texts). Until now openpyxl wrote the
  formulas only. Values are stored for arithmetic, comparisons, `&` and SUM, AVERAGE, MIN, MAX, COUNT,
  ROUND, ABS, IF, AND, OR and NOT. LibreOffice shows stored values without computing them again, so only
  results that are certain are stored: other functions and the formulas that use them, errors such as
  division by 0, leftovers such as 0.3 − 0.1 − 0.2, booleans in ranges (LibreOffice counts them as 1 and
  0, Excel skips them) and decimal numbers joined to text stay empty in the preview, as before. Excel
  computes all formulas when it opens the file. `ROUND` rounds the number as Excel shows it, with 15
  digits (`ROUND(655.595; 2)` is 655.60).

### Office File UI

- A Word, Excel or PowerPoint file opened from an answer no longer shows the "Content" tab (the
  extracted text, e.g. raw Excel formulas) for a moment before it switches to "Preview".
- The chip of a file that was deleted on the server since (by an admin, a cleanup job or retention rules)
  is greyed out and has no download icon. Its tooltip (Open WebUI's, with the file name) says "No longer
  available" below the name, also on a click; the click no longer opens an empty file view. The script
  checks each file of a shown answer once per page load (`/api/v1/files/<id>/process/status`, a small
  answer), and a download that gets a 404 greys out the chip as well. Open WebUI answers a file the user
  may not read with the same 404, hence the wording. Only Open WebUI's own 404 counts: a missing route or
  a proxy's error page leaves the chips as they are.
- A failed download shows "Download failed" below the name in the same tooltip for a few seconds
  instead of only turning the icon red (in the file view, which has no such tooltip, the icon's title
  says it).

## 2026-10-23, second update

Company Context 1.2.0. The other plugins are unchanged.

### Before you update

- **The date and the user's name move to the end of the note**, after the language, the form of
  address and `extra_rules`. So the start of the note is the same for all users and the model server
  can share its cached prompt between them. Nothing has to be changed.

### Company Context 1.2.0

- `company_profile`: short facts about the company for every user (sites, departments, abbreviations,
  who to ask). The model is told not to make up company details that are not given. A note above about
  2000 tokens (~8000 characters) gives a warning in the log.
- `group_profiles`: texts for single user groups such as departments, one `## <group name>` section per
  group. A user gets the sections of their groups (in Open WebUI 0.12 also through a subgroup), and with
  `include_group_names` (default on) the model learns which of these groups the user belongs to.
  Section names that match no group are written to the log.

## 2026-10-23

Office Documents 0.4.0 · Response Export 1.3.0 · Context Guard 1.4.0. The other plugins are unchanged.

### Before you update

- **Context Guard finds the context size on its own** where the server reports it: Ollama's `num_ctx`
  of the model and vLLM's `max_model_len`. If `context_tokens` was set lower than vLLM's
  `--max-model-len` on purpose, add the model to `model_context_tokens`, which still wins. This also
  works with Redis (several workers).
- **Office Documents writes the user's name into the file properties** (author), like Word does. Until
  now the files named `python-docx`, a python-pptx developer or `openpyxl` there.
- Office Documents: Excel cells may now be `true`/`false`, so the tool's parameters changed. Models see
  the new parameters with the next request; nothing else has to be done.
- Response Export reads numbers such as `2.000` by the other numbers of their column and, if those do not
  show it, by the language: in an English export a column of nothing but `2.000` and `15.000` becomes 2
  and 15 (German: 2000 and 15000).

### Office Documents 0.4.0

Excel:
- Amounts written as text in one column are read with one decimal mark: the values that show it
  (`2.350,50 €`, `980,00 €`) decide it for the whole column, also against the `language` valve. Until now
  `1.500 €` next to them became 1.5 with `language = en` (and `$1,500` next to `$2,350.50` 1.5 with `de`),
  without a word. A column whose values use both marks next to a value they decide (`1.500`) is sent back.
- Years, numbers and codes have no thousands separator: a `Jahr` column shows 2024 instead of 2.024, also
  `Nr.`, `Pos.`, `Kundennr.`, `Rechnungsnummer`, `ID`, `PLZ`, `Invoice no.` and similar identifier names
  (whole numbers in a column without a total). `Anzahl` or `Average number` stay amounts.
- A chart sent with only a type and title takes its labels from the first text or date column (the
  products next to `Pos.`, not 1, 2, 3) and does not plot years and numbers as values (no `Jahr` bars next
  to `Umsatz`).
- A cell that is one link (`[Intranet](https://…)`) is a clickable link; in other cells the address stays
  after the label (`Wiki (https://…)`). Until now only the label was left.
- `true` and `false` are TRUE and FALSE (WAHR and FALSCH in German Excel) instead of the numbers 1 and 0.
  In Word and slide tables they are Yes and No (Ja and Nein) instead of being sent back.
- ISO timestamps (`2026-10-23T08:30:00Z`, `…+02:00`, `… 08:30:15.250`) are dates with their time as
  written instead of being sent back; the message for other dates with a time names `YYYY-MM-DD hh:mm`.
- Dates before 1 March 1900 are written as text (`18.01.1871`); Excel showed `####` or a wrong day. A formula
  that computes with them (`=B{row}+14`) is sent back, as Excel would show #VALUE!.

Word:
- Code blocks between ``` or ~~~ lines in a paragraph (a language name may follow the opening one) are a
  monospace block with their lines and indentation as written. Until now a `# comment` in them became a
  heading (also in the table of contents), indentation was lost and the ``` stayed visible.
- `header_text` is plain text like the title and the footer: `**ACME** GmbH` showed the asterisks.

All formats:
- Windows paths keep their backslashes: `\\fs01\Vertrieb\_Archiv` became `\fs01\Vertrieb_Archiv`,
  `C:\Daten\*.xlsx` became `C:\Daten*.xlsx`.
- The file properties hold the document title, the user's name as author and the time the file was made.
- A comma before a closing bracket (`"b",]`) is repaired; with Python 3.13 and later the message for it,
  when the repair is not enough, says to delete it.

In the replay of all 4,605 calls of the three benchmark runs, one more call builds its file (a trailing
comma) and none is sent back newly. The files differ in their properties, in the number format of 157
position and number columns and in the paths of 15 Word files.

### Response Export 1.3.0

Excel:
- Numbers that can be read two ways (`1.085`, `2.000`, `1,234`) follow the other numbers of their column:
  next to `1.25`, `1.085` is 1.085 (before: 1085); next to `1.250.000`, `2.000` is 2000. If the column does
  not show it, the language decides (English: dot as decimal mark, German: comma). `0.125` is now a number
  instead of text.
- Dates such as `15.03.2024`, `1.3.2024` and `2024-03-12` become real dates, shown as `15.03.2024` in German
  and `2024-03-12` in English, so they can be sorted and filtered. `03/12/2024` stays text: month and day
  could be either way round.
- A cell that is one link (`[Docs](https://…)`) is a clickable link. Other links in a cell keep their
  address as `text (url)` instead of losing it.
- A heading directly above a table names its sheet instead of "Table 1", "Table 2" …

Word, PDF and e-mail:
- If the answer cites sources (`[1]`, `[2, 3]`), the file and the e-mail end with a numbered list of them,
  with name and web address as under the answer in the chat. Before, the numbers led nowhere. Numbers in
  brackets that are no source (`[2024]`) do not count.
- Code blocks and inline code are exported as written: an HTML example with `<details>` or a `<think>` tag
  in backticks is no longer cut out. Open WebUI's own reasoning and tool-call blocks are still always
  removed.
- Collapsible sections the model wrote itself (`<details>`, e.g. a solution) are exported open, with the
  summary as a bold line; before they were left out. Reasoning, tool calls and code interpreter runs are
  still removed.
- Embedded images and charts of the code interpreter show their description (or "[Image]"). Before, an
  embedded chart came out as pages of characters (a 400 KB chart gave a 113-page PDF) and a code
  interpreter chart as a file path.
- Formulas (`$a*b*c$`, `\[ E = mc^2 \]`) are kept as written; before, `*` turned into italics and `\[` lost
  its backslash. Amounts such as `$5 and $10` are not affected.
- `&amp;`, `&lt;`, `&nbsp;` and other HTML entities show as the characters they stand for.
- An answer with a broken emoji character no longer makes the Word and e-mail export fail.

PDF:
- Directory trees and boxes drawn with line characters keep their shape (`├── src/` becomes `+-- src/`)
  instead of turning into question marks. Arrows and symbols such as ↑ ↓ ▶ ▲ ▼ get a stand-in too.

### Context Guard 1.4.0

- The request's `max_tokens` is kept free if it is larger than `answer_reserve_tokens`. vLLM rejects a
  request whose prompt and `max_tokens` together exceed `--max-model-len`: with `max_tokens` 8192 as a
  model or default parameter and a 32k context, the guard let a 25,000-token prompt through and vLLM
  answered with an HTTP 400 error. A `max_tokens` at or above the context size is left to the server.
- The context size: an entry in `model_context_tokens`, else Ollama's `num_ctx` of the model (sent with
  the request), else the `max_model_len` that vLLM reports in its model list (also for workspace models
  on a vLLM model), else `context_tokens`. llama-server reports no size that fits, so it still needs
  `context_tokens` or `model_context_tokens`.
- `answer_reserve_tokens` keeps at most half the context free. With a small context (an Ollama `num_ctx`
  of 4096 and the default reserve of 6000), the guard used to switch itself off and log a warning on every
  request.
- Context Compaction: the room for the summary request is checked in the context of the compaction model
  if one is set (`CONTEXT_COMPACTION_MODEL`), with the summary length of the task model parameters
  (`max_tokens`, 1000 if they set none). Before, a small compaction model could hide the hint although its
  summary request did not fit.
- New integration test with a vLLM-like server that reports `max_model_len`, checked with 0.11.4 and
  0.12.3.

## 2026-10-22, third update

Response Export 1.2.2. The code of Context Guard, Reasoning Toggle and Office File UI was tidied up
without a change in behavior; updating them is optional. Company Context and Office Documents are
unchanged.

### Response Export 1.2.2

- Excel uses the same font as Word and PDF: a name in the `font` valve that is not a valid font name
  (or has extra spaces) now gives Calibri there too, instead of being written into the file as it is.
- Large files (a Word file on a template with images) download faster: the browser decodes them in a
  plain loop instead of calling a function for every byte.

## 2026-10-22, second update

Office Documents 0.3.2. The other plugins are unchanged.

### Office Documents 0.3.2

- A workbook in which no sheet has rows is sent back instead of being attached as an empty file. Models
  sometimes send such a call next to the real one (a "Test" workbook, or an empty copy with the same title),
  and the empty file then sat next to the real one in the chat. A template to fill in still works with rows
  of empty cells (`null`).

## 2026-10-22

Office Documents 0.3.1. The other plugins are unchanged.

### Office Documents 0.3.1

Excel sum rows:
- A formula sent in a row of a formula column is kept and replaces the column formula in that row. Until
  now the tool dropped it without a word, so net, VAT and gross rows below the items
  (`=SUM(F2:F6)`, `=F7*0.19`, `=F7+F8`) got the line item formula and showed 0. In two benchmark runs this
  decided 28 failed runs of offers and order lists.
- Column formulas with nothing to compute in such a sum row (quantity × price where the row has no
  quantity) are left out there instead of showing 0. The column formula written out for each row (`=B2*C2`)
  is no sum row; the column formula applies, as before.
- Charts and the auto filter end above the sum rows at the end of a table, so a pie chart does not show
  Netto, MwSt and Brutto as slices and sorting does not move them.
- A last row labelled "Summe", "Netto" etc. with its own formulas is the model's own total and no longer
  sent back as a total row; the message for a total row without formulas says how to write sums that
  `total` cannot build.
- Sent back: a formula that uses its own cell (`=SUM(F2:F7)` in F7), a subtotal (`SUM`, `SUMIF`) in a column
  with `total`, which counted it twice, and a formula that uses a cell of its own sheet that can only be
  meant by a wrong row count (`=F8*0.19` with an empty row 8 and the net in F9); the message lists the rows
  by their labels.

JSON repair:
- Two closing brackets in the wrong order (`null]}]` for `null]]}`), one bracket too many, an object whose
  `{` is missing, a stray quote after a number (`"level": 1"}`) and a key without its opening quote
  (`level": 2`) are repaired. Where a slip allows several repairs, they are tried in turn.
- A repair is only used if the result is valid arguments for the tool and no key comes twice with
  different values (a wrong repair can move the keys of one object into the one before it).
- Other tool parameters written after the value (`], "title": "T"`) are only used if nothing else follows;
  before, a text after them was dropped.
- In the replay of two benchmark runs, 73 of the 139 calls sent back for broken JSON text now build the
  file (Nemotron 3.5 Super 51 of 68, Qwen3.8-27B 17 of 32).

## 2026-10-21, third update

Context Guard 1.3.0 · Response Export 1.2.1 · Office File UI. The other plugins are unchanged.

### Context Guard 1.3.0

- The stop message is flowing text. Open WebUI shows errors as plain text, so the line breaks and
  bullet points of the old message ran together.
- If the user's own message makes up most of the prompt (a long pasted text), the stop message says
  so and suggests pasting only the relevant section or uploading the text as a file.
- Context Compaction: the threshold must leave about 1500 tokens for the summary request (summary
  prompt and the summary itself), not just be below the usable context. Otherwise the guard logs a
  warning and keeps showing its hint. With vLLM, a threshold close to the limit made the summary
  request fail.
- Recognizes when Context Compaction could not get a summary from the model and only cut the older
  messages short, and logs a warning once per chat.
- New integration tests with a real Open WebUI (`tests/integration/`), checked with 0.11.4, 0.12.0
  and 0.12.1.
- README: Open WebUI 0.12 sends its built-in tools with every chat, about 4,000 to 7,000 tokens
  depending on the model (measured for Qwen3.8, Gemma 4 and gpt-oss), and how to switch them off.

### Response Export 1.2.1

- Excel: all cells use the `font` valve. The header row had no font name, so LibreOffice showed it in
  a different font.
- Excel: numbers written with thousands separators (`1,240,000`, `1.234,56`) are shown with
  separators. Numbers without them, such as years and postcodes, stay as they are.

### Office File UI

- PowerPoint preview: titles and subtitles that the file aligns left are shown on the left. Open WebUI
  centres them even when the file says left ([open-webui#32873](https://github.com/open-webui/open-webui/issues/32873)).
  The fix is merged for the next Open WebUI release ([open-webui#32876](https://github.com/open-webui/open-webui/pull/32876));
  from then on the script's correction is not needed and does no harm.
- Pie and doughnut charts stay empty in the PowerPoint preview; the script cannot draw them
  ([open-webui#32884](https://github.com/open-webui/open-webui/issues/32884)).
- Spreadsheet preview in dark mode: grey instead of blue-violet, like the rest of the page. Open WebUI
  0.12.3 fixes this itself; the rule does no harm there.

## 2026-10-21, second update

Response Export 1.2.0. The other plugins are unchanged.

### Before you update

- **The PDF button downloads a PDF file** instead of opening the browser's print dialog, so nobody
  prints on paper by mistake any more. To print, open the file and print it from the PDF viewer.
- The PDF uses the standard PDF fonts Helvetica and Courier. The `font` and `word_template` valves
  still apply to Word only; `accent_color` now also colours the PDF.

### Response Export 1.2.0

- The plugin writes the PDF itself, without a new library: the content and design of the Word file
  (headings, lists numbered as in Word, quotes, code blocks, horizontal rules, tables with a coloured
  header row that is repeated on every page, banded rows and right-aligned number columns), clickable
  links (http, https, mailto), page numbers ("Seite 2 von 5") and the note of `print_footer`. A4 for
  German, Letter for English; the file name is made like the Word file's.
- Characters the standard PDF fonts lack: arrows and symbols such as → ≤ ✓ get stand-ins (`->`, `<=`,
  `+`), emoji are left out, other scripts become `?`. All German characters work.
- New icon for the PDF button: a document with the letters PDF.

## 2026-10-21

Office Documents 0.3.0 · Response Export 1.1.0 · Context Guard 1.2.0 · Company Context 1.1.0 ·
Office File UI (no version). Reasoning Toggle is unchanged.

### Before you update

- **New defaults apply to existing installations.** Open WebUI stores only the valves that were
  changed, so a valve that was never set takes the new default:
  - `language` is now `auto` in Response Export, Context Guard and Company Context: the plugins
    follow the user's interface language (English if there is no translation). Set `en` or `de` to
    keep one language for everybody.
  - `timezone` in Company Context was `UTC` and is now empty, which means the user's browser time
    zone. In Response Export, empty means the browser's time zone instead of the server's. Set an
    IANA name such as `Europe/Berlin` to keep a fixed zone.
- **Response Export makes different Word files.** They now look like the files of Office Documents:
  accent colour, coloured table header, banded rows, real links, A4 paper for German and Letter for
  English. The new valves `accent_color`, `font` and `word_template` work as in Office Documents.
- **Office Documents sends more mistakes back to the model.** Calls that 0.2 accepted can now get an
  error, for example a column formula that computes the same row everywhere, arithmetic on a text
  column, or a formula that is not allowed. The model usually fixes them in the next round. Formulas
  with HYPERLINK, WEBSERVICE, DDE or references to other workbooks are still written as text.
- XLOOKUP, IFS, MAXIFS, MINIFS, CONCAT and TEXTJOIN need Excel 2019/2021 or Microsoft 365 (XLOOKUP:
  Excel 2021 or 365, LibreOffice 24.8). Older versions show `#NAME?` in those cells.

### Office Documents 0.3.0

Excel formulas:
- Column formulas that refer to a fixed row instead of `{row}` (`=B3*C3`, `=B1*C1`) are sent back with
  the reason. A share like `=B2/SUM(B2:B4)` becomes `=B{row}/SUM($B$2:$B$4)`, and a range with one
  fixed end such as `$B$2:B9` becomes `$B$2:$B$9`.
- Arithmetic on a text column, on the header row or on single texts is sent back. Every formula
  message lists the columns (`A=Item, B=Quantity, …`), so the model can correct miscounted letters.
- Values the model sends in a formula column are compared with the result of simple formulas; a clear
  difference is sent back.
- A text such as `ca. 500` in a number column that is summed, charted or used by a formula is sent
  back instead of silently dropping out of the total.
- Percent columns: `C{row}%`, `*0.01` and `(C{row})/100` are recognised as dividing twice, like
  `/100` in 0.2.0.
- Formulas are parsed and checked for syntax and the number of arguments (`=B2*`, `=ROUND(B2)`).
- New functions, also with their German names: SUMIFS, COUNTIFS, AVERAGEIFS, VLOOKUP, INDEX, MATCH,
  DATE, TODAY, YEAR, MONTH, DAY, EDATE, EOMONTH, NETWORKDAYS, WEEKDAY, TIME, HOUR, MINUTE, LEFT,
  RIGHT, LEN, CONCATENATE, ISBLANK, ISNUMBER, ISTEXT, and XLOOKUP, IFS, MAXIFS, MINIFS, CONCAT,
  TEXTJOIN (see above).
- Whole columns (`SUM(B:B)`) stand for the data rows of the table; in a cell of that column they end
  above the cell. German quotes in formulas („ja“) are repaired, empty arguments are allowed.
- A formula that is not allowed is sent back to the model instead of being written as text.

Dates, times, numbers:
- New column type `time` (`08:00`), shown as `hh:mm`; durations in formula columns and totals as
  `[h]:mm`. The schema explains that hours are `(end − start) * 24`, and a time difference in a number
  column is sent back.
- `22.10.26` is read as a date; ambiguous or impossible dates (`05/11/2026`, `31.11.2026`) are sent
  back. ISO dates and times in a column without a type are recognised; values with a time get
  `DD.MM.YYYY hh:mm`.
- Formula columns without a type get the format of their result (date + days → date, quantity ×
  price → currency).
- Amounts in another currency (`$1,200.00`) keep that currency instead of the valve's; mixed
  currencies in one column are sent back. `’` and `'` work as thousands separators.

Other:
- Rows sent as objects with the column headers as keys, `null` in table cells and a `sheets` argument
  wrapped in an object are read. Keys like `Preis (€)` are no longer taken for garbled arguments.
- Up to four charts per sheet (`chart` may be a list); `unit` names the value axis, as on slides.
- Word and slides: items that already start with `1.`, `•` or `*` are not numbered twice, `#` lines
  in a paragraph become headings, `---` lines are dropped, Markdown tables in a paragraph have the
  same limits as table blocks, numbers in tables are right-aligned, and a letter without a title has
  no title block.
- Escaped line breaks are found in linear time; a long text full of `\n` could block the server.

### Response Export 1.1.0

- Word files in the design of Office Documents, with real links (http, https, mailto) and the valves
  `accent_color`, `font` and `word_template` (letterhead: header, footer, page size and margins).
- `---` becomes a line in Word; nested `<details>` blocks are removed completely.
- If the browser step fails (for example a Content Security Policy without `unsafe-eval`) or does not
  answer within 30 seconds, a message says so instead of the export hanging or failing silently.
- Excel: phone numbers like `+49 711 123` stay text, `CHF` amounts are numbers, `’` works as a
  thousands separator.
- A click without a sub-action exports to Word.
- `language: auto` (default) and an empty `timezone` are read from the browser.

### Context Guard 1.2.0

- Counts the model's system prompt, which Open WebUI adds only after the filter runs, so assistants
  with long instructions are no longer underestimated.
- Counts stored reasoning (`reasoning_content`, `reasoning`, `thinking`).
- Workspace models use the context size of their base model when they have no entry of their own.
- `language: auto` (default).
- Corrected: the filter priority does not decide whether Company Context's note is counted; it
  always is.

### Company Context 1.1.0

- The date is taken in the user's browser time zone (empty `timezone`, the new default), else the
  server's. An unknown time zone counts as empty instead of being an error.
- The note names the ISO calendar week.
- `language: auto` (default).

### Office File UI

- A chip without a known file reloads the chat at most three times (after 2, 8 and 32 seconds);
  failed requests do not count. Before, it reloaded every 2 seconds while an answer streamed.
- Page changes are handled at most every 150 ms.
- The tab bar of the preview is only hidden while the preview really shows, so a change in Open
  WebUI's markup cannot leave the user without the bar.
- The markup it relies on is the same in Open WebUI 0.11.4 and 0.12.1.
