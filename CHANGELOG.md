# Changelog

Each plugin has its own version (in the header of its file). Earlier changes are in the git history.

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
