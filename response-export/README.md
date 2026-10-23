# Response Export

An Open WebUI **action** that adds four buttons below every assistant message:

| Button | Result |
|---|---|
| **Word (.docx)** | Word document with headings, bullet and numbered lists (one nesting level, numbering restarts for each list), tables with a header row in the accent colour, bold/italic/strikethrough/inline code, clickable links, code blocks, quotes and horizontal rules. Same design as the Word files of Office Documents (`accent_color`, `font`), optionally on a company template such as a letterhead (`word_template`) |
| **Excel (.xlsx)** | One worksheet per Markdown table, named after a heading directly above the table (otherwise "Table 1", "Table 2" …). The header row is bold, shaded and frozen, with an auto filter and fitted column widths. Numbers become real numbers, including German and Swiss formats (`1.234,56`, `12,5 %`, `1.234,56 €`, `CHF 1'234.50`), `$`/`£` amounts and the codes `EUR`, `USD`, `GBP` and `CHF`. Phone numbers such as `+49 711 123456` stay text. Numbers written with thousands separators keep them in Excel; years and postcodes without separators stay as they are. Dates written as `15.03.2024`, `1.3.2024` or `2024-03-15` become real dates (shown as `15.03.2024` in German, `2024-03-15` in English). A cell that is one link becomes a clickable link; other links stay visible as `text (url)`. All cells use the `font` valve |
| **PDF** | PDF file with the content and design of the Word file (`accent_color`): headings, lists, quotes, code blocks, tables whose header row is repeated on every page (number columns right-aligned), clickable links, page numbers and an optional small note in the footer. A4 for German, Letter for English. It is downloaded like the Word file; to print, open it and print it from the PDF viewer |
| **E-Mail** | Opens the default mail program through a `mailto:` link. A `Subject:`/`Betreff:` line becomes the subject, and any preamble before it ("Here is a draft:") is dropped. If the answer is too long for a `mailto:` link, an `.eml` draft with plain-text and HTML parts is downloaded instead. Outlook opens it as a new, editable message |

Reasoning, tool-call and code interpreter blocks of Open WebUI (`<details type="reasoning">…</details>`
etc.) are always removed before exporting, also nested ones. `<think>…</think>` blocks are removed
outside code; in code blocks and inline code they stay as written. Collapsible sections the model wrote itself (`<details>` with a
`<summary>`, e.g. a solution) are exported open: the summary as a bold line, then the text.

If the answer cites its sources with Open WebUI's markers (`[1]`, `[2, 3]`), the Word file, the PDF and
the e-mail end with a numbered list of the sources (name and web address, as under the answer in the
chat). File names use the first heading of the answer plus a timestamp, e.g.
`Quarterly report 2026-10-12 1405.docx`.

The plugin writes the PDF itself, so no PDF library is needed. It uses the standard PDF fonts
Helvetica and Courier, which every PDF viewer has (see [Notes and limitations](#notes-and-limitations)).

![The four export buttons below an answer](images/buttons.png)

![The exported Word document, Excel sheet and PDF file](images/exported-files.png)

## Install

File: [`response_export.py`](response_export.py). Install it under **Admin Panel → Functions** and
enable it globally or per model, see [Installation](../README.md#installation).

## Valves

| Valve | Default | Description |
|---|---|---|
| `language` | `auto` | Language of the messages, default file names, PDF footer and the paper size of Word and PDF files (`en`: Letter, `de`: A4): `auto` (the user's interface language, English if there is no translation), `en` or `de` |
| `timezone` | _(empty)_ | IANA time zone for file name timestamps and the footer date, e.g. `Europe/Berlin`. Empty means the time zone of the user's browser. An unknown name counts as empty |
| `max_mailto_length` | `1900` | Longest `mailto:` link handed to the mail program (Outlook on Windows handles about 2000 characters). Longer answers are downloaded as `.eml` |
| `always_eml` | `false` | Always download a formatted `.eml` draft instead of a plain-text `mailto:` link |
| `print_footer` | `true` | Add the note "Created with the AI assistant on …" to the footer of PDF pages. The page number is always there |
| `accent_color` | `1F4E79` | Accent colour of Word and PDF files (headings, table header rows, links), as a hex value without `#` |
| `font` | `Calibri` | Font of Word documents and Excel sheets, e.g. a company font such as `Arial` or `Aptos`. It has to be installed where the files are opened; otherwise Office shows a similar one. PDF files always use Helvetica and Courier |
| `word_template` | _(empty)_ | Path of a Word file (`.docx` or `.dotx`) on the Open WebUI server, e.g. a letterhead: Word documents get its header, footer, page size and margins. Not used for PDF files |

`accent_color`, `font` and `word_template` have the same names and meaning as in
[Office Documents](../office-documents/#word-template), so set them to the same values to get the same
look from both plugins. With a template, its text is removed; the content keeps the design and uses
the styles *Body Text*, *Heading 1–4*, *Quote* and the list styles, which are replaced in the document.
If the file cannot be read (missing, not a Word file, larger than 20 MB), the document is made without
it and Open WebUI's log says why.

`language = auto` and an empty `timezone` are read from the browser with a small extra request when a
button is clicked.

## Security & privacy

- Everything is generated in memory on the server and sent straight to the user's browser. There is
  no network access and no shell, and nothing is written to disk or stored.
- The browser receives the generated file (base64) or the `mailto:` link, wrapped in a fixed
  JavaScript snippet delivered through Open WebUI's `execute` event. All dynamic values are
  JSON-encoded.
- The `word_template` valve reads one file from the server (at most 20 MB). Only an administrator can
  set it.
- **CSP requirement:** the `execute` event runs code with `new Function(...)`. If your reverse
  proxy sets a `Content-Security-Policy`, it must allow `'unsafe-eval'`, otherwise the buttons do
  nothing. Open WebUI sets no CSP by default. If the browser does not run the snippet or does not
  answer within 30 seconds, an error message with this hint is shown.
- Model output is never trusted as HTML:
  - raw HTML in the answer is always escaped (shown as text);
  - links are only created for `http`, `https` and `mailto` in Word and PDF files, and also for
    `tel` in the e-mail; other link targets stay text as `text (url)`;
  - images are not loaded, so there are no tracking pixels: images on the web are shown as links,
    embedded images and files of the chat (such as charts of the code interpreter) as their
    description, or "[Image]" ("[Bild]" in German) if they have none.
- Excel cells are written as text or numbers, never as formulas, so a cell such as `=HYPERLINK(...)`
  stays plain text.
- The e-mail is never sent by the plugin. The user's own mail program opens it as a draft.

## Notes and limitations

- **Ambiguous numbers in Excel:** a dot or comma followed by exactly three digits can be either a
  thousands separator or the decimal mark (`1.085`, `2.000`, `1,234`). Excel reads them the way the
  other numbers of the same column show: in a column with `1.25` or `1,234,567`, `1.085` is 1.085; in
  a column with `12,5` or `1.250.000`, `2.000` is 2000. If the column does not show it, the `language`
  decides: in English the dot is the decimal mark (`1.085` is 1.085, `1,234` is 1234), in German the
  comma (`2.000` is 2000, `1,234` is 1.234). `0.125` is always a decimal. Integers with a leading zero
  (postcodes, IDs) and numbers with more than 15 digits stay text. Dates written month or day first
  (`03/04/2024`) stay text.
- **Formulas (LaTeX)** such as `$E = mc^2$` or `\[ … \]` are exported as written, with their
  delimiters. `$5 and $10` are amounts, not a formula.
- **PDF fonts:** the standard PDF fonts are not embedded and cover the characters of Windows-1252
  (Western European languages, including all German characters, `€`, `„“` and `–`). Other characters
  are replaced: common symbols get a stand-in (`→` becomes `->`, `≤` `<=`, `✓` `+`, `✗` `-`, `▶` `>`),
  the lines of directory trees and boxes become `+`, `-` and `|` (`├── src/` becomes `+-- src/`),
  letters lose accents the fonts lack (`č` becomes `c`), emoji are left out, and other scripts (Greek,
  Cyrillic, Chinese …) become `?`. The Word file keeps all characters. Headings in the PDF, as in Word, show no
  links or formatting.
- **`.eml` drafts:** Outlook (Windows/Mac) opens them as editable drafts thanks to the
  `X-Unsent: 1` header. Thunderbird and Apple Mail open them as received messages, so use
  "Edit as new" or forward them.
- **Button labels** are language-neutral ("Word (.docx)", "Excel (.xlsx)", "PDF", "E-Mail"): Open WebUI
  builds the button list without loading the valves, so translated labels would only switch after the
  first click. Messages, file names and the PDF footer follow the `language` valve.
- **Without a sub-action** (e.g. when the action is called through the API with the function's own
  ID), the answer is exported as Word.
