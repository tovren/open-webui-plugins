# Open WebUI plugins: Response Export, Context Guard, Company Context, Reasoning Toggle

Four self-contained plugins for [Open WebUI](https://github.com/open-webui/open-webui):

| Plugin | Type | File | What it does |
|---|---|---|---|
| [Response Export](#response-export) | Action | `response_export.py` | Adds buttons to export an answer as Word or Excel, print or save it as PDF, or open it as an e-mail |
| [Context Guard](#context-guard) | Filter | `context_guard.py` | Warns before the context window overflows and stops the request with a clear message |
| [Company Context](#company-context) | Filter | `company_context.py` | Adds company name, date, language and form of address to every request |
| [Reasoning Toggle](#reasoning-toggle) | Filter (toggle) | `reasoning_toggle.py` | Adds a button to the chat input that lets local reasoning models (vLLM, llama-server) think before they answer |

Response Export, Context Guard and Company Context are in English by default and include a German
translation (`language` valve); Reasoning Toggle shows no texts of its own. All four use only the
Python standard library and libraries that already ship with Open WebUI, so nothing has to be
installed.

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

## Compatibility

- Tested with **Open WebUI 0.11.4** (the version in `required_open_webui_version`). Context Guard
  needs at least 0.11 because it uses the `request` filter hook.
- Python 3.11 and 3.12. Response Export was tested with python-docx 1.1.2 and newer, and
  openpyxl 3.1.
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

Recommended order when using both filters: Company Context `priority = 0`, Context Guard
`priority = 100`, so the guard also counts the added note.

## Tests

The `tests/` folder contains a pytest suite covering all four plugins (document structure, number
formats, e-mail/mailto/.eml, print HTML, both languages, warn/block thresholds, idempotence,
reasoning parameters per model …):

```bash
python -m venv /tmp/owui-test && /tmp/owui-test/bin/pip install python-docx openpyxl pydantic pytest
/tmp/owui-test/bin/python -I -m pytest -q tests
```

Optional checks run when the tools are installed:

- **Node.js** runs the generated browser snippets against a small DOM stub.
- **A headless Chromium** renders the print page to a PDF (`chrome-headless-shell`, or set
  `PRINT_TEST_BROWSER`).
- **LibreOffice** converts the Word file to PDF to check list numbering.
- **poppler-utils** is needed for both PDF checks.

## License

[MIT](LICENSE) © \<AUTHOR NAME\>
