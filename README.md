# Open WebUI plugins

Self-contained plugins for [Open WebUI](https://github.com/open-webui/open-webui), each in its own
folder with a README:

| Plugin | Type | What it does |
|---|---|---|
| [Office Documents](office-documents/) | Tool | Lets the model create Word, Excel and PowerPoint files and attaches them to its answer |
| [Response Export](response-export/) | Action | Adds buttons to export an answer as Word, Excel or PDF, or open it as an e-mail |
| [Context Guard](context-guard/) | Filter | Warns before the context window overflows and stops the request with a clear message |
| [Company Context](company-context/) | Filter | Adds company name, date, language, form of address, a short company profile and texts per user group to every request |
| [Office File UI](office-file-ui/) | Browser script | Optional: shows files attached to an answer as chips below the text and opens Office files in the preview |

Response Export, Context Guard, Company Context and Office Documents include English and German
texts (`language` valve); Response Export, Context Guard and Company Context follow the user's
interface language by default (`auto`). All of them use only the Python standard library and libraries that already ship with Open WebUI, so nothing
has to be installed.

What changed in each version, and what to check before updating, is in the [changelog](CHANGELOG.md).

![Files made by Office Documents: a Word document, an Excel sheet with a chart and slides](office-documents/images/files.png)

## Installation

1. In Open WebUI, open **Admin Panel → Functions** and click **+** (new function). Alternatively,
   use **Get** on the plugin's page on openwebui.com and import it into your instance.
2. Paste the content of the plugin file and save it. The ID and name come from the file's header.
3. Enable the function with its toggle.
4. Make it active:
   - **Global:** in the function's **⋯** menu, enable **Global** to apply it to all models.
   - **Per model:** in **Workspace → Models → (model) → Edit**, tick the plugin under
     *Actions* (Response Export) or *Filters* (Context Guard, Company Context).
5. Click the gear icon of the function to set its **Valves** (language, context size, company name …).

Office Documents is a tool, not a function: install it under **Workspace → Tools → +**, set its
valves with the gear icon there and enable it per model (see its [setup](office-documents/#setup)).
Office File UI is not a plugin at all; see its [README](office-file-ui/) for how to install it.

Context Guard always counts the note of Company Context, whatever their priorities: Open WebUI runs
all `inlet` hooks (Company Context) before any `request` hook (Context Guard). The priority only
orders several filters on the same hook.

## Compatibility

- Needs **Open WebUI 0.12.4** or newer (the version in `required_open_webui_version`). The whole
  test suite passes with 0.12.4 (Python 3.12 and the packages it bundles), Context Guard also with
  Context Compaction. For older versions, use the plugins from before 2026-10-27 (see the
  [changelog](CHANGELOG.md)).
- Python 3.11 and 3.12. Response Export was tested with python-docx 1.1.2 and newer, and
  openpyxl 3.1. Office Documents was tested with python-docx 1.2.0, openpyxl 3.1.5 and
  python-pptx 1.0.2, the versions bundled with Open WebUI 0.12.4.
- No extra packages: everything used is bundled with Open WebUI. The plugins do not depend on
  `ENABLE_PIP_INSTALL_FRONTMATTER_REQUIREMENTS`.

## Tests

The `tests/` folder contains a pytest suite for all plugins (document structure, number formats,
e-mail/mailto/.eml, PDF layout, both languages, warn/block thresholds, idempotence …):

```bash
python -m venv /tmp/owui-test && /tmp/owui-test/bin/pip install python-docx openpyxl python-pptx pydantic pytest pypdf
/tmp/owui-test/bin/python -I -m pytest -q tests
```

Office Documents has its own file, `tests/test_office_documents.py` (all three formats, formula
whitelist, limits, repaired model input, content checks, error messages, run time of the Markdown
parser). Its tests for attaching files to a chat replace Open WebUI's internals with stubs and need
`starlette` (`pip install starlette`); without it they are skipped.

The tests that read the PDF files of Response Export need `pypdf`; without it they are skipped.

`tests/integration/` checks Context Guard together with Open WebUI's Context Compaction in a real
Open WebUI: a mock model server with an 8k context (behaving like llama-server or like vLLM) and
chats with normal growth, a tool loop and a long paste. Each test starts its own Open WebUI, so the
tests take about 40 seconds. They are skipped unless `OWUI_BIN` points to an `open-webui`
executable, which allows checking several versions:

```bash
python -m venv /tmp/owui-0124 && /tmp/owui-0124/bin/pip install open-webui==0.12.4
OWUI_BIN=/tmp/owui-0124/bin/open-webui /tmp/owui-test/bin/python -I -m pytest -q tests/integration
```

The same folder also checks Company Context's `group_profiles` with real groups (with a subgroup)
and its form of address through Model Controls.

Optional checks run when the tools are installed:

- **Node.js** runs the generated browser snippets and the Office File UI script against small DOM
  stubs.
- **LibreOffice** converts the Word file to PDF to check list numbering, and renders the Word, Excel
  and PowerPoint files of Office Documents.
- **qpdf** checks the structure of the PDF files of Response Export.
- **poppler-utils** renders those PDF files, and is needed for the LibreOffice checks.

## License

[MIT](LICENSE) © tovren
