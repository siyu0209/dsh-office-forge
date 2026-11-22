# Office Forge

**Office document automation for DeepSeek Harness (DSH).**

Office Forge is a self-contained DSH bundle that gives an agent thirteen independent tools for working with Word, Excel, PowerPoint, and PDF files. It reads, creates, edits, audits, previews, converts, batch-processes, and templates documents without launching an Office UI. All document work runs in short-lived Python subprocesses; the Node layer only handles workspace authorization, argument validation, subprocess transport, and optional host integration.

The plugin is an independent implementation. It bundles no third-party office runtime and does not depend on any other document plugin.

## Highlights

- **One toolkit for four formats** — `.docx`, `.xlsx`, `.pptx`, and `.pdf`.
- **File-level by default** — works on disk files without requiring Microsoft Office.
- **Optional real backends** — Microsoft Office COM and LibreOffice are used only when present.
- **Template-friendly** — `{{key}}` placeholders and one-level `{{#each items}}...{{/each}}` loops for Word tables and Excel rows.
- **Audit built in** — Word heading quality, Excel formula health/profile/insight/diff, PowerPoint deck quality and theme catalogues.
- **Agent-friendly** — every tool accepts and returns lossless JSON, reports stable error codes, and is safe to call from a conversation.
- **Workspace-safe** — lexical and host-level path checks, atomic writes, bounded input/output, no recursion.
- **Configurable** — module and feature switches for Word, Excel, PowerPoint, and PDF capabilities.
- **Diagnosable** — `officeforge_doctor` and a read-only plugin-page card report exactly which optional backends are available.

## Requirements

| Component | Requirement |
|---|---|
| DSH | `>= 0.2.2-rc.5` |
| Node.js | `>= 20` |
| Python | Python 3.12 with `python-docx`, `openpyxl`, `python-pptx`; the DSH bundled runtime already includes these |
| PDF | `pypdf` is vendored inside the plugin; no external PDF service is required |
| Microsoft Office COM | Optional, Windows only, for native PivotTable, native SmartArt, revisions, PDF export, and legacy conversion |
| LibreOffice | Optional, for legacy `.doc/.xls/.ppt` conversion when COM is unavailable |
| OCR | Optional, requires PyMuPDF, pytesseract, and the Tesseract executable; `officeforge_doctor` reports missing pieces |

## Installation

### From a release tarball

```bash
dsh plugin --profile <profile> add /path/to/dsh-office-forge-0.14.0.tgz
```

Or add the package to the profile's dependencies and `dsh.profile.bundles` in the usual DSH way. The bundle exposes `dsh.bundle.patch`, `dsh.client`, locale files, and the icon required by the DSH plugin page.

### From source

```bash
git clone https://github.com/siyu0209/dsh-office-forge.git
cd dsh-office-forge
npm pack
```

Then install the produced `dsh-office-forge-*.tgz` into the target profile. Installing from source requires Node and a Python runtime with the three Office libraries; `officeforge_doctor` tells you what is missing.

### Verify the installation

After restarting DSH, call:

```json
{ "tool": "officeforge_doctor" }
```

Expected result: `ok: true` when required pieces are present. Optional backends (OCR, LibreOffice) may show `warn`; they never block the core document tools.

## Quick start

### Create a Word document

```json
{
  "path": "reports/weekly.docx",
  "type": "docx",
  "spec": {
    "title": "Weekly Report",
    "subtitle": "Week 47",
    "sections": [
      { "heading": "Progress", "bullets": ["Finished module A", "Validated flow B"] },
      { "heading": "Next week", "numbered": ["Fix C", "Ship D"] }
    ],
    "pageNumbers": true
  }
}
```

### Create an Excel workbook

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "spec": {
    "sheets": [
      {
        "name": "Metrics",
        "columns": ["Month", "Revenue"],
        "rows": [["Jan", 120], ["Feb", 180], ["Mar", 240]],
        "freeze": "A2",
        "autoFilter": true,
        "charts": [
          { "type": "bar", "dataRange": "B1:B4", "categoriesRange": "A2:A4", "anchor": "E2", "title": "Revenue" }
        ]
      }
    ]
  }
}
```

### Create a PowerPoint deck

```json
{
  "path": "decks/quarterly.pptx",
  "type": "pptx",
  "spec": {
    "theme": "swiss",
    "slides": [
      { "layout": "cover", "title": "Quarterly Review", "subtitle": "FY2026" },
      { "layout": "kpi", "title": "Highlights", "items": [
        { "value": "42%", "label": "Growth" },
        { "value": "18", "label": "New accounts" }
      ] }
    ]
  }
}
```

### Fill a Word template

```json
{
  "template": "templates/people.docx",
  "output": "out/people.docx",
  "data": {
    "title": "Team",
    "people": [
      { "name": "Alice", "dept": "Engineering" },
      { "name": "Bob", "dept": "Product" }
    ]
  },
  "options": { "expandLoops": true, "missing": "keep" }
}
```

### Audit a workbook

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "mode": "health"
}
```

### Merge PDFs

```json
{
  "operation": "merge",
  "inputs": ["out/a.pdf", "out/b.pdf"],
  "output": "out/merged.pdf"
}
```

## Tool reference

| Tool | What it does |
|---|---|
| `officeforge_read` | Read `.docx` / `.xlsx` / `.pptx` / `.pdf` in `content`, `outline`, or `search` mode; output markdown, text, or JSON. PDF supports text extraction and encryption passwords. |
| `officeforge_create` | Create a document from a declarative spec. Word supports sections, headings, bullets, numbered lists, tables, headers/footers, page setup, page numbers, TOC, images, hyperlinks, captions, styles, comments, revisions, mail merge, footnotes, bookmarks, cross-references, citations, bibliography, columns, watermarks, content controls, and `.dotx` templates. Excel supports sheets, CSV/TSV/JSON import, formulas, array formulas, charts (including combo charts), conditional formatting, data validation, comments, images, merged cells, tables, protection. PowerPoint supports slides, layouts, templates, themes, charts, images, tables, text boxes, shapes, connectors, transitions (including morph), animations, SmartArt-style diagrams, notes, sections, comments, media, and masters. |
| `officeforge_edit` | Apply bounded operations to an existing document. Word: replace, append, images, TOC, captions, styles, fields, comments, tracked revisions, mail merge, footnotes, bookmarks, cross-references, citations, bibliography, columns, watermark, content controls, templates, core properties. Excel: cells, rows, sheets, formats, freeze, filters, charts, conditional formatting, validations, comments, merges, images, formula columns, sorting, deduplication, blank filling, text cleanup, pivot summary, tables, protection. PowerPoint: text, slides, notes, sections, transitions, shapes, text boxes, connectors, animations, SmartArt-style diagrams, comments, media, masters, themes, core properties. |
| `officeforge_template` | Fill Word/Excel templates with JSON or `.json/.csv/.tsv` data: `{{key}}` placeholders and one-level `{{#each items}}...{{/each}}` loops. |
| `officeforge_audit` | Read-only audit. Word: `health` (heading jumps, duplicate/orphan headings, unresolved placeholders, long paragraphs, empty paragraph runs) and `profile`. Excel: `health` (formula errors, inconsistent patterns, hardcoded values in formula columns, blank-row breaks, circular references), `profile`, `insight`, `diff`. PowerPoint: `deck` quality, `profile`, `themes`, `layouts`, `templates`. |
| `officeforge_doctor` | Read-only runtime diagnostics: Python and Office libraries, bundled PDF engine, COM/pywin32, registered Office applications, LibreOffice, `officeKit`, client bundle, and active modules/features. Every check has `required`, `status`, `detail`, and `remedy`. |
| `officeforge_convert` | Convert legacy `.doc/.docm/.xls/.xlsm/.ppt/.pptm` through COM/LibreOffice; convert modern/PDF content through the host `officeKit`. |
| `officeforge_pdf` | File-level PDF operations: `merge`, `split`, `rotate`, `watermark`, `pageNumbers`, `formFields`, `fillForm`, `extractImages`, `ocr`, `encrypt`, `decrypt`. Uses vendored pypdf. |
| `officeforge_native` | Optional Microsoft Office COM operations: `convert`, `exportPdf`, `acceptRevisions`, `recalculate`, `pivot` (native PivotTable), `smartArt` (native SmartArt). |
| `officeforge_preview` | Render pages, slides, or worksheet ranges to PNG through `officeKit.renderImages`; no Office UI is launched. |
| `officeforge_recalculate` | Recalculate formulas through `officeKit.recalculate` and save a refreshed workbook. |
| `officeforge_batch` | Run up to 20 independent create/edit/template jobs in one Python call; each item succeeds or fails independently; nested batches are rejected. |
| `officeforge_export` | Convert through `officeKit` to PDF or another supported format; optionally copy the output back into the workspace. |

## Configuration

All fields are optional. Every module and feature defaults to `true`, so an empty config enables the full toolkit.

| Field | Default | Description |
|---|---|---|
| `pythonPath` | auto-detect | Python executable with `python-docx`, `openpyxl`, `python-pptx` installed. |
| `engineTimeoutMs` | `120000` | Timeout for one Python engine call, in milliseconds. |
| `maxBatchItems` | `20` | Maximum items in one `officeforge_batch` call. |
| `enableExport` | `true` | Register `officeforge_export`. |
| `enableBatch` | `true` | Register `officeforge_batch`. |
| `enableBackends` | `true` | Register `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate`; missing backends return explicit errors. |
| `modules` | all `true` | `word` / `excel` / `ppt` / `pdf` document-family switches. |
| `features` | all `true` | Fine-grained Word / Excel / PowerPoint capability switches. |

Example profile override:

```yaml
- insert:
    - id: dsh-office-forge
      name: dsh-office-forge
      config:
        modules: { word: true, excel: true, ppt: true, pdf: true }
        features:
          word: { images: true, toc: true, comments: true }
          excel: { charts: true, audit: true, import: true }
          ppt: { themes: true, audit: true, layouts: true }
```

Full feature names are defined in `resources/features.json`; theme presets are defined in `resources/themes.json`.

### Python discovery order

1. `pythonPath` from the bundle config
2. `OFFICEFORGE_PYTHON`
3. `DSH_PYTHON`
4. `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe`
5. `$DSH_HOME/dsh-runtimes/*/dependencies/python`
6. `python3` / `python` on `PATH`

## Architecture

```
DSH host
  └─ lib/index.js                 Node host layer
       ├─ Config schema           modules + features + runtime options
       ├─ path policy             ctx.fs.resolve / contains / processPath
       ├─ runEngine()             short-lived Python subprocess
       │    └─ lib/office_engine.py
       │         ├─ office_common.py      paths, atomic writes, placeholders, feature gates
       │         ├─ office_docx.py        Word read/create/edit/template/audit
       │         ├─ office_xlsx.py        Excel read/create/edit/template/audit
       │         ├─ office_pptx.py        PowerPoint read/create/edit/template/audit
       │         ├─ office_pdf.py         PDF text extraction
       │         ├─ office_pdf_tools.py   PDF merge/split/rotate/watermark/forms/images/OCR/encryption
       │         ├─ office_backends.py    optional COM / LibreOffice backends
       │         ├─ office_com_worker.py  isolated COM worker
       │         └─ office_migrations.py  spec schema migration
       ├─ officeKit bridge         export / preview / recalculate
       └─ lib/client.js            read-only plugin-page status card + environment check
```

The Node layer does four jobs: configuration trimming, workspace path policy, argument validation, and subprocess lifecycle. All document parsing and generation happen in Python. The Python engine is stateless between calls; each tool dispatches exactly one action.

## Security and isolation

- **Workspace-only writes.** Every input and output path goes through both the host's `ctx.fs.resolve` / `ctx.fs.contains` checks and a lexical `PATH_OUTSIDE_WORKSPACE` guard.
- **Atomic writes.** New files are written to a same-directory temp file and `os.replace`d. Existing files are refused unless `overwrite: true`; optional backups are created next to the output.
- **Bounded execution.** 50 MiB input, 100 MiB output, 200,000 text characters, 5,000 × 300 Excel cells per sheet, 300 slides, 20 batch items.
- **No recursion.** Tools never call each other; each tool sends exactly one action to the engine; templates expand one loop level only; nested batches are rejected; table text replacement uses an explicit stack.
- **No host hooks.** The plugin injects only `tools`, `fs`, and `workingDirectory`; it does not register `systemPrompt`, LLM stream hooks, tool middleware, or host hooks.
- **No tool-name collisions.** Every tool starts with `officeforge_`; built-in `office_*` tools and `run_code` are untouched.
- **Per-tool isolation.** Tools are registered one by one; a failed registration logs a warning and never aborts activation of the rest of the plugin.
- **Lossless JSON.** All tool results are normalized so absent optional fields are omitted instead of leaking `undefined`, which the DSH host rejects as `value is not lossless JSON`.
- **Read-only plugin page.** The client bundle registers only `plugins.detail.section` with a read-only status card; it never writes configuration and never performs document work.

## Error model

All failures return a stable JSON error object with `code`, `message`, and optional `detail`. Common codes include:

| Code | Meaning |
|---|---|
| `INPUT_NOT_FOUND` | Source file does not exist or is outside the workspace. |
| `PATH_OUTSIDE_WORKSPACE` | Path resolved outside the allowed workspace. |
| `UNSUPPORTED_FORMAT` | File extension cannot be handled by the requested tool. |
| `OUTPUT_EXISTS` | Target exists and `overwrite` was not requested. |
| `FEATURE_DISABLED` | The requested capability is disabled in `features`. |
| `BACKEND_UNAVAILABLE` | Optional COM/LibreOffice/OCR backend is not installed or not detected. |
| `SCHEMA_VERSION_UNSUPPORTED` | Spec uses a schema version newer than the engine supports. |
| `INVALID_PDF` / `PDF_PASSWORD_REQUIRED` / `PDF_PASSWORD_INVALID` | PDF is malformed, encrypted without a password, or has a wrong password. |
| `ENGINE_BAD_OUTPUT` | The Python engine returned a value that cannot be serialized losslessly. |

## Testing and benchmarks

The repository ships ten test suites plus a benchmark suite. They use temporary directories and mock hosts; none starts DSH or writes to a profile.

| Suite | Purpose |
|---|---|
| `tests/test_engine.py` | End-to-end Python engine: all formats, templates, audit, PDF, native, batch, copy. |
| `tests/test_fuzz.py` | Malformed JSON, path escapes, unknown actions, nested batches, stable error codes. |
| `tests/test_com_cleanup.py` | Optional COM process reclamation test; skips without Office/pywin32. |
| `tests/test_plugin.mjs` | Node plugin layer with a mock host and mock `officeKit`. |
| `tests/test_plugin_config.mjs` | Module/feature normalization and type trimming. |
| `tests/test_compatibility.mjs` | Real host schema validation, tool-name collisions, isolation static checks. |
| `tests/test_client.mjs` | Client bundle manifest, lazy-CJS shape, slot registration. |
| `tests/test_client_host.mjs` | Drives the client bundle through the real `dsh-client-modules` client system. |
| `tests/test_dsh_profile.mjs` | One-shot install in an isolated `DSH_HOME`; verifies profile manifest and bundle declaration. |
| `tests/run_regression.py` | Runs every suite against `tests/regression_baseline.json`; `--quick` skips optional suites. |
| `benchmarks/run_benchmarks.py` | 12-case performance/robustness baseline; thresholds in `benchmarks/baseline.json`. |

### Local verification

Set `PYTHON` to a Python 3.12 executable with `python-docx`, `openpyxl`, and `python-pptx` installed. If you use the DSH bundled runtime:

```bash
export DSH_HOME="${DSH_HOME:-$HOME/.dsh}"
export PYTHON="$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe"
```

Windows PowerShell:

```powershell
$env:DSH_HOME = "$env:USERPROFILE\.dsh"
$env:PYTHON = "$env:DSH_HOME\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
```

Run everything:

```bash
"$PYTHON" tests/run_regression.py
"$PYTHON" tests/run_regression.py --quick
```

Run individual suites:

```bash
"$PYTHON" tests/test_engine.py
"$PYTHON" tests/test_fuzz.py
"$PYTHON" tests/test_com_cleanup.py
node tests/test_plugin.mjs
node tests/test_plugin_config.mjs
node tests/test_compatibility.mjs
node tests/test_client.mjs
node tests/test_client_host.mjs
"$PYTHON" benchmarks/run_benchmarks.py
node tests/test_dsh_profile.mjs
```

`tests/test_client_host.mjs` and `tests/test_dsh_profile.mjs` need a DSH installation. Point them at it with `DSH_APP_ROOT` (the directory containing `node_modules/@deepseek-ai/dsh`). They skip cleanly when the host is unavailable.

## Compatibility and limitations

- **DSH**: tested against `0.2.2-rc.5`; package metadata declares `engines.dsh >= 0.2.2-rc.5`.
- **Node**: tested with Node 20+; the DSH host runs Node 26.
- **Python**: tested with CPython 3.12.14; `python-docx 1.2.0`, `openpyxl 3.1.5`, `python-pptx 1.0.2`.
- **Office COM**: Windows + Microsoft Office + pywin32 only. Without it, `officeforge_native` returns `BACKEND_UNAVAILABLE`; the core file-level tools still work.
- **LibreOffice**: optional; only used for legacy conversion when present.
- **OCR**: optional; requires PyMuPDF, pytesseract, and the Tesseract executable + language data.
- **Preview**: uses the host `officeKit.renderImages`; exact pixel sizes and font availability depend on the host and installed fonts.
- **Native features**: native PivotTable/PivotChart/Slicer and native SmartArt require COM; the file-level Python path provides pivot summaries and SmartArt-style editable shapes, not full native objects.
- **Word watermarks**: VML watermarks are visible in Word; some LibreOffice previews may not render them.
- **Macros**: macros are not executed; editing `.xlsm` preserves the VBA container where openpyxl supports it.

## Troubleshooting

| Symptom | Check |
|---|---|
| `officeforge_doctor` says `BACKEND_UNAVAILABLE` | The optional backend is not installed. Core tools remain usable. |
| `INPUT_NOT_FOUND` for a relative path | Paths are resolved relative to the session workspace; use a path inside it. |
| `PATH_OUTSIDE_WORKSPACE` | The target escapes the workspace; write inside it or use an allowed host path. |
| `FEATURE_DISABLED` | A feature is turned off in `features`; enable it in the bundle config. |
| `value is not lossless JSON` | A tool returned an absent optional field; update to a version that normalizes results with `lossless()`. |
| Python import errors | Run `officeforge_doctor`; point `pythonPath` or `OFFICEFORGE_PYTHON` at a Python with `python-docx`, `openpyxl`, and `python-pptx`. |
| COM output fails | Verify Windows Office is installed, pywin32 is available, and no protected-view dialog blocks the file. |

## Repository layout

```
office-forge/
  package.json                 # DSH bundle and npm metadata
  cordis.patch.yml             # inserts dsh-office-forge into a profile
  icon.svg
  locale/
    zh.json
    en.json
  resources/
    themes.json                # 31 built-in theme presets
    features.json              # module/feature registry
    layouts.json               # 38 PowerPoint layout presets
    templates.json             # 10 deck templates
  lib/
    index.js                   # tool registration, path policy, Python transport, officeKit bridge
    client.js                  # read-only plugin-page card and environment check
    office_engine.py           # JSON stdin/stdout engine entrypoint
    office_common.py           # paths, atomic writes, placeholders, validation
    office_pdf.py              # PDF text extraction
    office_pdf_tools.py        # PDF file operations
    office_backends.py         # optional COM / LibreOffice backends
    office_com_worker.py       # isolated COM worker
    office_migrations.py       # spec schema migration
    office_docx.py             # Word
    office_xlsx.py             # Excel
    office_pptx.py             # PowerPoint
    vendor/pypdf/              # vendored pypdf runtime
  benchmarks/
    run_benchmarks.py
    baseline.json
  tests/
    test_engine.py
    test_fuzz.py
    test_com_cleanup.py
    test_plugin.mjs
    test_plugin_config.mjs
    test_compatibility.mjs
    test_client.mjs
    test_client_host.mjs
    test_dsh_profile.mjs
    run_regression.py
    regression_baseline.json
  docs/
    DESIGN.md
    COMPATIBILITY.md
    COMPATIBILITY-MATRIX.md
    COVERAGE.md
    CHANGELOG.md
    BENCHMARKS.md
  README.md
  README.zh.md
  LICENSE
```

## Development

The plugin is plain ESM + Python 3.12; there is no build step.

```bash
git clone https://github.com/siyu0209/dsh-office-forge.git
cd dsh-office-forge
npm install --ignore-scripts
```

Run the verification commands in the table above before opening a pull request. Keep feature changes covered by `tests/test_engine.py` and update `tests/regression_baseline.json` only when the new assertion floor is intentional.

## Versioning

Semantic versioning: patch versions fix bugs and compatibility; minor versions add capabilities; major versions may change document spec semantics. `docs/CHANGELOG.md` records every release and the migration rules.

## License

MIT. See [LICENSE](LICENSE).
