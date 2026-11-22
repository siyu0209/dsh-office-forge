# Office Forge

Office Forge is an Office document bundle for DeepSeek Harness (DSH). It provides thirteen independent `officeforge_*` tools for reading, creating, editing, template-filling, auditing, batch-processing, PDF merge/split/rotate/watermark, legacy conversion, visual preview, formula recalculation, PDF export, and optional native Microsoft Office COM operations across Word, Excel, PowerPoint, and PDF files. All document logic runs in short-lived Python subprocesses; the Node layer only handles workspace authorization, argument validation, subprocess transport, and optional export through the host `officeKit` service.

> This repository contains the plugin source and offline tests. Installing into a DSH profile is opt-in and described under *Installation*.

## Tools

| Tool | Purpose |
|---|---|
| `officeforge_read` | Read `.docx` / `.xlsx` / `.pptx` / `.pdf` in `content`, `outline`, or `search` mode for Office files; render as markdown, text, or JSON. PDF supports text/markdown extraction. |
| `officeforge_create` | Create a document from a declarative spec: docx uses `title/subtitle/sections`, xlsx uses `sheets` (or `import` from CSV/TSV/JSON), pptx uses `slideSize/slides`. Excel chart specs support `combo` with primary/secondary axes. |
| `officeforge_edit` | Apply bounded operations: docx `replace/append/addImage/addToc/addCaption/setStyle/updateFields/addComment/trackRevision/mailMerge/addFootnote/addBookmark/addCrossReference/addCitation/addBibliography/setColumns/addWatermark/addContentControl/saveAsTemplate/setCore`; xlsx `setCells/appendRows/replace/createSheet/deleteSheet/renameSheet/setColumnWidths/setFreeze/setAutoFilter/addChart/addConditionalFormatting/addDataValidation/addComment/mergeCells/addImage/addFormulaColumn/sortRange/removeDuplicates/fillBlanks/cleanRange/pivotSummary/addTable/protectSheet/protectWorkbook`; pptx `replace/appendSlides/setNotes/setNotesBatch/addSection/listSections/deleteSlide/moveSlide/setTransition/addShape/addTextBox/addConnector/addAnimation/addSmartArt/addComment/listComments/addMedia/applyMaster/applyTheme/setCore`. A sibling backup is made by default. |
| `officeforge_template` | Fill templates with JSON or `.json/.csv/.tsv` data: `{{key}}` placeholders, plus one-level `{{#each items}}...{{/each}}` table/sheet row loops for Word and Excel. |
| `officeforge_audit` | Read-only audit: Word `health` (heading jumps, duplicate headings, orphan headings, unresolved placeholders, long paragraphs, empty paragraph runs) and `profile`; Excel `health` (formula errors, inconsistent formula patterns, hardcoded values in formula columns, blank row breaks, circular references), `profile`, `insight`, `diff`; PowerPoint `deck` quality checks, `profile`, `themes`. |
| `officeforge_doctor` | Read-only runtime diagnostics: Python with python-docx/openpyxl/python-pptx, bundled PDF engine, COM/LibreOffice backends, Office application registration, `officeKit` availability, client bundle, and the active modules/features configuration. Every check carries required/optional, status, and a remedy when missing. Run it after installation to see which capabilities are available. |
| `officeforge_convert` | Convert legacy `.doc/.docm/.xls/.xlsm/.ppt/.pptm` through an optional COM/LibreOffice backend, or modern/PDF content through the host `officeKit`. |
| `officeforge_pdf` | File-level PDF operations: `merge`, `split`, `rotate`, `watermark`, `pageNumbers`, `formFields`, `fillForm`, `extractImages`, `ocr`, `encrypt`, `decrypt`, powered by vendored pypdf with no external PDF service. |
| `officeforge_native` | Optional native Microsoft Office COM operations: `convert`, `exportPdf`, `acceptRevisions`, `recalculate`, `pivot` (native PivotTable), and `smartArt` (native SmartArt); requires Windows + Office + pywin32. |
| `officeforge_preview` | Render pages, slides, or worksheet areas to PNG through `officeKit.renderImages` without launching an Office UI. |
| `officeforge_recalculate` | Recalculate formulas through `officeKit.recalculate` and save a refreshed workbook instead of approximating formula text. |
| `officeforge_batch` | Run several independent create/edit/template jobs in one Python call. Each item succeeds or fails independently; nested batches are rejected. |
| `officeforge_export` | Convert through the host `officeKit` service to PDF or another supported format; optionally copy the result back into the workspace. |

## Module and feature switches

Office Forge lets you choose document families and advanced capabilities in the DSH bundle config. Every switch defaults to `true`, so leaving the config empty enables all v0.6 capabilities; `resources/features.json` and `resources/themes.json` are externalized.

```json
{
  "modules": { "word": true, "excel": true, "ppt": true, "pdf": true },
  "features": {
    "word": { "images": true, "hyperlinks": true, "pageBreaks": true, "toc": true, "captions": true, "styles": true, "fields": true, "comments": true, "revisions": true, "mailMerge": true, "footnotes": true, "bookmarks": true, "crossReferences": true },
    "excel": { "charts": true, "comboCharts": true, "conditionalFormatting": true, "dataValidation": true, "comments": true, "images": true, "mergedCells": true, "formulas": true, "arrayFormulas": true, "dataTools": true, "audit": true, "pivots": true, "tables": true, "protection": true, "import": true },
    "ppt": { "charts": true, "images": true, "transitions": true, "shapes": true, "themes": true, "audit": true, "animations": true, "smartart": true, "notes": true, "sections": true }
  }
}
```

- Disabling `modules.word` / `modules.excel` / `modules.ppt` removes that document family and its tool entry points; disabling `modules.pdf` removes `officeforge_export`.
- Disabling a `features` item makes the engine return `FEATURE_DISABLED` if that capability is requested; there is no silent downgrade or partial output.
- Advanced capabilities: Word images/hyperlinks/page breaks/table of contents/captions/styles/field refresh/comments/tracked changes/mail merge/footnotes/bookmarks/cross-references; Excel charts/combo secondary-axis charts/conditional formatting/data validation/comments/images/merged cells/formula columns/array formulas/sorting/deduplication/blank fill/text cleanup/formula health audit/pivot summaries/CSV/TSV/JSON import/native table objects/sheet and workbook protection; PowerPoint charts/images/transitions/shapes/text boxes/connectors/theme presets/deck quality checks/animations/SmartArt-style diagrams/batch speaker notes/sections.
- Generated files carry neutral core properties: `author`/`lastModifiedBy` default to `Office Forge`, `created`/`modified` use the current UTC time, and `revision` starts at 1. Pass `properties: {author, title, subject, keywords, comments, category}` to any create spec to override; docx also keeps its historical top-level `author` spelling.
- Optional real backends: `enableBackends` defaults to `true`; `officeforge_pdf` always works with vendored pypdf, while `officeforge_convert`/`officeforge_preview`/`officeforge_recalculate` use the host `officeKit` or a local COM/LibreOffice installation. `officeforge_native` requires Windows + Microsoft Office + pywin32; when unavailable it returns `BACKEND_UNAVAILABLE` and never blocks the native tools.

## Quick examples

Create a Word document:

```json
{
  "path": "reports/weekly.docx",
  "type": "docx",
  "spec": {
    "title": "Weekly Report",
    "subtitle": "2026-11-21",
    "sections": [
      { "heading": "Progress", "bullets": ["Finished module A", "Validated flow B"] },
      { "heading": "Next week", "numbered": ["Fix C", "Ship D"] }
    ]
  }
}
```

Create an Excel workbook with a chart, conditional formatting, and an image:

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
        "charts": [{ "type": "bar", "dataRange": "B1:B4", "categoriesRange": "A2:A4", "anchor": "E2", "title": "Revenue" }],
        "conditionalFormatting": [{ "range": "B2:B4", "type": "cellIs", "operator": "greaterThan", "formula": "150", "fill": "FFC7CE" }],
        "images": [{ "path": "logo.png", "anchor": "G2", "width": 120, "height": 40 }]
      }
    ]
  }
}
```

Fill a Word template with a table loop:

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

Export a PDF:

```json
{
  "path": "reports/weekly.docx",
  "output": "out/weekly.pdf",
  "format": "pdf"
}
```

Audit Excel formula health:

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "mode": "health"
}
```

Check a PowerPoint deck and list themes:

```json
{
  "path": "decks/quarterly.pptx",
  "type": "pptx",
  "mode": "deck"
}
```

```json
{
  "type": "pptx",
  "mode": "themes"
}
```

## Configuration

All fields have defaults and can be overridden in the bundle config:

| Field | Default | Description |
|---|---|---|
| `pythonPath` | auto-detect | Optional Python executable with `python-docx`, `openpyxl`, and `python-pptx` installed. |
| `engineTimeoutMs` | `120000` | Timeout for one Python engine call, in milliseconds. |
| `maxBatchItems` | `20` | Maximum items in one `officeforge_batch` call. |
| `enableExport` | `true` | Register `officeforge_export`. |
| `enableBatch` | `true` | Register `officeforge_batch`. |
| `enableBackends` | `true` | Register `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate`; missing backends fail explicitly. |
| `modules` | all `true` | `word` / `excel` / `ppt` / `pdf` document-family switches. |
| `features` | all `true` | Fine-grained Word / Excel / PowerPoint capability switches. |

When installing into a profile, override the settings in that bundle's `cordis.patch.yml` entry, for example to enable Word and Excel only, and disable Excel charts:

```yaml
- insert:
    - id: dsh-office-forge
      name: dsh-office-forge
      config:
        modules: { word: true, excel: true, ppt: false, pdf: true }
        features:
          excel: { charts: false, dataTools: true }
```

Python discovery order: `pythonPath` → `OFFICEFORGE_PYTHON` → `DSH_PYTHON` → `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe` → scan `$DSH_HOME/dsh-runtimes/*/dependencies/python` → `PATH`. The bundled Python 3.12.14 on this machine already includes all required packages.

## Boundaries and design guarantees

- **Workspace-only writes**: every input and output path passes both `ctx.fs.resolve` / `ctx.fs.contains` and a lexical relative-path check. Escapes fail with `PATH_OUTSIDE_WORKSPACE`.
- **Atomic writes**: new files are written to a same-directory temporary file and `os.replace`d. Existing files are refused unless `overwrite` is explicit.
- **Bounded execution**: 50 MiB input, 100 MiB output, 200,000 text characters, 5,000 × 300 Excel cells per sheet, 300 slides, 20 batch items.
- **No recursion**: each tool calls exactly one Python action; tools never call each other; `officeforge_batch` rejects nested batches.
- **No interference with the armor-piercing plugin**: no prompt variables, no LLM stream hooks, no tool middleware. `@adwmc/helm-d` hcot / attack-scheduler / advisory behavior is untouched.
- **No built-in tool collisions**: every tool starts with `officeforge_`, so built-in `office_render` / `office_convert` / `office_recalculate` remain unchanged.
- **Per-tool isolation**: tools are registered one by one; a failed registration logs a warning and never aborts activation of the rest of the plugin or the profile.
- **Compatibility hardening (v0.6.1)**: `officeforge_export` uses the root session as the `officeKit` output owner, checks the source version around conversion, and rejects unmappable `officeKit` paths on remote workspaces; invalid profile config falls back to defaults instead of blocking activation; Excel chart data uses a unique hidden worksheet and never reuses or overwrites a user sheet.
- **Plugin-page status card and environment check (v0.14.0)**: `lib/client.js` registers exactly one read-only card in `plugins.detail.section` showing the version, the tool inventory, and a **Check environment** button. The button calls the read-only `GET /officeforge/doctor` route and renders health/summary plus per-check status and remedy. It reads and writes no configuration, performs no document work, and registers no other slot.

## Compatibility and isolation (v0.14.0)

- Injects only `tools`, `fs`, and `workingDirectory`; it does not register `systemPrompt`, does not take over the LLM stream, and does not register middleware/hooks, so `@adwmc/helm-d` hcot, attack-scheduler, and advisory behavior remain untouched.
- The client bundle injects only `slots` and registers only `plugins.detail.section` (id `officeforge.status`, order 40). It never registers `plugins.bundle.config` and never mutates configuration, so the plugin page cannot drift away from the tools that were actually registered.
- All thirteen tool names start with `officeforge_`; they do not collide with built-in `office_render` / `office_convert` / `office_recalculate`, `run_code`, or common `word_*` / `excel_*` / `ppt_*` / `univer_*` tools.
- `tests/test_compatibility.mjs` validates all thirteen output schemas with the real `assertSupportedJsonSchema`, verifies zero tools when all modules are off, checks that the type enum narrows to Word-only, and confirms no host tool is shadowed.
- Every document operation runs in a short-lived Python subprocess; no process environment, global configuration, DSH process, or profile is modified.
- A failed tool registration only logs a warning; invalid config falls back to defaults; batch errors are isolated per item. When several plugins are installed, callers should still avoid concurrent writes to the same file; Office Forge does not provide a cross-plugin file lock.

## Layout

```
office-forge/
  package.json           # DSH bundle and npm metadata
  cordis.patch.yml       # inserts dsh-office-forge into a profile
  icon.svg
  locale/
    zh.json              # { "meta": { "title", "description" } }
    en.json
  resources/
    themes.json          # 31 externalized themes (24 from dsh-ppt + 7 legacy themes)
    features.json        # module/feature registry
  lib/
    index.js             # DSH tool registration, path policy, Python transport, officeKit export/preview/recalculate
    client.js            # plugin-page read-only status card + environment check (plugins.detail.section)
    office_engine.py     # JSON stdin/stdout engine entrypoint
    office_common.py     # paths, atomic writes, placeholders, data files, shared validation
    office_pdf.py        # PDF text extraction (pypdf/fitz/pdfplumber/fallback)
    office_backends.py   # optional COM / LibreOffice legacy conversion
    office_migrations.py # spec schema migration
    office_docx.py       # Word read/create/edit/template
    office_xlsx.py       # Excel read/create/edit/template
    office_pptx.py       # PowerPoint read/create/edit/template
  benchmarks/
    run_benchmarks.py   # performance/robustness baseline (12 real engine cases)
    baseline.json       # per-case maxSeconds / maxPeakMB thresholds
  tests/
    test_engine.py       # Python engine end-to-end tests (temporary workspace)
    test_fuzz.py         # Malformed input and path-escape robustness
    test_com_cleanup.py  # COM process reclamation (needs Office + pywin32)
    run_regression.py    # Every suite against the recorded baseline
    regression_baseline.json # Per-suite assertion floors
    test_plugin.mjs      # Node plugin-layer mock-host tests
    test_plugin_config.mjs # module/feature switch tests
    test_compatibility.mjs # host contract, tool-name collision and isolation checks
    test_client.mjs      # client bundle shape and slot registration
    test_client_host.mjs # drives the client bundle with the real dsh-client-modules
    test_dsh_profile.mjs # one-shot install verification in an isolated DSH_HOME
  docs/
    BENCHMARKS.md        # performance and robustness baseline
    DESIGN.md            # research sources, architecture, recursion constraints, compatibility
    COMPATIBILITY.md     # v0.14.0 host contract, tool namespaces and isolation audit
    COMPATIBILITY-MATRIX.md # plugin capability/format/backend compatibility matrix
    CHANGELOG.md         # version changes and migration notes
    RESEARCH-v0.4.md     # v0.4 candidate-plugin survey: absorbed formula audit + PPT theme/deck checks; no COM/exceljs/LibreOffice
  README.md
  README.zh.md
  LICENSE
```

## Local verification

From `office-forge/`:

```bash
# Everything at once: assertion floors plus each suite's completion marker
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/run_regression.py

# Skip the two suites that need Microsoft Office and the DSH CLI
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/run_regression.py --quick

# Individually
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/test_engine.py
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/test_fuzz.py
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/test_com_cleanup.py
node tests/test_plugin.mjs
node tests/test_plugin_config.mjs
node tests/test_compatibility.mjs
node tests/test_client.mjs
node tests/test_client_host.mjs
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" benchmarks/run_benchmarks.py
node tests/test_dsh_profile.mjs
```

All ten suites (including the benchmark suite) use temporary directories and a mock host. They do not start DSH and do not write to any profile. `test_dsh_profile.mjs` creates a one-shot `DSH_HOME` under the system temp directory, installs the local tarball into a temporary profile, verifies the result, and deletes it again. `test_com_cleanup.py` really drives Microsoft Office to export a PDF and compares `WINWORD/EXCEL/POWERPNT` process counts before and after, so a leaked COM instance fails the run; without Office or pywin32 it skips instead.

## Installation

Add the packed tarball (or this directory) as a local bundle to the target profile and enable `dsh-office-forge`. On the machine that produced this build the plugin is installed in the local desktop profile; `officeforge_doctor` reports the live configuration.

## License

MIT
