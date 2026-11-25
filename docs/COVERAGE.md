# Office Forge Coverage & Regression Baseline

This project treats coverage as a capability matrix plus deterministic test
counts, not as a single line-coverage number. The bundled Python runtime has no
`coverage` package; `tests/run_regression.py` can still be used in any Python
environment that installs one.

## Capability coverage

| Capability | Engine tests | Node tests | Backend requirement |
|---|---|---|---|
| Word read/create/edit/template, images, hyperlinks, TOC, captions, styles, fields, comments, tracked changes, mail merge, footnotes, bookmarks, cross references, document audit | `test_docx`, `test_docx_features`, `test_docx_v03`, `test_docx_v05_plus`, `test_docx_v08`, `test_docx_deepening`, `test_docx_audit` | `test_plugin.mjs` | python-docx |
| Excel read/create/edit/template, charts, combo/secondary-axis charts, CSV/TSV/JSON import, array formulas, conditional formatting, validation, comments, merged cells, images, formula columns, data tools, audit, pivot summaries | `test_xlsx`, `test_xlsx_features`, `test_xlsx_v03`, `test_xlsx_v05_plus`, `test_xlsx_audit`, `test_xlsx_v10` | `test_plugin.mjs` | openpyxl |
| PowerPoint read/create/edit/template, themes, deck check, transitions, shapes, text boxes, connectors, animations, SmartArt-style diagrams, dual engines, aesthetic score, beautify, HTML deck | `test_pptx`, `test_pptx_features`, `test_pptx_v03`, `test_pptx_v05_plus`, `test_pptx_audit`, `test_ppt_v015` | `test_plugin.mjs` | python-pptx |
| Deck IR / Adapter Registry / Template Pack (listing, validation) / rendered-slide verification / JSON+HTML verification reports / deck audit-repair loop | `test_deck_v017`, `test_deck_v018`, `test_deck_storyboard`, `test_deck_v100`, `test_deck_zbsp`, `test_deck_anon1009`, `test_deck_guizang`, `test_deck_composer`, `test_deck_template_pack`, `test_deck_builtin_packs`, `test_deck_template_validation`, `test_deck_capabilities_loss_audit`, `test_deck_import_design`, `test_deck_v306`, `test_deck_adapter_auto`, `test_deck_edit`, `test_deck_verify` | `test_plugin.mjs` (`officeforge_deck`) | python-pptx + optional dsh-ppt + bundled Anon1009 audit/beautify + host officeKit for verify |
| PDF read | `test_pdf_v05` | — | pypdf/fitz/pdfplumber or fallback |
| PDF merge/split/rotate/watermark/pageNumbers/formFields/fillForm/extractImages/ocr/encrypt/decrypt | `test_pdf_tools`, `test_pdf_tools_v090` | `test_plugin.mjs` | vendored pypdf; OCR optional |
| Native COM export/updateFields/recalculate/pivot/SmartArt | manual smoke tests; `test_native_update_fields` | — | Windows + Office + pywin32; `BACKEND_UNAVAILABLE` otherwise |
| Legacy conversion | `test_convert_v05` | `test_plugin.mjs` mock | COM/LibreOffice or explicit `BACKEND_UNAVAILABLE` |
| Preview / recalculate | `test_plugin.mjs` mock officeKit | `test_plugin.mjs` mock officeKit | host officeKit |
| Schema migration | `test_schema_migration_v05` | — | none |
| Runtime diagnostics | `test_plugin.mjs` (`officeforge_doctor`) | `test_plugin.mjs` | none |
| Client bundle (plugin-page status card) | — | `test_client.mjs`, `test_client_host.mjs` | `@deepseek-ai/dsh-client-modules` |
| One-shot profile install | — | `test_dsh_profile.mjs` | DSH CLI + isolated `DSH_HOME` |
| Isolation / compatibility | `test_compatibility.mjs` | — | none |
| Malformed input handling | `test_fuzz.py` | — | none |
| COM process reclamation | `test_com_cleanup.py` | — | Windows + Office + pywin32 |

## Baseline (2026-11-24)

| Suite | PASS assertions | Expected status | Needs |
|---|---:|---|---|
| `tests/test_engine.py` | 334 | `ALL TESTS PASSED` | bundled Python |
| `benchmarks/run_benchmarks.py` | 13 | `BENCHMARK PASSED` | bundled Python |
| `tests/test_fuzz.py` | 45 | `FUZZ TESTS PASSED` | bundled Python |
| `tests/test_plugin.mjs` | 72 | `PLUGIN TESTS PASSED` | test junctions |
| `tests/test_ppt_bridges.mjs` | 140 | `PPT BRIDGE TESTS PASSED` | dsh-ppt + dsh-ppt-studio |
| `tests/test_plugin_config.mjs` | 14 | `PLUGIN CONFIG TESTS PASSED` | test junctions |
| `tests/test_compatibility.mjs` | 54 | `COMPATIBILITY TESTS PASSED` | test junctions |
| `tests/test_client.mjs` | — | `CLIENT UI TESTS PASSED` | none |
| `tests/test_client_host.mjs` | — | `CLIENT HOST MACHINERY TESTS PASSED` | DSH client modules |
| `tests/test_com_cleanup.py` | 3 | `COM CLEANUP TESTS PASSED` | Office + pywin32 |
| `tests/test_dsh_profile.mjs` | — | `DSH PROFILE TESTS PASSED` | DSH CLI, isolated `DSH_HOME` |

`python tests/run_regression.py` runs every suite above and enforces both the
assertion floor and each suite's completion marker, so a suite that silently
reports zero assertions fails instead of passing. `--quick` skips the two
suites that need Microsoft Office and the DSH CLI. A count change must update
this file, `docs/DESIGN.md` and `tests/regression_baseline.json` together.
