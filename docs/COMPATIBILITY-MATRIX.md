# Office Forge Compatibility Matrix

| Area | Native core | Optional backend | Notes |
|---|---|---|---|
| `.docx` read/create/edit/template | python-docx 1.2.0 | — | images, hyperlinks, TOC, captions, styles, fields, comments, tracked changes, mail merge |
| `.xlsx` read/create/edit/template | openpyxl 3.1.5 | officeKit.recalculate | charts, conditional formatting, data validation, comments, merged cells, images, audit, pivot summaries |
| `.pptx` read/create/edit/template | python-pptx 1.0.2 | officeKit.renderImages | 31 external themes, deck check, transitions, shapes, text boxes, connectors, animations, SmartArt-style diagrams |
| `.pdf` read | pypdf/fitz/pdfplumber if present; built-in fallback | — | read-only; malformed/encrypted PDFs may be unavailable |
| `.pdf` merge/split/rotate/watermark | vendored pypdf | — | `officeforge_pdf`; no external PDF service |
| Native open-document operations | — | Windows COM worker | convert/exportPdf/acceptRevisions/recalculate/native pivot/native SmartArt |
| `.doc/.docm/.xls/.xlsm/.ppt/.pptm` | — | Windows COM or LibreOffice `soffice` | `BACKEND_UNAVAILABLE` when neither exists |
| Export PDF/office formats | — | `officeKit.convertTo` | host service required |
| Visual preview | — | `officeKit.renderImages` | PNG, page/sheet/range bounds |
| Formula recalculation | — | `officeKit.recalculate` | real spreadsheet engine required |
| Legacy `.doc/.xls/.ppt` on this machine | unsupported without backend | system Python has win32com but bundled Python does not | install Office/COM or LibreOffice to enable |
| Runtime diagnostics | `officeforge_doctor` | — | Python, PDF vendor, COM/LibreOffice, officeKit and active feature switches |

## Isolation contract

- Tool prefix `officeforge_`; no built-in `office_*` or `run_code` collisions.
- No `systemPrompt`, LLM stream hook, tool middleware, hcot, attack-scheduler, or advisory registration.
- All writes stay inside the session workspace; atomic write + optional backup.
- Every optional backend failure is explicit and does not affect Native tools.
