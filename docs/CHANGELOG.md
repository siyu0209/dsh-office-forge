## 3.5.7 — 2026-11-25

- **Public web research is allowed**: the PPT intake contract now explicitly permits searching authoritative public sources when key content is missing, provided the assistant announces what it is checking, cites source names/URLs plus access dates, never presents public facts as user-provided data, and never uploads local/private files. Internal company data still requires user files.
- **Skills**: `officeforge-ppt` and `officeforge` gain a `2.6 联网查找资料` section with the same sourcing, citation and privacy rules.
- **Tests**: the system-prompt assertion now also checks the public-web-research clause.

## 3.5.6 — 2026-11-25

- **Progressive PPT intake**: the mandatory `officeforge:ppt-intake` prompt now uses layered questioning instead of one fixed list. First round asks only the broad direction (audience/scenario, new vs existing file, `.pptx` vs `.html`); after the answer it restates the direction and plans the next round from those answers (content source, must-have sections/data/page count, brand/template, edit constraints). Rounds are unlimited: each round is planned from the previous answer and stops when the brief is sufficient; visual details still move to a 2–3 page sample.
- **Skills**: `officeforge-ppt` and `officeforge` carry the same layered intake wording; `whenToUse` metadata now describes the broad-first/plan-next behavior.
- **Tests**: `tests/test_plugin.mjs` asserts the system prompt contains `progressive`, `First round` and `next round`; plugin suite remains 72 PASS.

## 3.5.5 — 2026-11-25

- **Mandatory PPT intake system prompt**: the plugin now contributes an `officeforge:ppt-intake` system-prompt section. Every new model step sees that a bare "make me a PPT" must first confirm 1–3 facts (audience/scenario, new vs existing file, content source and must-have data, `.pptx` vs `.html`, brand/template) before any `officeforge_*` PPT generation or edit; it must not invent files/numbers. Visual choices move to a 2–3 page sample after the brief. This is additive and does not touch other plugins' prompt sections or the `@adwmc/helm-d` hook chain.
- **Skill routing metadata**: `officeforge` and `officeforge-ppt` now declare `whenToUse`, and `lib/skill.js` forwards it to `ctx.skills.register()` so the loader catalog can surface the intake requirement even before the skill body is opened.
- **Tests**: `tests/test_plugin.mjs` adds a system-prompt section assertion and a `whenToUse`-aware skill mock; plugin suite 72 PASS. No engine or document behavior changed.

## 3.5.1 — 2026-11-25

- **Guizang is now a real HTML engine**: `officeforge_deck` and `officeforge_web` accept `engine: "guizang"` and render through the bundled `guizang-ppt-skill` `template-swiss.html` (CSS, motion runtime, presenter shell and speaker notes included), instead of using the old `guizang-*` native-PPTX palette approximation.
- **Skill intake rules**: `skills/officeforge-ppt/SKILL.md` asks 1–3 key questions before choosing an engine, routes Guizang/magazine/Swiss/web decks to real Guizang HTML, and forbids pretending `engine: "guizang"` native PPTX is full Guizang.
- **Bundled skill set**: `guizang-ppt-skill`, `guizang-page-skill`, `guizang-social-card-skill` and `humanizer-zh` are packaged under `skills/` and registered by `lib/skill.js`.
- **Engine tests**: `test_deck_guizang` now proves a real `template-swiss.html` output with slides and speaker notes; all engine tests pass.

## 3.4.1 — 2026-11-25

- **Bundled PPT intake skill**: adds `skills/officeforge-ppt/SKILL.md` and `lib/skill.js`; the plugin registers the `officeforge-ppt` runtime skill and exposes it through `exports["./skills/*"]`.
- **Adaptive, non-form intake**: asks only the 1–2 missing facts that would materially change the deck; visual choices move to a 2–3 page sample; no fixed questionnaire.
- **Packaging**: `skills` is included in `files`; plugin tests now 66 with skill registration assertions.

## 3.4.0 — 2026-11-25

- **Design-aware deck editing**: `office_deck_ir.apply_deck_operations` and `office_deck.edit_deck` support `setTheme`, `setTitle`, `addSlide`, `removeSlide`, `replaceSlide`, `appendBlock`, `replaceBlock`, and `setNotes` against an imported `.pptx` or spec deck, then re-render through the normal adapter pipeline.
- **Node surface**: `officeforge_deck` gains `operation: "edit"` with `output` and `operations` parameters; `test_deck_edit` covers import→edit→render, spec-only base, appended blocks, notes, and invalid index rejection.
- Engine tests 334, including `test_deck_edit` (7 assertions).

## 3.3.0 — 2026-11-25

- **Advanced Office objects**: `office_deck_ir` and legacy renderer now handle SmartArt, transitions, animations, media lists and comments; `media` can be an array of audio/video objects.
- **Reverse import fidelity**: `office_ppt_import` preserves shape fill/line/rotation, copies embedded media into the workspace, emits media/smartart blocks, records transitions and animation presence, and keeps comments.
- **Design adapter full path**: `AnonPptAdapter.create` and `.edit` now render a temporary legacy deck then run Anon1009 beautify instead of returning `ADAPTER_UNSUPPORTED`; loss audit includes `smartart`, `animations` and `morph` for the correct legacy ownership.
- **Web deck coverage**: `office_webdeck` renders charts, shapes, media and comments in addition to text/tables/images.
- Engine tests 327, including `test_deck_v306` for anon create, advanced Office object render, import counts and Deck IR roundtrip.

## 3.2.0 — 2026-11-24

- **zbsph full schema**: table headers, line/bar/pie chart options, shapes/paths, lines/arrows, notes, solid/image backgrounds, safe-area/overlap declarations; Anon1009's four built-in templates (`business-blue`, `academic-white`, `tech-dark`, `pitch-bold`) are resolved from their `deck.yaml` themes.
- **Full catalog verification**: added `tests/test_ppt_bridges.mjs`; real render checks cover 24 themes, 8 master presets, 20 scenario templates, 82 layouts and 4 Vate templates — 140 PASS.
- **Studio bridge fix**: error responses are emitted once (no duplicated JSON), so failed renders produce a stable single error object.
- `office_deck_registry` advertises zbsph safeArea/overlap/tableHeader/chartOptions features so loss audit and auto-selection reflect the actual schema.

## 3.1.0 — 2026-11-24

- **Unified plugin-page UI**: `lib/client.js` now has both environment check and capabilities panel actions; `GET /officeforge/capabilities` exposes dsh-ppt counts and the design-token catalog to the plugin detail page.
- **Offline and release packaging**: `scripts/build_offline.py` builds a cross-machine zip with install/check scripts; `.github/workflows/release.yml` builds and attaches release assets; `docs/PUBLISHING.md` documents npm, GitHub, awesome-dsh-plugin and dsh-plugin-catalog paths.
- **Hardening**: fuzz coverage extended to deck import/verify/lossAudit/render malformed inputs (45 cases), added a deck-import benchmark (13 cases), and updated the regression baseline.

## 3.0.0 — 2026-11-24

- **Design-token catalog**: `officeforge_deck operation: "designCatalog"` lists the bundled Anon1009 themes, Guizang palettes, Composer/Office themes and their families; `capabilities` also includes the same catalog.
- **Reverse import**: `officeforge_deck operation: "import"` parses an existing `.pptx` into Deck IR and returns text/table/chart/image/shape/comment/notes/media/transition fidelity counts plus explicit warnings for objects that cannot be rebuilt exactly.
- **Advanced Office objects**: Deck IR accepts shape/line/media/comment blocks; `create_pptx` now writes slide-level comments, media, transitions, animations, SmartArt and backgrounds instead of silently dropping them.
- Engine tests 317, including design-catalog counts, PPTX import roundtrip and comment preservation assertions.

## 2.0.0 — 2026-11-24

- **Full dsh-ppt catalog**: `resources/dsh_ppt_catalog.json` records 24 themes, 8 master presets, 20 scenario templates and 82 runtime layouts; `officeforge_deck operation: "capabilities"` exposes it together with the adapter matrix.
- **Studio masters/templates**: studio bridge accepts `master` (named preset, mapped to theme/brand/chrome) and `template` (scenario template id); raw legacy master objects explicitly fall back to legacy.
- **Static loss audit**: `officeforge_deck operation: "lossAudit"` returns selected adapter, required/supported/lost feature sets and `lossless`, with no file rendering.
- Engine tests 311, including dsh-ppt catalog counts, master/template rendering and loss audit assertions.
- **zbsph/PPTD**: shapes, lines, backgrounds, table/chart/image and notes remain on the native PPTD path; unsupported Office-only objects still fall back to legacy.

## 1.17.0 — 2026-11-24

- **zbsph shapes/background**: `office_ppt_zbsp_bridge.mjs` now maps basic `shapes` (`rect`/`roundRect`/`ellipse`/`triangle`/arrow/chevron/pentagon/hexagon/flowchart shapes), `lines`, and per-slide `background` into native PPTD elements. Unsupported media/comments/master/morph/animation/SmartArt still fall back to legacy.
- **Adapter selection**: `office_deck_registry` now treats `shapes`/`background` as zbsph-capable features and auto selects zbsph for shape/background decks when no legacy-only objects are present.
- Engine tests 304, including zbsph shape/line and background preservation plus auto-selection for shape/background decks.

## 1.16.0 — 2026-11-24

- **Designed PPT by default**: `pptxEngine` now defaults to `studio`; a plain `officeforge_create` PowerPoint uses the designed dsh-ppt renderer instead of the legacy path.
- **Anon1009 beautify by default**: new `beautifyEngine` setting defaults to `anon1009`; `officeforge_beautify` routes to the bundled Anon1009 aesthetic adapter, with `engine: "builtin"` retaining the original theme/normalize path.
- Adds `engine` parameter to `officeforge_beautify`; aliases friendly theme names (`forge`, `consulting`, `ocean`, etc.) to Anon1009 themes.
- Engine tests 301, including the Anon1009 beautify adapter test.

## 1.15.0 — 2026-11-24

- **Render auto-verify**: `officeforge_deck` `operation: "render"` accepts `verify: true`; after writing the `.pptx` it renders every slide through the host officeKit and reuses the same verifier, returning the result as `verification`.
- Adds `verifyOutputDir` to keep auto-verify PNGs separate from the deck output; `report`/`reportFormat` continue to work for both `verify` and auto-verify.
- Plugin tests 61.

## 1.14.0 — 2026-11-24

- **Verification reports**: `officeforge_deck` `operation: "verify"` accepts `report` and optional `reportFormat`; the rendered-slide result is written as JSON or HTML (extension selects the default format) and returned as a workspace-relative `report` path.
- HTML reports contain the score/valid status, checks, issues, and per-slide summary; JSON reports keep the full machine-readable result.
- Plugin tests 60, including JSON and HTML report assertions.

## 1.13.0 — 2026-11-24

- **Rendered-slide verification**: `officeforge_deck` adds `operation: "verify"`; it renders every slide to PNG through the host officeKit, then checks page count, blank pages, image dimensions, shape bounds, expected title coverage, and missing fonts.
- Results include `score`, `valid`, structured `checks`, per-slide summaries, and issue codes (`BLANK_RENDER`, `EMPTY_SLIDE`, `SHAPE_OUT_OF_BOUNDS`, `TITLE_MISSING`, `RENDER_PAGE_MISSING`, `RENDER_TOO_SMALL`).
- Engine tests 300, including clean and blank rendered-page verification.

## 1.12.0 — 2026-11-24

- **Automatic PPT adapter selection**: `officeforge_deck` with `engine: "auto"` now inspects Deck IR features (theme/template, table/chart/image, Office-only objects) and selects `legacy`, `studio`, `zbsph`, or `guizang`; the result includes `adapterDecision` with required/unsupported features and a human-readable reason.
- Explicit engine requests remain unchanged; unavailable external adapters still raise `ADAPTER_UNAVAILABLE` with an installation hint.
- Engine tests 294, including plain/designed/table/Office-only auto-selection cases.

## 1.11.0 — 2026-11-24

- **Template Pack validation**: `officeforge_deck` adds `operation: "validate"` with optional `pack`; it checks manifest fields, asset files, template ids/slides, and layout references without rendering.
- `validate_template_pack()` reports machine-readable `errors`/`warnings` and per-pack `themeCount/layoutCount/templateCount`; unknown pack ids return `TEMPLATE_PACK_NOT_FOUND`.
- Engine tests 290, including built-in pack validation and a deliberately broken external pack.

## 1.10.0 — 2026-11-24

- Word `updateFields` via COM refreshes fields and table of contents, then saves a DOCX copy.
- zbsph adapter discovery now checks pnpm-hoisted profile packages, not just the plugin-local `node_modules`.
- Engine tests 285.

## 1.9.1 — 2026-11-24

- 修复 Word 审计对文档标题的误报：标题后接更深层级标题属于正常大纲结构，不再记为 `orphan_heading`。
- 新增 `document title followed by chapter is not an orphan heading` 测试；engine 282 项 PASS。

## 1.9.0 — 2026-11-24

- **Storyboard Markdown**：`officeforge_deck` 的 `storyboard` 新增 `format: "markdown"`，返回可直接放进文档/PR 的 Markdown 提纲（标题、主题、每页 layout/blocks/notes/时长）。
- **测试**：`test_deck_storyboard` 增加 2 项 Markdown 断言，engine 281 项 PASS。

## 1.8.0 — 2026-11-24

- **Built-in template packs**：`resources/template_packs.json` 新增 `officeforge-academic`（论文汇报/学位答辩）与 `officeforge-pitch`（融资路演/产品发布）；pack 内模板通过 `templatePack` + `template` 解析，渲染仍走现有 legacy/studio 后端。
- **打包**：`files` 增加 `resources/packs/**`，包内自带模板资产。
- **测试**：`test_deck_builtin_packs` 4 项 PASS，engine 279 项 PASS。

## 1.7.0 — 2026-11-24

- **Deck storyboard**：`officeforge_deck` 新增 `operation: "storyboard"`，返回每页的 layout、blockTypes、notes 与 `estimatedSeconds`，并给出 `totalEstimatedSeconds`；纯规划，不写文件。
- **测试**：`test_deck_storyboard` 4 项 PASS，engine 275 项 PASS。

## 1.6.0 — 2026-11-24

- **zbsph adapter**：`lib/office_ppt_zbsp_bridge.mjs` 现在把 `table`、`chart`、`image` 转成 PPTD 原生元素（`elementType: table/chart/image`），不再回退 legacy；`media/comments/master/background/shapes/transition/animation/SmartArt/morph` 仍明确 `ZBSPH_UNSUPPORTED` 并回退。
- **测试**：`test_deck_zbsp` 增加 table/chart/image 持久化断言，engine 271 项 PASS。

## 1.5.1 — 2026-11-24

- **修复**：`probe` 的 `pptx.engines` 现在返回 `auto` 加全部已注册 adapter（legacy/studio/zbsph/anon1009/guizang/composer），doctor 不再显示旧引擎清单。

## 1.5.0 — 2026-11-23

- **Bundled composer web adapter**：`ComposerPptAdapter` 用内置 `office_webdeck` 渲染单文件 HTML deck；`officeforge_deck` 的 `engine: "composer"` 可用于 `render`，输出 `.html`，并保留 `officeforge_web` 的独立入口。
- **Adapter Registry**：`composer` 从描述条目变为内置 web-deck adapter；`guizang` 的内置 9 套调色板仍可用。
- **测试**：新增 `test_deck_composer`（4 项：available、kind、adapter/type、html 落盘），engine 268 项 PASS；plugin 55、config 14、compatibility 54。
- **文档**：README/README.zh、DESIGN、COVERAGE 更新到 v1.5.0；`composer` 的 HTML 输出走 `office_webdeck`，不依赖外部 npm 包。

## 1.4.0 — 2026-11-23

- **Bundled 歸藏 design tokens**：`resources/guizang-themes.json` 内置 9 套 magazine/swiss 调色板（ink-classic、indigo-porcelain、forest-ink、kraft-paper、dune、ikb、lemon、lemon-green、safety-orange）；`office_pptx._load_themes()` 同时加载主主题与歸藏主题，`officeforge_deck` 可直接用 `engine: "guizang"` + `theme: "ikb"` 渲染原生 `.pptx`。
- **Adapter Registry**：`guizang` 从描述条目变为内置可用 adapter（`GuizangPptAdapter`），仍保留 `dsh-guizang-design` 完整 skill 作为可选安装；`composer` 继续为可选外部条目。
- **测试**：新增 `test_deck_guizang`（5 项：catalog available、kind、adapter、palette render、文件落盘），engine 264 项 PASS；plugin 55、config 14、compatibility 54。
- **文档**：README/README.zh、DESIGN、COVERAGE 更新到 v1.4.0；`resources/guizang-themes.json` 保留上游 MIT 来源说明。

## 1.3.0 — 2026-11-23

- **Bundled Anon1009 aesthetic adapter**：`lib/vendor/anon1009` 内置 MIT 许可的 `ppt_audit.py` / `ppt_beautify.py` / `themes.py` 与主题资产；`lib/office_ppt_anon.py` 提供 `audit_anon` 与 `beautify_anon`，通过 `officeforge_deck` 的 `engine: "anon1009"` 调用。
- **审计/修复**：`audit` 返回 0–100 score、grade、issueCounts 与逐条 issues；`repair` 执行非破坏式主题、字体、对比度、对齐与页码美化，并写到新的 `.pptx`。
- **测试**：新增 `test_deck_anon1009`（5 项：fixture、catalog available、audit score、repair output、文件落盘），engine 259 项 PASS；plugin 55、config 14、compatibility 54。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v1.3.0；`lib/vendor/anon1009/LICENSE` 保留上游 MIT 许可。

## 1.2.0 — 2026-11-23

- **Real zbsph adapter**：`lib/office_ppt_zbsp_bridge.mjs` 调用本机 `dsh-ppt-studio` 的 `resolveDeck`/`exportPptx`，支持文本、标题、列表、段落、备注的 deck 生成；复杂对象（table/chart/image/media/comments/master/morph/animations/SmartArt）明确 `ZBSPH_UNSUPPORTED` 并回退 legacy。
- **Adapter selection**：`officeforge_deck` 的 `engine` 可选 `zbsph`；安装 `dsh-ppt-studio` 时使用真实 adapter，未安装时仍返回 `ADAPTER_UNAVAILABLE`。
- **测试**：`test_deck_zbsp` 与更新后的 `test_deck_v100`，engine 254 项 PASS；plugin 55、config 14、compatibility 54。

## 1.1.0 — 2026-11-23

- **Template Packs**：`template_packs.json` 支持内置包与外部包；外部包通过 `OFFICEFORGE_TEMPLATE_PACKS` 指向目录（目录内放 `template_pack.json`、`themes.json`、`layouts.json`、`templates.json`）或直接指向 manifest 文件；`officeforge_deck` 的 `catalog` 展示每个包的 theme/layout/template 实际数量。
- **Deck IR**：`normalize_deck` 支持 `templatePack` + `template` 解析包内模板；`deck_to_backend_spec` 不再把已解析的 template 标记传给 legacy 后端，避免二次解析。
- **测试**：新增 `test_deck_template_pack`（外部包发现、资产计数、包内模板 plan、legacy render），engine 250 项 PASS；plugin 55、config 14、compatibility 54。

## 1.0.0 — 2026-11-23

- **Versioning**：项目结束 0.x 阶段，进入稳定 API 版本线；后续功能版本使用 `1.x.0`，修复版本使用 `1.x.y`。
- **Adapter Registry**：`lib/office_deck_registry.py` 增加 `zbsph`、`anon1009`、`guizang`、`composer` 四个外部 adapter 描述条目（kind/capabilities/available/hint），`officeforge_deck` 的 `catalog` 可完整列出；未安装的 adapter 明确返回 `ADAPTER_UNAVAILABLE` 并附带安装提示，不会静默回退。
- **Deck engine 选择**：`officeforge_deck` 的 `engine` 枚举扩展为 `auto|legacy|studio|zbsph|anon1009|guizang|composer`；`auto` 仍只选择 legacy/studio，外部 adapter 必须显式选择。
- **测试**：新增 `test_deck_v100`（外部 adapter 目录、kind 和 unavailable 行为），engine 246 项 PASS；plugin 55、config 14、compatibility 54。

## 0.18.0 — 2026-11-23

- **Studio adapter 扩容**：`lib/office_ppt_studio_bridge.mjs` 现支持 Deck IR 的 table/chart/image blocks（含 headers/rows、categories/series、图片 path/alt/fit），并支持 `text`/`timeline`/`matrix` 块映射；真正无法由 dsh-ppt 表达的 comments/media/master/morph/animations/SmartArt/shapes 仍明确 `STUDIO_UNSUPPORTED` 并回退 legacy。
- **Deck IR 完善**：`lib/office_deck_ir.py` 增加 `DECK_BLOCK_TYPES` 校验与 `_blocks_to_legacy_slide` 展平；`resources/deck_schema.json` 同步 block 类型与字段。
- **测试**：新增 `test_deck_v018`（table/chart/image blocks + 未知 block 拒绝），engine 242 项 PASS；plugin 55、config 14、compatibility 54 通过。

## 0.17.0 — 2026-11-23

- **新增**：Deck IR + Adapter Registry + Template Pack 架构。`lib/office_deck_ir.py` 提供 `normalize_deck`/`deck_to_backend_spec`/`summarize_deck`/`catalog`，支持 `{deck:{...}}` 包装与模板展开；`lib/office_deck_registry.py` 注册 legacy/studio adapter；`lib/office_deck.py` 提供 plan/render/audit/repair/catalog/web。
- **新增**：`officeforge_deck` 工具（plan/render/audit/repair/catalog/web），共十六个工具；`resources/template_packs.json` 与 `resources/deck_schema.json`。
- **测试**：`tests/test_engine.py` 新增 `test_deck_v017`（9 项），engine 236 项 PASS；plugin 55、config 14、compatibility 54 PASS；回归基线同步。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.17.0。

## 0.16.0 — 2026-11-23

- **Studio 后端真正接入 dsh-ppt**：`StudioPptBackend` 通过 `dsh-ppt/deck-advanced` 的 `buildDeckAsync` 渲染支持的纯文本/KPI/items deck；遇到 comments/media/master/morph/animations/SmartArt/图片/表格/图表等对象时自动回退 legacy python-pptx。
- **桥接层**：新增 `lib/office_ppt_studio_bridge.mjs`，stdin/stdout JSON，明确 `STUDIO_UNSUPPORTED` 信号，不注册 prompt/LLM hook。
- **依赖**：`dsh-ppt@0.33.0` 作为可选渲染器依赖；不新增 DSH bundle、不改变 `dsh.profile.bundles`。
- **测试**：engine 227 项 PASS；新增 `studio falls back for unsupported slide objects`；plugin 54、config 14、compatibility 52 PASS。

---

## 0.15.0 — 2026-11-23

- **新增**：PowerPoint 双后端抽象 `lib/office_ppt_backends.py`（`auto` / `legacy` / `studio`），配置项 `pptxEngine`；`auto` 在指定 template/theme 时选择 studio，否则保持 legacy。
- **新增**：`officeforge_beautify`（主题套用、标题/正文字号与段落间距规范化、before/after deck score、默认备份）与 `officeforge_web`（自包含 HTML deck，16:9、键盘/按钮翻页、打印样式）。
- **新增**：`officeforge_audit` 的 `pptx` `aesthetic` 模式，附加标题字号检查与 0–100 审美分数。
- **测试**：`tests/test_engine.py` 新增 `test_ppt_v015`（legacy/studio/auto 选择、aesthetic、beautify、HTML deck、feature 关闭拒绝），engine 226 项 PASS；plugin 54 项、config 14 项、compatibility 52 项 PASS。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.15.0。

## 0.14.0 — 2026-11-22

- **新增**：插件页一键环境检查卡片。`lib/client.js` 的 `OfficeForgeStatus` 增加“检查环境”按钮，点击 `fetch('/officeforge/doctor')` 并渲染 `health`、`summary` 与逐项 `checks`（含 `remedy`）；卡片保持只读，不写配置、不执行文档操作。
- **新增**：host 侧只读 `GET /officeforge/doctor` 路由。`lib/index.js` 复用 `buildDoctorReport`，输出与 `officeforge_doctor` 一致的结构；非 GET 返回 405，异常时仍返回诊断 JSON。
- **测试**：`tests/test_plugin.mjs` 新增 doctor HTTP route registered / rejects non-GET / live checks 三项，plugin 52 项 PASS；`tests/test_client.mjs` 与 `tests/test_client_host.mjs` 覆盖按钮与 bundle 契约；回归基线 pluginPass 52。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.14.0。

## 0.13.5 — 2026-11-22

- **新增**：`officeforge_doctor` 增加可选 OCR 后端检查（PyMuPDF、pytesseract、Tesseract 可执行文件与语言包），未安装时给出安装提示，不把 OCR 定为必需依赖。
- **测试**：`tests/test_plugin.mjs` 新增 `doctor reports OCR backend`，plugin 49 项 PASS；回归基线 pluginPass 49。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.13.5。

## 0.13.4 — 2026-11-22

- **修复**：重新校准 `benchmarks/baseline.json`。原阈值大约只有实测值的 1.1–1.2 倍，正常机器波动就会让 12 个基准中的 9 个误报 FAIL；现按实测值的 1.5 倍重设（最低 1 秒），并将峰值内存留出约 50% 余量。
- **验证**：`benchmarks/run_benchmarks.py` 全量 PASS；engine 216、fuzz 37、plugin 48、config 14、compatibility 48 全部 PASS。
- **文档**：README/README.zh、DESIGN、COMPATIBILITY 更新到 v0.13.4。

# Office Forge Changelog

## 0.13.1 — 2026-11-22

- **修复**：`officeforge_preview` 在真实宿主中报 `value is not lossless JSON`。根因是 host 的 `snapshotJsonValue`（`@deepseek-ai/dsh-util-values`）拒绝任何值为 `undefined` 的对象成员、稀疏数组与非有限数字，而 `dsh-office-kit` 的 `renderImages` 只在实际适用时才带 `page`/`sheet`/`range`（`lib/index.js:1215-1229`），`convertTo` 也只在有页数时带 `pageCount`；插件把 `image.page` 等原样放进返回值，`undefined` 就漏给了宿主，整次调用被判定为非 lossless JSON。
- **修复**：新增 `lossless()` 深度归一化函数，丢弃 `undefined` 成员与非有限数字、把原型非 `Object.prototype` 的值判为不可序列化；`runEngine` 在把引擎结果交回宿主前统一过一遍，`officeforge_preview`/`convert`/`recalculate`/`export`/`doctor`/`batch` 的返回体也各自归一化。缺字段仍然缺字段，不编造占位值。
- **修复**：`runEngine` 遇到不可序列化的引擎结果时改为抛 `ENGINE_BAD_OUTPUT`，不再把坏值传给宿主。
- **测试**：`tests/test_plugin.mjs` 的 officeKit mock 改为按真实行为省略 `page`/`sheet`/`range` 与 `pageCount`/`dpi`，新增 6 项断言（可选字段被丢弃而非转发 `undefined`、preview/convert/recalculate/doctor 输出通过 host lossless 闸门），plugin 46 项 PASS，回归基线 pluginPass 46。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.13.1。

## 0.13.0 — 2026-11-22

- **PPT 深化**：新增 `resources/layouts.json`（38 个版式：cover/section/agenda/bullets/双栏/三栏/四象限/kpi/timeline/roadmap/process/comparison/matrix/quote/image-*/gallery/table/chart/chart-split/cards/team/closing 等）与 `resources/templates.json`（10 套成品 deck：board/project/pitch/academic/training/sales/report/product/strategy/minimal），`officeforge_create` 的 pptx spec 支持 `template`、`master`、`layout` 走版式库；`officeforge_edit` 新增 `addComment`/`listComments`（原生 `p:cmLst` + `p:cmAuthorLst` 部件）、`addMedia`（音频/视频 + poster + autoplay/loop）、`applyMaster`（母版背景与页眉页脚）；`setTransition` 新增 `morph` 效果（支持 `option`）。
- **审计**：`officeforge_audit` 新增 pptx `layouts`/`templates` 目录模式，可无 path 调用。
- **测试**：`tests/test_engine.py` 新增 `test_pptx_deepening`（20 项：版式目录/模板目录/模板展开/KPI 与时间线渲染/批注写入与列举/媒体内嵌/morph 持久化/母版页脚/重新打开/comments feature 关闭/TEMPLATE_NOT_FOUND），引擎 216 项 PASS，回归基线 enginePass 216。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY、locale 更新到 v0.13.0 与 PPT 深化能力。

## 0.12.0 — 2026-11-22

- **Word 深化**：`officeforge_edit` 新增 `addCitation`（上标引文标记）、`addBibliography`（参考文献列表）、`setColumns`（分栏）、`addWatermark`（页眉水印）、`addContentControl`（结构化文档标签）、`saveAsTemplate`（导出 `.dotx` 模板）；`officeforge_create` 的 docx spec 支持 `columns`、`watermark`、`contentControls`、`bibliography`，`.dotx` 输出写入模板内容类型。
- **测试**：`tests/test_engine.py` 新增 `test_docx_deepening`（分栏/水印/内容控件/参考文献/引文/模板内容类型/feature 关闭拒绝），引擎 196 项 PASS，回归基线 enginePass 196。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY、locale 更新到 v0.12.0 与 Word 深化能力。
## 0.11.0 — 2026-11-22

- **性能基线**：新增 `benchmarks/run_benchmarks.py` 与 `benchmarks/baseline.json`，12 个真实引擎用例覆盖 Word 创建/读取/审计、Excel 创建/读取/审计、PPT 创建/读取/审计、损坏/空/加密 PDF；记录墙钟时间与 Windows 峰值工作集，超阈值即失败。
- **回归门禁**：`tests/run_regression.py` 完整模式新增 `benchmarks` 可选套件（`--quick` 跳过），`tests/regression_baseline.json` 增加 `benchmarkPass=12`；`tests/test_engine.py` 新增 `test_pdf_invalid`，引擎基线升至 186。
- **PDF 修复**：`lib/office_pdf.py` 对不以 `%PDF-` 开头的文件显式抛 `INVALID_PDF`，修复损坏/空文件被 fallback 当成空文档的问题。
- **文档**：新增 `docs/BENCHMARKS.md`，README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.11.0。

## 0.10.0 — 2026-11-22

- **Excel 原生能力**：`officeforge_create` 的 xlsx spec 支持 `import` 从 CSV/TSV/JSON 导入（`source`/`format`/`sheet`/`header`），CSV 数值自动解析；`officeforge_edit` 的 `setCells` 支持 `arrayFormula` + `ref` 写入数组公式；`charts` 支持 `type:"combo"`，按 series 的 `type`/`axis` 生成柱线组合图，自动设置次坐标轴（`axId 200`）。
- **测试**：`tests/test_engine.py` 新增 `test_xlsx_v10`（CSV/JSON 导入、数组公式持久化、组合图主/次坐标轴、import feature 关闭拒绝），引擎 184 项 PASS；回归基线 enginePass 同步为 184。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.10.0 与 Excel 导入/数组公式/组合图说明；`officeforge_native` 的原生 PivotTable 仍依赖可选 COM 后端。

## 0.9.0 — 2026-11-22

- **PDF 完整工具链**：`officeforge_pdf` 在 merge/split/rotate/watermark 之外新增 `pageNumbers`（页码，支持格式/位置/起始值/字号/边距）、`formFields`（表单字段清单）、`fillForm`（表单填写，支持 flatten）、`extractImages`（页面图片提取为 PNG）、`ocr`（可选 PyMuPDF + pytesseract，缺依赖抛 `BACKEND_UNAVAILABLE`）、`encrypt`/`decrypt`（用户/所有者密码）；加密 PDF 读取无密码抛 `PDF_PASSWORD_REQUIRED`，错误密码抛 `PDF_PASSWORD_INVALID`。
- **读取支持**：`officeforge_read` 增加 `password` 参数并支持加密 PDF；`lib/office_pdf.py` 的 pypdf 路径显式处理 `is_encrypted`。
- **测试**：`tests/test_engine.py` 新增 `test_pdf_tools_v090`（页码、加密/解密、表单字段/填写、图片提取、OCR 缺依赖分类），引擎 176 项 PASS；回归基线 enginePass 176。
- **文档**：README/README.zh、DESIGN、COVERAGE、COMPATIBILITY 更新到 v0.9.0 与完整 PDF 工具链说明。


## 0.8.4 — 2026-11-22

- **跨机可用性自检**：`officeforge_doctor` 现在逐项检查 Python 解释器与版本、`python-docx`/`openpyxl`/`python-pptx`、内置 pypdf、COM/pywin32、已注册的 Word/Excel/PowerPoint、LibreOffice、`officeKit` 服务与客户端 bundle；每项返回 `required`、`ok`、`status`（ok/warn/missing/disabled）、`detail` 与 `remedy`，并汇总 `health`（ok/degraded/error）和 `summary`。即使引擎本身无法启动，doctor 也会返回结构化诊断而不是抛异常。
- **后端状态扩展**：`office_backends.status()` 增加 Python 可执行文件/版本、模块版本映射与 Office 应用注册信息；不启动 Office 进程。
- **测试**：`tests/test_plugin.mjs` 增加 doctor 的 health、pythonModules、checks 字段与 summary 断言；plugin 测试 40 项 PASS，回归基线同步更新。
- **文档**：README/README.zh 的 `officeforge_doctor` 行改为跨机自检说明；DESIGN/COVERAGE/COMPATIBILITY 更新到 v0.8.4。

## 0.8.3 — 2026-11-22

- **Word 文档审计**：`officeforge_audit` 现在支持 `type: "docx"`。`health` 检查空文档、标题跳级、同层级重复标题、标题后无正文（孤立标题）、未清除的 `{{...}}` 模板占位符、TODO/TBD/FIXME 草稿标记、超过 1500 字符的长段落、连续空段落，并输出 score / issueCount / summary / issues；`profile` 返回块/段落/标题/表格/图片/节统计、标题层级分布、样式使用频次、页面尺寸/方向、表格行列数与核心元数据。
- **feature 开关**：新增 `word.audit`（默认开启）；关闭时引擎抛 `FEATURE_DISABLED`。
- **测试**：`tests/test_engine.py` 增加 `test_docx_audit`（12 项），覆盖全部 health 检查、profile 统计、默认 health 模式与 feature 关闭拒绝；engine 全量 166 项 PASS，config 14 项 PASS，回归基线同步更新。

## 0.8.2 — 2026-11-21

- **文档元数据默认值统一**：三个创建路径（docx/xlsx/pptx）现在都会覆盖模板自带的核心属性，不再把 `python-docx` / `python-pptx` / `openpyxl` 的占位作者和 2013 年模板时间带进生成文件。未显式指定时 `author`/`lastModifiedBy` 为 `Office Forge`，`created`/`modified` 为当前 UTC 时间，`revision` 为 1。
- **`spec.properties` 统一支持**：docx 现在和 xlsx/pptx 一样接受 `properties: {title, subject, author, keywords, comments, category, lastModifiedBy}`；docx 历史上使用的顶层 `author`/`subject`/`keywords`/`comments` 继续兼容。显式值始终优先于默认值。
- **docx 模板注释清理**：python-docx 默认模板的 `comments = "generated by python-docx"` 会在未提供 `comments` 时被清空，避免把库内部说明写进文档属性。
- **新增测试**：`tests/test_engine.py` 增加 `test_document_properties`（20 项），覆盖三种格式的默认作者、时间、revision、显式值覆盖与 docx 旧写法兼容；engine 全量 154 项 PASS，回归基线同步更新。
- 本机构建环境已把 0.8.2 安装到 desktop profile 并实机验证；源码与离线测试仍不依赖任何 DSH profile。

## 0.8.1 — 2026-11-21

- **COM 进程回收（安全加固）**：`office_com_worker.py` 现在登记每一个由它启动的 Office automation 实例，退出路径统一走 `_release_all()`；`_safe_quit` 在 `Quit()` 失败或进程仍存活时，按窗口句柄解析 PID 并用 `taskkill /F /T` 强制回收。此前 worker 被超时杀掉后，它拉起的 `WINWORD/EXCEL/POWERPNT` 会变成孤儿进程留在本机。
- **超时整树回收**：`office_backends.run_com` 从 `subprocess.run` 改为显式 `Popen` + `communicate(timeout)`，超时后对 worker 执行 `taskkill /F /T` 再收尸，确保 Office 子进程不逃逸；同时保留 `BACKEND_TIMEOUT` 错误码与 stderr 摘要。
- **相对路径解析修正**：`office_engine._checked` 此前用 `os.path.abspath()` 解析相对路径，会落到引擎进程自己的 cwd，导致 `missing.docx` 这类请求被误报为 `PATH_OUTSIDE_WORKSPACE`。现在相对路径一律相对 workspace 解析，错误码恢复为 `INPUT_NOT_FOUND` / `INVALID_ARGUMENT` / `UNSUPPORTED_FORMAT` 等真实原因。
- **新增 `tests/test_fuzz.py`**（37 项）：覆盖空 payload、未知 action、路径逃逸、类型错配、畸形 JSON、截断 JSON、嵌套 batch、过高 schemaVersion 等 36 种畸形输入，断言引擎始终返回单个 JSON 且带稳定错误码、不泄漏 traceback、不在 workspace 外落盘，并验证同一坏请求两次响应完全一致。
- **新增 `tests/test_com_cleanup.py`**（3 项）：真实 COM `exportPdf` 前后比对 `WINWORD/EXCEL/POWERPNT` 进程数，确认无泄漏；并用挂起脚本验证超时后整棵进程树被回收。
- **回归脚本重写**：`tests/run_regression.py` 现在一次跑完 engine/fuzz/plugin/config/compatibility/client/clientHost，可选 comCleanup 与 dshProfile，除断言数下限外还要求每个套件打印完成标记，避免“零断言也算通过”。支持 `--quick` 跳过需要 Office 与 DSH CLI 的套件。
- **测试**：engine 134、fuzz 37、plugin 36、config 13、compatibility 48、comCleanup 3；回归基线同步更新。
- 仍不安装到真实 DSH profile。

## 0.8.0 — 2026-11-21

- **Word 脚注**：新增 `addFootnote`，手工注入 `/word/footnotes.xml` part（含 `w:footnoteRef` 与 `FootnoteText` 样式），正文写入 `w:footnoteReference`；无该 part 的文档会自动创建并登记 content type 与 relationship。
- **Word 书签与交叉引用**：新增 `addBookmark`（`bookmarkStart`/`bookmarkEnd`，重复名返回 `BOOKMARK_EXISTS`）与 `addCrossReference`（`REF <name> \h` 域并打开 `updateFields`，缺失书签返回 `BOOKMARK_NOT_FOUND`）。Word 只接受 ASCII 书签名，非 ASCII 请求会经 SHA-1 稳定映射为 `bm_<8hex>`。
- **Excel 原生表格对象**：新增 `addTable` 与 create spec 的 `tables`，写入真正的 `openpyxl.worksheet.table.Table` + `TableStyleInfo`（可回读为 ListObject），重复名返回 `TABLE_EXISTS`。
- **Excel 保护**：新增 `protectSheet` / `protectWorkbook` 与 create spec 的 `protection` / `workbookProtection`，映射 `SheetProtection`、`WorkbookProtection` 的密码与格式/排序/筛选等选项。
- **PowerPoint 备注与节**：新增 `setNotesBatch`（批量写/清空演讲者备注）与 `addSection` / `listSections`；节写入 `p:extLst` 中 `p14:sectionLst` 扩展（URI `{521415D9-36F7-43E2-AB2F-B90AF26B5E84}`），重复占用返回 `SECTION_CONFLICT`。
- **修复**：`p14:sectionLst` 最初被挂到 `p:sldIdLst` 内，会让 python-pptx 再次打开时报 `AttributeError: 'lxml.etree._Element' object has no attribute 'rId'`；已改挂到 `p:presentation` 根节点的 `p:extLst`。
- **配置**：`resources/features.json` 增加 word `footnotes`/`bookmarks`/`crossReferences`、excel `tables`/`protection`、ppt `notes`/`sections`；引擎 probe capabilities 与 `officeforge_edit` 描述同步。
- **测试**：engine 134、plugin 37、config 14、compatibility 49；新增 docx/xlsx/pptx v0.8 三组用例，覆盖持久化、feature 关闭拒绝、重复/冲突错误码；回归基线同步更新。
- 仍不安装到真实 DSH profile。

## 0.7.0 — 2026-11-21

- **插件页状态卡**：新增 `lib/client.js` 客户端 bundle（`dsh.client.platform = "web"`、`exports["./client"]`），在插件详情页通过 `plugins.detail.section` slot 渲染只读状态卡，展示版本、工具清单与 `officeforge_doctor` 自检入口；不修改配置、不在插件页执行文档操作。
- **客户端契约对齐**：bundle 采用宿主 lazy-CJS 形态（`window.__ModuleLoader__.load({id, factory})`，factory 返回 `module.exports`，`inject = ["slots"]`）；`ClientModuleSystem.register` 的重复 id 校验与 `createClientModuleSystem` 的 boot manifest 解析均已实测通过。
- **新增测试**：`tests/test_client.mjs`（manifest + bundle 形态 + slot 注册）、`tests/test_client_host.mjs`（用真实 `@deepseek-ai/dsh-client-modules` 驱动自家 bundle）。
- **安装验证**：`tests/test_dsh_profile.mjs` 在临时 `DSH_HOME` + 临时 profile 中安装 0.7.0 tarball，额外校验 `dsh.client.platform`、`exports["./client"]` 与 bundle 内容，结束后清理；未触碰真实 desktop profile。
- 未给 `Config` 加 `.volatile()`：工具在 `apply` 时静态注册，避免出现「UI 开关生效但工具未重注册」的不一致。

## 0.6.1 — 2026-11-21

- **运行诊断**：新增第 13 个工具 `officeforge_doctor`，一次报告 Python、内置 PDF 引擎、COM/LibreOffice、officeKit 和当前模块/功能开关状态，便于安装后自检。
- **测试**：engine 105、plugin 36、config 13、compatibility 48；新增 doctor 注册与 pdfVendor/checks 断言。
- **安装验证**：在隔离 DSH_HOME 的临时 profile 中用 npm tarball 完成安装，patch 可被 `loadOverlayPatches` 识别；未触碰真实 desktop profile，未安装到真实 DSH。

## 0.6.0 — 2026-11-21

- **PDF 文件操作**：新增 `officeforge_pdf`，支持 `merge`、`split`、`rotate`、`watermark`，包内 vendor `pypdf`，不依赖外部 PDF 服务。
- **原生 COM Worker**：新增 `officeforge_native` 与 `lib/office_com_worker.py`，支持 `convert`、`exportPdf`、`acceptRevisions`、`recalculate`、`pivot`（原生 PivotTable）、`smartArt`（原生 SmartArt）；自动探测系统 Python + pywin32，缺后端返回 `BACKEND_UNAVAILABLE`。
- **后端增强**：`office_backends.py` 新增 `com_python_path`、`run_com`、`export_pdf`；旧格式转换优先 COM，失败再 LibreOffice。
- **Engine/Node**：新增 `pdfop`/`native` action，`officeforge_pdf`/`officeforge_native` 两个工具，工具总数 12；Config 增加 `comPythonPath`；`officeforge_pdf` 只受 `modules.pdf` 控制。
- **测试**：engine 105、plugin 32、config 13、compatibility 46；新增 PDF 工具与 native exportPdf 断言，COM 手动实测 exportPdf/recalculate/pivot/SmartArt；仍不安装 DSH。

## 0.5.1 — 2026-11-21

- **Word 审阅**：新增 `addComment`、`trackRevision`、`mailMerge` 三个编辑操作；批注写入 comments part，修订写入 `w:del`/`w:ins`，邮件合并复用 `{{key}}` 占位符替换。
- **Excel 透视摘要**：新增 `pivotSummary`，按行字段分组并对数值列做 sum/count/avg/min/max 聚合，写入新工作表；这是可移植的值汇总，不是原生 PivotTable 对象。
- **PPT 动画与 SmartArt 风格图**：新增 `addAnimation`（fade/fly/wipe/zoom/appear）与 `addSmartArt`（可编辑圆角矩形+连接线，`nativeSmartArt:false`），`features` 增加 `animations`/`smartart`。
- 测试：engine 100、plugin 30、config 13、compatibility 42；仍不安装 DSH。

## 0.5.0 — 2026-11-21

- **PDF 读取**：`officeforge_read` 支持 `.pdf`，优先 pypdf/fitz/pdfplumber，缺依赖时使用内置 zlib/content-stream fallback。
- **旧格式兼容**：新增 `officeforge_convert`，通过可选 Windows COM / LibreOffice 后端转换 `.doc/.docm/.xls/.xlsm/.ppt/.pptm`；后端缺失返回 `BACKEND_UNAVAILABLE`，不阻断 Native 核心。
- **视觉预览与重算**：新增 `officeforge_preview`（officeKit.renderImages）与 `officeforge_recalculate`（officeKit.recalculate）。
- **主题外部化**：`resources/themes.json` 提供 31 套内置主题，`resources/features.json` 提供模块/能力注册表；新增主题/开关不需要改 Python 代码。
- **Schema 迁移**：`office_migrations.py` 将缺失 `schemaVersion` 的旧 spec 视为 v1，迁移到 v2，并记录 `migratedFrom`；高于 v2 返回 `SCHEMA_VERSION_UNSUPPORTED`。
- **兼容性**：新增后端探测、预览/重算参数校验、源文件版本检查、远程 workspace 映射检查；工具总数从 7 增至 10。
- 测试：engine 覆盖 PDF/迁移/convert 降级；Node 覆盖十工具注册、convert/preview/recalculate mock；仍不安装 DSH。

## 0.4.1 — 2026-11-21

- 兼容性加固：`officeforge_export` 使用根会话 owner、源文件 `stat.version` 前后校验、远程 workspace `processPathFromHostPath` 检查。
- 非法 profile 配置回退默认值；Excel 图表数据使用唯一隐藏工作表，不覆盖用户表。
- 新增 `docs/COMPATIBILITY.md` 与 `tests/test_compatibility.mjs`（schema、工具名、模块裁剪、隔离静态检查）。

## 0.4.0 — 2026-11-21

- 新增 `officeforge_audit`：Excel `health/profile/insight/diff`，PPT `deck/profile/themes`。
- PowerPoint 8 套主题、`applyTheme`；Excel 公式审计（错误/模式不一致/硬编码/空行断链/循环引用）。

## 0.3.0 — 2026-11-21

- Word `addCaption`、`setStyle`、`updateFields`；Excel `addFormulaColumn`、`sortRange`、`removeDuplicates`、`fillBlanks`、`cleanRange`；PPT `addTextBox`、`addConnector`。
- 参考公开的 Office 自动化能力清单，保持纯文件级实现。

## 0.2.0 — 2026-11-21

- 新增 `modules` / `features` 配置、Word 图片/超链接/TOC、Excel 图表/条件格式/数据验证/批注/合并/图片、PPT 切换/形状；主题/功能开关化。

## 0.1.0 — 2026-11-21

- 首个版本：`officeforge_read/create/edit/template/batch/export` 六工具，Python 子进程 + 原子写入 + 路径隔离。

## Migration

- v0.4 → v0.5：spec 可带 `schemaVersion: 2`；旧 spec 无需修改，引擎自动补 `migratedFrom: 1`。
- v0.3 → v0.4：新增 `officeforge_audit`，原有工具参数不变。
- v0.2 → v0.3：新增编辑操作，旧操作保持不变。
