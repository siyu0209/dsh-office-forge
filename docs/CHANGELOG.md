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
- **主题外部化**：`resources/themes.json` 提供 31 套主题（dsh-ppt 24 套 + 7 套旧主题），`resources/features.json` 提供模块/能力注册表；新增主题/开关不需要改 Python 代码。
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
- 参考 `@copylee/dsh-office` 的 COM 深度能力清单，但保持纯文件级实现。

## 0.2.0 — 2026-11-21

- 新增 `modules` / `features` 配置、Word 图片/超链接/TOC、Excel 图表/条件格式/数据验证/批注/合并/图片、PPT 切换/形状；主题/功能开关化。

## 0.1.0 — 2026-11-21

- 首个版本：`officeforge_read/create/edit/template/batch/export` 六工具，Python 子进程 + 原子写入 + 路径隔离。

## Migration

- v0.4 → v0.5：spec 可带 `schemaVersion: 2`；旧 spec 无需修改，引擎自动补 `migratedFrom: 1`。
- v0.3 → v0.4：新增 `officeforge_audit`，原有工具参数不变。
- v0.2 → v0.3：新增编辑操作，旧操作保持不变。
