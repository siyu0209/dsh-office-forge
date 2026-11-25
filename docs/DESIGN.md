# Office Forge 设计说明

## 1. 目标与边界

Office Forge 是一个独立的 DSH Office 文档插件，把常用的 Word、Excel、PowerPoint、PDF 文档能力收敛成十六个互相独立的工具。它只处理磁盘上的文档文件，不依赖 Office UI，也不依赖其他文档插件。

核心边界：

- 读取、创建、编辑 `.docx`、`.xlsx`、`.pptx`；只读解析 `.pdf`。
- 用 JSON / CSV / TSV 数据填充 Word/Excel 模板。
- 在一次受控调用中执行多个独立文档任务。
- 通过宿主 `officeKit` 导出 PDF 等格式，并可选使用 `renderImages` / `recalculate`。
- 文件级 PDF 合并、拆分、旋转、水印、页码、表单、图片提取、加密解密、可选 OCR。
- PowerPoint 双后端（legacy / studio）、审美评分与一键美化、自包含 HTML deck 输出。
- 可选 COM/LibreOffice 后端：旧格式转换、COM 导出 PDF、接受修订、Word 域/目录更新、公式重算、原生 PivotTable、原生 SmartArt。
- 按 `modules` / `features` 配置启用或禁用 Word、Excel、PowerPoint、PDF 能力。

不注册 prompt 变量、不接管 LLM stream、不注册 tool middleware 或 hook，也不与内置 `office_*` 工具重名。

## 2. 设计原则

1. **独立实现**：所有文档逻辑由本插件自己的 Python 模块完成；不嵌入、不转发到其他 Office 插件运行时。
2. **文件级优先**：核心能力只依赖 python-docx、openpyxl、python-pptx、pypdf；Office COM 与 LibreOffice 只是可选增强。
3. **单一职责**：每个 `officeforge_*` 工具只向引擎发送一个 action；工具之间不互相调用。
4. **可预测失败**：所有异常转换成语义稳定的错误码，不返回 traceback，不把失败伪装成成功。
5. **工作区边界**：所有输入输出路径经过宿主与词法双重检查；默认不覆盖已有文件。
6. **可诊断**：`officeforge_doctor` 与插件页只读卡片报告 Python、库、PDF、COM、LibreOffice、officeKit 与客户端 bundle 状态。

## 3. 架构

```
DSH host
  └─ lib/index.js
       ├─ Config schema: modules + features + runtime options
       ├─ enabledTypes() / allowedExtensionsFor() / featuresFor()
       ├─ resolveTarget(): ctx.workingDirectory + ctx.fs.resolve/contains/processPath
       ├─ runEngine(): spawn 短生命周期 Python 子进程
       │    └─ lib/office_engine.py
       │         ├─ office_common.py      路径、原子写、占位符、feature gate
       │         ├─ office_deck_ir.py     Deck IR 规范化、模板包、catalog
       │         ├─ office_deck_registry.py PPT adapter 注册与分发
       │         ├─ office_deck.py        deck plan/render/audit/repair/catalog/web
       │         ├─ office_docx.py        Word 读取/创建/编辑/模板/审计
       │         ├─ office_xlsx.py        Excel 读取/创建/编辑/模板/审计
       │         ├─ office_pptx.py        PowerPoint 读取/创建/编辑/模板/审计
       │         ├─ office_ppt_anon.py    内置 Anon1009 审美审计/美化 adapter
       │         ├─ office_pdf.py         PDF 文本提取
       │         ├─ office_pdf_tools.py   PDF 文件操作
       │         ├─ office_backends.py    可选 COM / LibreOffice 后端
       │         ├─ office_com_worker.py  隔离 COM Worker
       │         └─ office_migrations.py  spec schema 迁移
       ├─ officeKit bridge: export / preview / recalculate
       └─ lib/client.js: 插件页只读状态卡 + 环境检查
```

Node 层只负责配置裁剪、路径策略、参数校验、子进程生命周期与 officeKit 桥接。文档解析和生成全部在 Python 子进程中完成；引擎每次调用无状态。

## 4. 工具分发与无递归约束

每个工具只调用 `office_engine.py` 的一个 action：

- `deck` → `office_deck.plan/render/audit/repair/catalog/render_html`
- `read` → `office_docx.read_docx` / `office_xlsx.read_xlsx` / `office_pptx.read_pptx` / `office_pdf.read_pdf`
- `create` → 对应模块的 `create_*`
- `edit` → 对应模块的 `edit_*`
- `template` → 对应模块的 `template_*`
- `audit` → 对应模块的 `audit_*`
- `pdfop` → `office_pdf_tools.pdf_operation`
- `native` → `office_backends.run_com` / `export_pdf`
- `copy`、`probe`、`doctor` 为引擎内部动作

硬约束：

1. 工具之间不互相调用。
2. `dispatch` 只做一层，不允许 action 再触发工具。
3. `officeforge_batch` 的 item 只允许 `create` / `edit` / `template`；传 `action: "batch"` 直接返回 `INVALID_OPERATION`。
4. `{{#each items}}` 只展开一层，不递归解析内层循环。
5. `_replace_table_text` 使用显式 stack，不使用 Python 递归。
6. 所有输入、输出、文本、行数、页数、batch 条目都有上限。

## 5. 模块与功能开关

`lib/index.js` 从 `resources/features.json` 读取注册表并生成 `Config` schema：

- `modules`：`word` / `excel` / `ppt` / `pdf`，默认全部 `true`。
- `features`：按文档族组织的高级能力，默认全部 `true`。
- `enableBackends`：是否注册 `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate`。
- `officeforge_pdf` 只受 `modules.pdf` 控制，使用包内 vendored pypdf。

关闭模块会从工具参数 enum 与注册集合中移除对应类型；关闭 feature 时引擎返回 `FEATURE_DISABLED`，不静默降级。主题、版式、模板分别从 `resources/themes.json`、`resources/layouts.json`、`resources/templates.json` 加载，新增资源不需要改 Python 代码。

## 6. 路径、写入与错误模型

- `ctx.fs.resolve` / `ctx.fs.contains` 保证宿主层工作区约束；`assert_inside` 再做词法路径检查，拒绝 `PATH_OUTSIDE_WORKSPACE`。
- 输出使用同目录临时文件 + `os.replace`，默认拒绝覆盖；`overwrite: true` 时才替换；可按参数生成同目录备份。
- 常见错误码：`INPUT_NOT_FOUND`、`PATH_OUTSIDE_WORKSPACE`、`UNSUPPORTED_FORMAT`、`OUTPUT_EXISTS`、`FEATURE_DISABLED`、`BACKEND_UNAVAILABLE`、`SCHEMA_VERSION_UNSUPPORTED`、`INVALID_PDF`、`PDF_PASSWORD_REQUIRED`、`PDF_PASSWORD_INVALID`、`ENGINE_BAD_OUTPUT`。
- 宿主输出必须是无损 JSON：`lossless()` 会丢弃 `undefined`、非有限数字等无法序列化的内容；不可序列化时返回 `ENGINE_BAD_OUTPUT`。

## 7. 可选后端与降级

- **COM**：`office_backends.com_python_path()` 探测 `OFFICEFORGE_COM_PYTHON`、`DSH_COM_PYTHON`、PATH 与常见 Python 安装；`run_com()` 用独立子进程执行 `office_com_worker.py`，支持 `convert`、`exportPdf`、`acceptRevisions`、`recalculate`、`pivot`、`smartArt`、`updateFields`。
- **LibreOffice**：探测 `soffice` / `libreoffice`，用于旧格式转换与预览；不存在时只影响可选功能。
- **OCR**：`office_pdf_tools.ocr_pdf` 需要 PyMuPDF、pytesseract 与 Tesseract；`officeforge_doctor` 把它标记为 optional warn。
- 缺少任何可选后端时，核心文件级工具保持可用；`officeforge_native` / `officeforge_convert` 返回 `BACKEND_UNAVAILABLE`。

## 8. 兼容性与隔离

- `package.json` 使用官方 `dsh.bundle.patch` 与 `dsh.client` 声明；`cordis.patch.yml` 只插入一个最小 entry。
- `apply(ctx, config)` 使用 `ctx.effect()` 注册并按逆序清理；十六个工具逐个 `safeRegister()`，单工具失败只记录 warning。
- 所有工具名以 `officeforge_` 开头，不覆盖内置 `office_render` / `office_convert` / `office_recalculate` / `run_code`。
- 不注册 `systemPrompt`、LLM stream、middleware、host hook；不修改 DSH 安装目录或 profile。
- 客户端 bundle `lib/client.js` 只注入 `slots`，只注册 `plugins.detail.section` 只读卡片（id `officeforge.status`，order 40）；不注册 `plugins.bundle.config`，不写配置。

## 9. 测试与基准

| 层 | 测试 | 结果 |
|---|---|---|
| Python 引擎 | `tests/test_engine.py` | `ALL TESTS PASSED`（334 项） |
| 畸形输入 | `tests/test_fuzz.py` | `FUZZ TESTS PASSED`（45 项） |
| 插件层 | `tests/test_plugin.mjs` | `PLUGIN TESTS PASSED`（72 项） |
| PPT 桥接实测 | `tests/test_ppt_bridges.mjs` | `PPT BRIDGE TESTS PASSED`（140 项） |
| 配置 | `tests/test_plugin_config.mjs` | `PLUGIN CONFIG TESTS PASSED`（14 项） |
| 兼容性 | `tests/test_compatibility.mjs` | `COMPATIBILITY TESTS PASSED`（54 项） |
| 客户端 | `tests/test_client.mjs` / `tests/test_client_host.mjs` | `CLIENT UI TESTS PASSED` / `CLIENT HOST MACHINERY TESTS PASSED` |
| 一次性安装 | `tests/test_dsh_profile.mjs` | `DSH PROFILE TESTS PASSED` |
| 基准 | `benchmarks/run_benchmarks.py` | `BENCHMARK PASSED`（13 项） |
| 回归 | `tests/run_regression.py` | `REGRESSION BASELINE PASSED` |

完整阈值见 `tests/regression_baseline.json` 与 `benchmarks/baseline.json`。

## 9.1 版本策略

从 `1.0.0` 起，插件使用稳定 API 版本线：新增向后兼容能力用 `1.x.0`，修复用 `1.x.y`；`0.x` 仅保留在历史 changelog 中。Deck IR 的 `schemaVersion` 与 npm 包版本独立演进，旧 spec 由 `lib/office_migrations.py` 迁移。

### 9.2 Template Packs

`resources/template_packs.json` 注册内置模板包。每个包用 `themes`、`layouts`、`templates` 三个相对路径声明资产；`OFFICEFORGE_TEMPLATE_PACKS` 可追加外部包，值为目录列表（目录内放 `template_pack.json`）或 manifest 文件路径。`officeforge_deck` 的 `catalog` 返回每个包的 `themeCount`、`layoutCount`、`templateCount`；`plan`/`render` 接受 `templatePack` + `template`，由 Deck IR 统一解析，legacy 与 studio 后端都不需要知道包格式。`validate` 对内置与外部包逐项检查 manifest、资产文件、模板与 layout 引用，返回 `valid`、`errors`、`warnings` 与计数；未知包返回 `TEMPLATE_PACK_NOT_FOUND`。`render` 的 `engine: "auto"` 会按 deck 的 theme/template、table/chart/image/shape/background 与 Office-only 对象选择 legacy/studio/zbsph/guizang，并在 `adapterDecision` 中返回 `selected`、`requiredFeatures`、`unsupportedFeatures` 和原因。`verify` 先用宿主 `officeKit.renderImages` 把 `.pptx` 每页渲染成 PNG，再由 Python 侧检查页数、空页、图片尺寸、形状越界、标题覆盖与缺失字体，返回 `score`、`checks`、`issues`、`valid`；`report` 参数把同一份结果写成 JSON 或 HTML 报告（按扩展名或 `reportFormat` 决定格式）；`render` 传 `verify: true` 时会在写出 `.pptx` 后自动跑同一套渲染复核并把结果放入 `verification`，`verifyOutputDir` 可单独指定复核 PNG 目录。\n\nPPT 默认路由（`1.16.0` 起）：`Config.pptxEngine` 默认 `studio`，`officeforge_create` 的普通 `.pptx` 直接走设计化 dsh-ppt 渲染器，不再默认落回 legacy；`Config.beautifyEngine` 默认 `anon1009`，`officeforge_beautify` 默认调用内置 Anon1009 审美适配器（可用 `engine: \"builtin\"` 切回原主题/规范化路径）。`officeforge_edit` 仍执行结构性编辑，视觉重排请走 `officeforge_beautify`。

### 9.3 设计感知编辑

`officeforge_deck` 新增 `edit` 操作：`path` 为现有 `.pptx`（可省略并改用 `spec`），`output` 为输出 `.pptx`，`operations` 接受 `setTheme`、`setTitle`、`addSlide`、`removeSlide`、`replaceSlide`、`appendBlock`、`replaceBlock`、`setNotes`。Python 侧由 `office_deck_ir.apply_deck_operations` 在 Deck IR 上做 deepcopy 修改，再由 `office_deck.edit_deck` 统一走 adapter 渲染；非法 slide/block 索引返回 `DECK_INDEX_OUT_OF_RANGE`，不修改源文件。

### 9.4 插件页 UI 与发布

`lib/client.js` 只注入 `slots`，注册 `plugins.detail.section` 下的只读 `officeforge.status` 卡片；卡片提供两个只读动作：`GET /officeforge/doctor` 返回环境诊断，`GET /officeforge/capabilities` 返回 adapter、dsh-ppt 主题/母版/模板/版式计数与设计 token 目录。两个路由都由 `lib/index.js` 在 `ctx.inject(['webServer'])` 中注册，非 GET 请求返回 405，不写配置、不执行文档操作、不注册 prompt/LLM hook。

发布链见 `docs/PUBLISHING.md`：npm tarball、GitHub tag/release、`awesome-dsh-plugin` PR、`dsh-plugin-catalog` 索引，以及由 `scripts/build_offline.py` 生成的可离线安装 zip（含 PowerShell 安装/检查脚本和 SHA256SUMS）。`.github/workflows/release.yml` 在 tag 推送时构建 npm tarball 与 offline zip，并把两者附加到 GitHub release。

### 9.5 PPT 4–6 融合补强

`officeforge_deck` 的 `import` 把 `.pptx` 反向解析为 Deck IR：标题/段落/表格/图表/图片/形状保留几何与填充/线条/旋转，媒体复制到 workspace，批注保留位置，transition/animation 会记录为显式 block 或 slide 元数据；`fidelity.counts` 与 `warnings` 说明无法重建的对象。`AnonPptAdapter.create` 与 `AnonPptAdapter.edit` 现在真正执行 legacy 渲染后再走 Anon1009 审美 adapter，不再抛 `ADAPTER_UNSUPPORTED`；`office_deck_registry` 的 feature 集合与 `lossAudit` 会把 `smartart`/`animations`/`morph` 计入，legacy 作为高级 Office 对象的保真 owner。`office_webdeck` 新增 chart/shapes/media/comments 的 HTML 渲染。

## 10. 运行环境与移植性

DSH 版本 `0.2.2-rc.5` 上验证；Node 20+；Python 3.12.14（`python-docx 1.2.0`、`openpyxl 3.1.5`、`python-pptx 1.0.2`）。Python 发现顺序：`pythonPath` → `OFFICEFORGE_PYTHON` → `DSH_PYTHON` → `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe` → `$DSH_HOME/dsh-runtimes/*/dependencies/python` → `PATH`。

所有机器相关的路径、用户名与安装位置都通过环境变量或自动探测获得，仓库不包含任何个人路径。
