# Office Forge 设计说明

## 1. 目标与边界

Office Forge 是一个 **DSH Office 特化插件**，把常用的 Word / Excel / PowerPoint / PDF 文档能力收敛成十三个互相独立的工具。

- v0.14.0：插件页增加一键环境检查卡片（只读 GET `/officeforge/doctor`，`buildDoctorReport` 返回 health/summary/checks/remedy，客户端 `OfficeForgeStatus` 渲染按钮与结果）；v0.13.5：`officeforge_doctor` 增加可选 OCR 后端检查（`office_backends._ocr_status()`：PyMuPDF、pytesseract、Tesseract 可执行文件与语言包），缺失时给出安装提示，不把 OCR 定为必需依赖。
- v0.13.4：修复 `officeforge_preview` 在真实宿主的 `value is not lossless JSON` 失败——宿主 `snapshotJsonValue` 拒绝 `undefined` 成员，而 officeKit 的 `renderImages`/`convertTo` 只在实际适用时才带 `page`/`sheet`/`range`/`pageCount`；新增 `lossless()` 归一化层，在 `runEngine` 出口与各工具返回处丢弃缺省字段而非转发 `undefined`。
- v0.13.0：PPT 深化——`officeforge_create` 支持 `template`（10 套成品 deck）、`master`、版式库（38 个 layout，`resources/layouts.json`）；`officeforge_edit` 新增 `addComment`/`listComments`/`addMedia`/`applyMaster`，`setTransition` 支持 `morph`；`officeforge_audit` 新增 pptx `layouts`/`templates` 目录模式。
- v0.12.0：Word 深化——`officeforge_edit` 新增 `addCitation`/`addBibliography`/`setColumns`/`addWatermark`/`addContentControl`/`saveAsTemplate`；`officeforge_create` 的 docx spec 支持 `columns`、`watermark`、`contentControls`、`bibliography`，输出 `.dotx` 时写入模板内容类型。
- v0.11.0：性能与鲁棒性基线——新增 `benchmarks/run_benchmarks.py` 与 `benchmarks/baseline.json`，默认 12 个真实引擎用例覆盖 8,000 段 Word、12,000×10 Excel、120 页 PPT 及损坏/空/加密 PDF；完整回归把耗时/峰值内存阈值纳入门禁，并修复损坏/空 PDF 误报成功。
- v0.10.0：Excel 原生能力扩展——CSV/TSV/JSON 导入、数组公式（`ArrayFormula`）、组合图/次坐标轴（`type:"combo"`）、原生表格对象与保护继续保留；原生 PivotTable 仍通过 `officeforge_native` 的可选 COM 后端提供，文件级 `pivotSummary` 为汇总表。
- v0.9.0：PDF 文件级工具链扩展——页码、表单字段读取/填写、图片提取、OCR（可选 tesseract）、加密/解密、加密 PDF 密码错误分类；`officeforge_read` 支持加密 PDF 密码。
- v0.8.4：`officeforge_doctor` 做跨机自检——逐项检查 Python、python-docx/openpyxl/python-pptx、内置 pypdf、COM/pywin32、Office 应用注册、LibreOffice、officeKit、客户端 bundle；每项输出 required/optional、status、detail 与 remedy，且缺 Python 依赖时仍能返回诊断报告。
- v0.8.3：新增 Word 文档审计——`officeforge_audit` 的 docx `health`（标题跳级/重复标题/孤立标题/未清除占位符/超长段落/连续空段落）与 `profile`（块统计/标题层级/样式分布/节/表格/元数据），并接入 feature 开关 `word.audit`。
- v0.8.2：统一三个创建路径的文档核心属性（neutral author、当前 UTC 时间、revision 1，`spec.properties` 跨格式一致）。
- v0.8.1：安全加固——COM 进程回收、相对路径解析修正、fuzz 与回归闸门。
- v0.8.0：格式纵深——Word 脚注/书签/交叉引用，Excel 原生表格对象与工作表/工作簿保护，PowerPoint 演讲者备注批量与节。
- v0.7.0：插件页只读状态卡（`lib/client.js`）。
- v0.6.1：PDF 文件操作与可选原生 COM 后端——PDF 合并/拆分/旋转/水印、旧格式转换、COM 导出 PDF/重算/原生 PivotTable/原生 SmartArt；同时保留 `dsh-excel-chat` 的公式审计模型、`dsh-ppt` 的主题/版面模型与宿主 `officeKit` 的预览/重算能力。

它只做文档处理：

- 读取、创建、编辑 `.docx` / `.xlsx` / `.pptx`，只读解析 `.pdf`
- 用 JSON / CSV / TSV 数据填充模板
- 在一次受控调用中执行多个独立文档任务
- 通过宿主 `officeKit` 导出 PDF 等格式，并可选使用 `renderImages` / `recalculate`
- 文件级 PDF 合并/拆分/旋转/水印（内置 vendor pypdf）
- 通过可选 COM/LibreOffice 后端转换旧版 `.doc/.xls/.ppt`
- 通过可选 Microsoft Office COM Worker 执行导出 PDF、接受修订、重算、原生 PivotTable、原生 SmartArt
- 按 `modules` / `features` 配置启用或禁用 Word、Excel、PowerPoint 的细粒度能力

它不注册 prompt 变量、不接管 LLM stream、不注册 tool middleware 或 hook，因此不会改变 `@adwmc/helm-d`（破甲）的 `hcot`、`attack-scheduler`、`advisory` 行为，也不会与内置 `office_render` / `office_convert` / `office_recalculate` 重名。

## 2. 调研来源与能力映射

本项目参考了当前生态中下载量较高、维护活跃的 Office 相关包和 DSH 插件，只吸收“能力模型与工具设计”，不直接复制它们的运行时依赖：

| 来源 | 观察到的能力 | 在本插件中的落点 |
|---|---|---|
| `dsh-office-tools`（npm，周下载约 2.1k） | 8 个 `word_*` / `excel_*` / `ppt_*` 工具；`ctx.fs.resolve` + `contains` + `processPath` 路径策略 | 十三个 `officeforge_*` 工具的统一参数、路径授权、错误码 |
| `dsh-univer-office`（npm，周下载约 13k） | 统一 Office 运行时；文档富文本/列表/表格/图片/图表/页眉页脚，表格和幻灯片编辑，实时预览与 PDF 打印 | v0.2 的模块/功能分层、Word 图片/超链接/TOC、Excel 图表/条件格式/图片、PPT 形状/切换能力模型 |
| `@copylee/dsh-office` 1.2.2（Windows COM 单点插件） | Word 题注/脚注/引用/样式/页面设置/目录/更新字段/流程图；Excel 透视/高级图表/公式/条件格式/验证；PPT 主题/SmartArt/成品版式/文本框/形状 | v0.3 的 Word 题注/样式/字段、Excel 公式列/数据工具、PPT 文本框/连接线能力清单；仅参考能力，不引入 COM、SmartArt 或主题引擎 |
| `dsh-excel-chat` 0.43.0 | 25 个 Excel 工具：公式体检/编译/解释/追踪/修复、profile/semantic profile/insight、workbook 操作/任务/undo/diff、图表/透视；核心为纯 XML/exceljs，COM 仅用于图表/PDF/视觉 | v0.4 的 `officeforge_audit` Excel `health/profile/insight/diff`，不复制其 systemPrompt 注册，不引入 exceljs 依赖或 COM |
| `dsh-ppt` 0.33.0 | 12 个 PPT 工具：主题/模板/创建/编辑/检查/渲染/导入/吸收/引用；24 主题、6 母版骨架、35+ 版式 | v0.5 的 31 套外部化主题（含 dsh-ppt 24 套）、`deck` 版面检查、`applyTheme`；不引入其大依赖链或 LibreOffice 渲染 |
| `dsh-officecli-connection` | 包装本机 OfficeCLI 二进制，按 pptx/word/excel/表单/论文/数据面板/财务模型分层提供运行时技能 | 按文档族和场景拆分能力开关的设计参考；本插件不依赖外部二进制 |
| `docxtemplater`（npm，周下载约 1.8M） | `{{key}}` 占位符与循环模板 | `officeforge_template` 的 `{{key}}`、`{{#each items}}...{{/each}}` |
| `exceljs`（npm，周下载约 29.7M） | 工作簿/工作表模型、单元格写入、表格样式 | `office_xlsx.py` 的 sheets、columns、rows、样式与编辑操作 |
| `pptxgenjs`（npm，周下载约 7.3M） | 声明式 slide、layout、table、chart | `office_pptx.py` 的 slides、layouts、table、chart |
| `mammoth`（npm，周下载约 17.1M） | Office 文档转可读文本/markdown | `officeforge_read` 的 `markdown` / `text` / `outline` 输出 |
| `pdf-lib`（npm，周下载约 34.8M） | PDF 生成/处理的通用能力模型 | 导出不重复造轮子，统一委托宿主 `officeKit.convertTo()` |
| `DSH-Office`、`deepseek-harness-office`、`dsh-plugin-better-sidebar-plugin-office` 等 GitHub 插件 | DSH 工具注册、Office 读写/编辑/预览的工程结构 | `defineTool` + `ctx.effect` + 单工具隔离注册 |

> 版本与下载量为调研时（2026-11-13 ~ 2026-11-19）的 npm 周下载数据；它们只用于选型，不构成运行时依赖。`dsh-univer-office` 的完整运行时约 168 MB 且带 native bindings，因此没有直接嵌入。

## 3. 架构

```
DSH host
  └─ lib/index.js
       ├─ Config: modules + features + pythonPath / timeout / batch / export / enableBackends
       ├─ enabledTypes() / allowedExtensionsFor() / featuresFor()
       ├─ resolveTarget(): ctx.workingDirectory + ctx.fs.resolve/contains/processPath
       ├─ runEngine(): spawn 短生命周期 Python 子进程
       │    └─ lib/office_engine.py
       │         ├─ office_common.py    路径、原子写、占位符、feature gate
       │         ├─ office_docx.py      Word 读取/创建/编辑/模板
       │         ├─ office_xlsx.py      Excel 读取/创建/编辑/模板
       │         ├─ office_pptx.py      PowerPoint 读取/创建/编辑/模板/主题
       │         ├─ office_pdf.py       PDF 文本提取
       │         ├─ office_backends.py  可选 COM/LibreOffice 转换
       │         └─ office_migrations.py spec 版本迁移
       └─ officeforge_export / preview / recalculate → ctx.get("officeKit")
```

Node 层只做四件事：配置裁剪、路径策略、参数校验、子进程生命周期与 officeKit 导出。文档解析和生成全部在 Python 子进程中完成。Python 使用本机 DSH 捆绑运行时，不要求系统 Python 安装额外包。

## 4. 不写死、不递归的设计

这是针对之前“工具无限互相递归把程序搞崩”问题的硬约束：

1. **工具之间不互相调用**：每个 `officeforge_*` 工具只向 `office_engine.py` 发送一个 action。
2. **引擎 dispatch 只做一层**：`read/create/edit/template/batch/copy/probe` 均直接落到具体模块函数。
3. **batch 禁止嵌套**：`officeforge_batch` 的 item 只允许 `create` / `edit` / `template`，传 `action: "batch"` 直接返回 `INVALID_OPERATION`。
4. **模板循环只展开一层**：`{{#each items}}` 内部不再解析新的 `{{#each}}`，避免模板自引用导致无限展开。
5. **表格文本替换使用显式 stack**：`_replace_table_text` 不依赖 Python 递归，嵌套表格也不会触发递归深度问题。
6. **所有长度/深度有上限**：输入 50 MiB、输出 100 MiB、文本 200k 字符、Excel 5k×300、PPT 300 页、batch 20 项。

## 5. 不影响现有生态

- 插件入口 `apply(ctx, config)` 使用 `ctx.effect()` 注册并在返回的 disposer 中按逆序清理。
- 十三个工具逐个 `safeRegister()`；某个工具因宿主版本差异注册失败，只记录 warning，不会抛出到 profile 激活流程。
- 工具名全部以 `officeforge_` 开头，不与内置 `office_*` 冲突。
- `package.json` 使用官方 `dsh.bundle.patch` 机制，`cordis.patch.yml` 只插入一个最小 entry，所有行为通过 Config 覆盖。
- 不修改 DSH 安装目录、不写 profile、不安装依赖；交付物是可打包的插件源码。
- `officeforge_export` 使用 `subagents.rootOf(agent).session.id` 作为 `officeKit` 输出 owner，转换前后校验源文件 `stat.version`，远程 workspace 的 `processPathFromHostPath` 不可映射时返回 `OFFICEKIT_REMOTE_WORKSPACE`。
- `normalizeConfig()` 捕获非法 profile 配置并回退 `Config({})`；单工具注册失败只记录 warning，不会阻断 profile 激活。
- Excel 图表数据使用 `_officeforge_chart_data` / `_officeforge_chart_data_1` 等唯一隐藏工作表，不覆盖用户已有工作表。
- `tests/test_compatibility.mjs` 用真实 `assertSupportedJsonSchema` 检查十三个工具 schema、内置工具名冲突、模块全关仅保留 doctor、Word-only 类型枚举收窄，并静态确认无 prompt/LLM/middleware/hook 代码。
- 客户端 bundle `lib/client.js` 只注入 `slots`，只注册一个 `plugins.detail.section` 只读卡片（id `officeforge.status`，order 40）；不注册 `plugins.bundle.config`、不写配置、不发起文档操作。`dsh.client.platform` 固定为 `web`，`exports["./client"]` 指向该文件；`ClientModuleSystem.register` 的重复 id 保护与 `createClientModuleSystem` 的 boot manifest 解析已在 `tests/test_client_host.mjs` 中用真实宿主代码验证。
- 没有给 `Config` 加 `.volatile()`：工具在 `apply` 时静态注册，配置变更不会重新注册工具；若强行暴露可编辑表单，会出现“开关生效但工具未重注册”的不一致，因此插件页只提供只读状态与 `officeforge_doctor` 自检入口。

## 6. 模块与功能开关

`lib/index.js` 的 `Config` 提供三组开关：

- `modules`: `word` / `excel` / `ppt` / `pdf`，默认全部 `true`。关闭 Word/Excel/PPT 会从工具参数 enum 和注册集合中移除对应类型；关闭 `pdf` 会移除 `officeforge_export`。
- `enableBackends`: 是否注册 `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate` / `officeforge_native`；缺少可选后端时工具显式返回 `BACKEND_UNAVAILABLE`，核心 Native 路径不受影响。`officeforge_pdf` 只受 `modules.pdf` 控制，使用包内 vendor pypdf。
- `features`: 按文档族组织的高级能力，默认全部 `true`：
  - Word: `images`、`hyperlinks`、`pageBreaks`、`toc`、`captions`、`styles`、`fields`、`comments`、`revisions`、`mailMerge`
  - Excel: `charts`、`conditionalFormatting`、`dataValidation`、`comments`、`images`、`mergedCells`、`formulas`、`dataTools`、`audit`、`pivots`
  - PowerPoint: `charts`、`images`、`transitions`、`shapes`、`themes`、`audit`、`animations`、`smartart`

Node 层从 `resources/features.json` 读取模块/能力注册表并生成 Config schema；归一化后的 `features` 随每次 create/edit/batch 请求传给 Python 引擎；引擎内的 `require_feature()` 在能力被关闭时返回 `FEATURE_DISABLED`，不会静默降级。主题从 `resources/themes.json` 加载，新增主题不需要修改 Python 代码。

## 7. 本机环境适配

| 项目 | 值 |
|---|---|
| DSH | `0.2.2-rc.5` |
| Node | `v26.3.0`（宿主使用系统 `C:\\Program Files\\nodejs\\node.exe`） |
| 捆绑 Python | `C:\\Users\\zhoutao\\DSH 官方 0.2.2\\app\\profile\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe`，`3.12.14` |
| Python 包 | `python-docx 1.2.0`、`openpyxl 3.1.5`、`python-pptx 1.0.2`、`Pillow 12.3.0`、`lxml 6.1.3`、`XlsxWriter 3.2.9` |
| officeKit | `@deepseek-ai/dsh-office-kit 0.2.2-rc.5`，服务名 `officeKit`，提供 `convertTo/renderImages/convert/recalculate` |

Python 发现顺序：`pythonPath` → `OFFICEFORGE_PYTHON` → `DSH_PYTHON` → `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe` → 扫描 `$DSH_HOME/dsh-runtimes/*/dependencies/python` → `PATH`。探测命令只导入 `docx`、`openpyxl`、`pptx` 并打印哨兵值。

## 8. 验证矩阵

| 层 | 测试 | 结果 |
|---|---|---|
| Python 引擎 | `tests/test_engine.py`：probe、docx/xlsx/pptx 全流程、PDF 读取、schema 迁移、legacy convert 降级、PDF 合并/拆分/旋转/水印/页码/表单/图片/OCR/加密解密、native exportPdf、batch 隔离、copy、Word 图片/超链接/TOC/题注/样式/字段/批注/修订/邮件合并/脚注/书签/交叉引用/引文/参考文献/分栏/水印/内容控件/.dotx 模板/文档审计、Excel 图表/条件格式/数据验证/批注/合并/图片/公式列/排序/去重/空白填充/文本清洗/公式体检/profile/insight/diff/透视摘要/表格对象/保护/CSV-JSON 导入/数组公式/组合图次坐标轴、PPT 切换/形状/文本框/连接线/主题/版面质检/动画/SmartArt 风格图/备注批量/节/版式库/模板库/批注/媒体/母版/morph、feature 关闭拒绝 | `ALL TESTS PASSED`（216 项） |
| Node 插件层 | `tests/test_plugin.mjs`：inject/Config/十三工具注册、create/read、审计、PDF/native 入口、路径逃逸拒绝、officeKit export/convert/preview/recalculate、batch 单项隔离 | `PLUGIN TESTS PASSED`（52 项） |
| 客户端 bundle | `tests/test_client.mjs`：manifest 的 `dsh.client.platform`/`exports["./client"]`、bundle lazy-CJS 形态、`slots` 注入、状态卡 slot 注册 | `CLIENT UI TESTS PASSED` |
| 真实客户端宿主 | `tests/test_client_host.mjs`：用 `@deepseek-ai/dsh-client-modules/lib/client.js` 的 `createClientModuleSystem` 驱动自家 bundle，验证 factory 注册与 `apply` 激活 | `CLIENT HOST MACHINERY TESTS PASSED` |
| 一次性安装 | `tests/test_dsh_profile.mjs`：临时 `DSH_HOME` + 临时 profile 安装 tarball，校验 `dsh.profile.bundles`、`--dump-config`、`dsh.bundle.patch`、`dsh.client`/`exports["./client"]`，结束后清理 | `DSH PROFILE TESTS PASSED` |
| 配置层 | `tests/test_plugin_config.mjs`：modules/features 归一化、v0.3/v0.4/v0.5 feature 默认值、类型裁剪、export 跳过、禁用 TOC 拒绝 | `PLUGIN CONFIG TESTS PASSED`（14 项） |
| 兼容性/隔离 | `tests/test_compatibility.mjs`：真实 schema 校验、十三工具注册、内置工具名不覆盖、模块全关仅保留 doctor、Word-only 枚举、无 prompt/LLM/middleware/hook | `COMPATIBILITY TESTS PASSED`（48 项） |
| 静态检查 | `python -m py_compile` 八个模块、`node --check lib/index.js lib/client.js` | 通过 |
| 打包 | `npm pack`：v0.14.0（含 `resources/`、`lib/vendor/`、`lib/client.js`、`docs/`），最终 SHA-256 见交付说明 | 通过 |

## 9. 安装状态

当前**没有**安装到任何 DSH profile，也没有修改 `C:\\Users\\zhoutao\\DSH 官方 0.2.2\\app\\profile`。`README.zh.md` 的安装章节只是未来操作说明；交付阶段不执行任何安装命令。
## Safety hardening (v0.8.1)

- **COM process reclamation**: the isolated worker records every Office
  automation instance it starts and quits them all on exit. A refused `Quit()`
  is escalated to `taskkill /F /T` by resolved PID, and `run_com` reclaims the
  whole worker tree on timeout, so no `WINWORD/EXCEL/POWERPNT` orphan survives a
  failed conversion.
- **Workspace-relative resolution**: `_checked` resolves relative paths against
  `OFFICEFORGE_WORKSPACE`, never the engine's own cwd, so a missing file reports
  `INPUT_NOT_FOUND` instead of a misleading `PATH_OUTSIDE_WORKSPACE`.
- **Classified failures only**: `tests/test_fuzz.py` proves that empty payloads,
  unknown actions, type mismatches, truncated JSON, nested batches, oversized
  `schemaVersion` and traversal attempts all return one JSON object with a
  stable error code, no traceback and no file outside the workspace.
- **Regression gate**: `tests/run_regression.py` enforces per-suite assertion
  floors from `tests/regression_baseline.json` plus each suite's completion
  marker, so a suite that silently emits zero assertions fails.

