# Office Forge

**面向 DeepSeek Harness（DSH）的 Office 文档自动化工具箱。**

[![npm](https://img.shields.io/npm/v/dsh-office-forge.svg)](https://www.npmjs.com/package/dsh-office-forge)
[![GitHub release](https://img.shields.io/github/v/release/siyu0209/dsh-office-forge.svg)](https://github.com/siyu0209/dsh-office-forge/releases/latest)

**安装：** `dsh plugin --profile <profile> add dsh-office-forge` · [npm](https://www.npmjs.com/package/dsh-office-forge) · [GitHub Release](https://github.com/siyu0209/dsh-office-forge/releases/latest)

Office Forge 是一个自包含的 DSH 插件包，为 Agent 提供十六个互相独立的工具，用来处理 Word、Excel、PowerPoint 和 PDF 文件。它可以读取、创建、编辑、审计、预览、转换、批量处理和填充模板，不需要启动 Office 界面。所有文档处理都在短生命周期的 Python 子进程中完成；Node 层只负责工作区授权、参数校验、子进程传输和可选的宿主集成。

插件为独立实现，不捆绑任何第三方 Office 运行时，也不依赖其他文档插件。

## 主要特性

- **一套工具覆盖四种格式**：`.docx`、`.xlsx`、`.pptx`、`.pdf`。
- **默认文件级处理**：直接处理磁盘上的文件，不要求安装 Microsoft Office。
- **可选真实后端**：只有检测到 Microsoft Office COM 或 LibreOffice 时才使用。
- **模板友好**：支持 `{{key}}` 占位符和一层 `{{#each items}}...{{/each}}` 表格/工作表行循环。
- **内置审计**：Word 标题质量、Excel 公式健康/画像/洞察/差异、PowerPoint 页面质量与主题目录。
- **适合 Agent 调用**：所有工具接收并返回无损 JSON，提供稳定错误码。
- **工作区安全**：词法路径检查 + 宿主路径检查、原子写入、输入输出限额、无递归。
- **可配置**：Word、Excel、PowerPoint、PDF 的模块与细粒度功能开关。
- **内置 `officeforge-ppt` 技能**：把 PPT 制作/修改的询问规则、样张流程和工具路由随插件分发；不固定问卷，只问会影响制作的关键问题。
- **可诊断**：`officeforge_doctor` 和插件页只读卡片会明确报告缺少哪些可选后端。

## 环境要求

| 组件 | 要求 |
|---|---|
| DSH | `>= 0.2.2-rc.5` |
| Node.js | `>= 20` |
| Python | Python 3.12，安装 `python-docx`、`openpyxl`、`python-pptx`；DSH 捆绑运行时已包含 |
| PDF | 插件内置 `pypdf`，不需要外部 PDF 服务 |
| Microsoft Office COM | 可选，仅 Windows；用于原生 PivotTable、原生 SmartArt、Word 域/目录更新、接受修订、PDF 导出、旧格式转换 |
| LibreOffice | 可选；COM 不可用时用于旧 `.doc/.xls/.ppt` 转换 |
| OCR | 可选，需要 PyMuPDF、pytesseract 和 Tesseract 可执行文件；`officeforge_doctor` 会报告缺失项 |

## 安装

### 从 npm 安装

包已发布在 <https://www.npmjs.com/package/dsh-office-forge>。

```bash
# 安装到 DSH profile
dsh plugin --profile <profile> add dsh-office-forge

# 固定版本
dsh plugin --profile <profile> add dsh-office-forge@3.4.1
```

### 从 GitHub Releases 安装

```bash
dsh plugin --profile <profile> add https://github.com/siyu0209/dsh-office-forge/releases/download/v3.4.1/dsh-office-forge-3.4.1.tgz
```

### 从 DSH 插件页安装

打开 **插件 → 安装**，输入 `dsh-office-forge`，或把 Release 里的 `.tgz` 拖进安装器，然后重启 DSH。桌面版 `desktop` profile 由 Electron 管理，请通过插件页安装该 profile。

### 从本地 tarball 安装

```bash
dsh plugin --profile <profile> add /path/to/dsh-office-forge-3.4.1.tgz
```

也可以按 DSH 的常规方式把包加入 profile 的 dependencies 与 `dsh.profile.bundles`。插件包声明了 `dsh.bundle.patch`、`dsh.client`、locale 和 icon，可被 DSH 插件页识别。

### 从源码构建

```bash
git clone https://github.com/siyu0209/dsh-office-forge.git
cd dsh-office-forge
npm pack
```

然后把生成的 `dsh-office-forge-*.tgz` 安装到目标 profile。源码运行需要 Node 和带三个 Office 库的 Python；`officeforge_doctor` 会指出缺少什么。

### 验证安装

重启 DSH 后调用：

```json
{ "tool": "officeforge_doctor" }
```

必需组件齐全时返回 `ok: true`。OCR、LibreOffice 等可选后端可能显示 `warn`，不会影响核心文档工具。

## 快速开始

### 创建 Word 文档

```json
{
  "path": "reports/weekly.docx",
  "type": "docx",
  "spec": {
    "title": "周报",
    "subtitle": "第 47 周",
    "sections": [
      { "heading": "进展", "bullets": ["完成模块 A", "验证流程 B"] },
      { "heading": "下周", "numbered": ["修复 C", "发布 D"] }
    ],
    "pageNumbers": true
  }
}
```

### 创建 Excel 工作簿

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "spec": {
    "sheets": [
      {
        "name": "指标",
        "columns": ["月份", "收入"],
        "rows": [["1月", 120], ["2月", 180], ["3月", 240]],
        "freeze": "A2",
        "autoFilter": true,
        "charts": [
          { "type": "bar", "dataRange": "B1:B4", "categoriesRange": "A2:A4", "anchor": "E2", "title": "收入" }
        ]
      }
    ]
  }
}
```

### 创建 PowerPoint 演示文稿

```json
{
  "path": "decks/quarterly.pptx",
  "type": "pptx",
  "spec": {
    "theme": "swiss",
    "slides": [
      { "layout": "cover", "title": "季度回顾", "subtitle": "FY2026" },
      { "layout": "kpi", "title": "亮点", "items": [
        { "value": "42%", "label": "增长" },
        { "value": "18", "label": "新客户" }
      ] }
    ]
  }
}
```

### 填充 Word 模板

```json
{
  "template": "templates/people.docx",
  "output": "out/people.docx",
  "data": {
    "title": "团队",
    "people": [
      { "name": "张三", "dept": "研发" },
      { "name": "李四", "dept": "产品" }
    ]
  },
  "options": { "expandLoops": true, "missing": "keep" }
}
```

### 审计工作簿

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "mode": "health"
}
```

### 合并 PDF

```json
{
  "operation": "merge",
  "inputs": ["out/a.pdf", "out/b.pdf"],
  "output": "out/merged.pdf"
}
```

## 工具参考

| 工具 | 作用 |
|---|---|
| `officeforge_read` | 读取 `.docx` / `.xlsx` / `.pptx` / `.pdf`，支持 `content`、`outline`、`search` 模式，输出 markdown、text 或 JSON；PDF 支持文本提取与加密密码。 |
| `officeforge_create` | 用声明式 spec 创建文档。Word 支持章节、标题、项目符号、编号、表格、页眉页脚、页面设置、页码、目录、图片、超链接、题注、样式、批注、修订、邮件合并、脚注、书签、交叉引用、引文、参考文献、分栏、水印、内容控件、`.dotx` 模板。Excel 支持工作表、CSV/TSV/JSON 导入、公式、数组公式、图表（含组合图）、条件格式、数据验证、批注、图片、合并单元格、表格对象、保护。PowerPoint 支持幻灯片、版式、模板、主题、图表、图片、表格、文本框、形状、连接线、切换（含 morph）、动画、SmartArt 风格图、备注、节、批注、媒体、母版。 |
| `officeforge_edit` | 对现有文档执行有界编辑。Word：替换、追加、图片、目录、题注、样式、字段、批注、修订、邮件合并、脚注、书签、交叉引用、引文、参考文献、分栏、水印、内容控件、模板、核心属性。Excel：单元格、行、工作表、格式、冻结、筛选、图表、条件格式、数据验证、批注、合并、图片、公式列、排序、去重、空白填充、文本清洗、透视摘要、表格对象、保护。PowerPoint：文本、幻灯片、备注、节、切换、形状、文本框、连接线、动画、SmartArt 风格图、批注、媒体、母版、主题、核心属性。 |
| `officeforge_template` | 用 JSON 或 `.json/.csv/.tsv` 数据填充 Word/Excel 模板：`{{key}}` 占位符和一层 `{{#each items}}...{{/each}}` 循环。 |
| `officeforge_audit` | 只读审计。Word：`health`（标题跳级、重复/孤立标题、未清除占位符、超长段落、连续空段落）与 `profile`。Excel：`health`（公式错误、模式不一致、硬编码、空行断链、循环引用）、`profile`、`insight`、`diff`。PowerPoint：`deck` 质量、`aesthetic` 审美评分、`profile`、`themes`、`layouts`、`templates`。 |
| `officeforge_beautify` | 原地美化现有 `.pptx`：默认使用内置 Anon1009 审美适配器（主题配色、字体层级、对比度、对齐、表格/页眉强调、页码），返回美化前后评分；`engine: "builtin"` 可切回原 Office Forge 主题/规范化路径。 |
| `officeforge_web` | 用同一份 slides spec 生成自包含 `.html` 演示：支持键盘/按钮翻页和打印样式，不依赖外部运行时。 |
| `officeforge_deck` | Deck IR 编排入口：`plan` 规范化并统计 deck；`storyboard` 生成带备注与时长估算的演示提纲（JSON 或 Markdown）；`render` 通过 `legacy`、`studio` 或可选 `zbsph`（完整 PPTD 文本/形状/线条/图片/表格/图表/背景/备注 schema，并支持 Anon1009 模板）渲染 `.pptx`；`engine: "auto"` 会检查 Deck IR、自动选择兼容渲染器并返回 `adapterDecision`；`audit`/`repair` 使用 `legacy`、内置 `anon1009` 审美 adapter 或内置 `guizang` design tokens；`composer` 通过内置 web 渲染器输出单文件 HTML；`catalog` 查看版式/模板/模板包/adapter 及安装提示；`web` 直接输出 HTML；`validate` 校验指定或全部 Template Pack；`verify` 通过宿主 officeKit 把 `.pptx` 真实渲染为 PNG，再检查页数、空白页、尺寸、形状越界与标题覆盖，返回 `score`/`checks`/`issues`；传 `report` 可把结果写成 JSON 或 HTML 报告，`render` 传 `verify: true` 会自动跑同一套检查；`designCatalog` 列出 Anon1009/歸藏/Composer/Office 设计 token；`import` 把现有 `.pptx` 解析成 Deck IR，并返回保真计数与警告；`edit` 对导入或 spec deck 应用设计感知的 Deck IR 操作（`setTheme`/`setTitle`/`addSlide`/`removeSlide`/`replaceSlide`/`appendBlock`/`replaceBlock`/`setNotes`）并重新渲染。 |
| `officeforge_doctor` | 只读运行时诊断：Python 与 Office 库、内置 PDF 引擎、COM/pywin32、已注册 Office 应用、LibreOffice、`officeKit`、客户端 bundle、当前 modules/features。每项包含 `required`、`status`、`detail`、`remedy`。 |
| `officeforge_convert` | 通过 COM/LibreOffice 转换旧 `.doc/.docm/.xls/.xlsm/.ppt/.pptm`；通过宿主 `officeKit` 转换现代格式/PDF。 |
| `officeforge_pdf` | 文件级 PDF 操作：`merge`、`split`、`rotate`、`watermark`、`pageNumbers`、`formFields`、`fillForm`、`extractImages`、`ocr`、`encrypt`、`decrypt`；使用内置 pypdf。 |
| `officeforge_native` | 可选 Microsoft Office COM 操作：`convert`、`exportPdf`、`acceptRevisions`、`updateFields`（Word 域和目录更新）、`recalculate`、`pivot`（原生 PivotTable）、`smartArt`（原生 SmartArt）。 |
| `officeforge_preview` | 通过 `officeKit.renderImages` 把页面、幻灯片或工作表区域渲染为 PNG，不启动 Office UI。 |
| `officeforge_recalculate` | 通过 `officeKit.recalculate` 重算公式并保存刷新后的工作簿。 |
| `officeforge_batch` | 在一次 Python 调用中执行最多 20 个独立的 create/edit/template 任务；每项独立成功或失败；禁止嵌套 batch。 |
| `officeforge_export` | 通过 `officeKit` 转换为 PDF 或其他支持格式；可选择把结果复制回工作区。 |

## 配置

所有字段都可选，所有模块与功能默认 `true`；空配置即启用完整能力。

| 字段 | 默认值 | 说明 |
|---|---|---|
| `pythonPath` | 自动探测 | 带 `python-docx`、`openpyxl`、`python-pptx` 的 Python 可执行文件。 |
| `engineTimeoutMs` | `120000` | 单次 Python 引擎调用超时（毫秒）。 |
| `maxBatchItems` | `20` | 单次 `officeforge_batch` 的最大条目数。 |
| `enableExport` | `true` | 是否注册 `officeforge_export`。 |
| `enableBatch` | `true` | 是否注册 `officeforge_batch`。 |
| `enableBackends` | `true` | 是否注册 `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate`；缺少后端时显式报错。 |
| `modules` | 全部 `true` | `word` / `excel` / `ppt` / `pdf` 文档族开关。 |
| `features` | 全部 `true` | Word / Excel / PowerPoint 细粒度能力开关。 |

profile 配置示例：

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

完整功能名见 `resources/features.json`，主题预设见 `resources/themes.json`。

模板包：内置包声明在 `resources/template_packs.json`（`officeforge-core`、`officeforge-academic`、`officeforge-pitch`）。要添加本地模板包而不改包体，把 `OFFICEFORGE_TEMPLATE_PACKS` 指向包含 `template_pack.json`、`themes.json`、`layouts.json`、`templates.json` 的目录（或直接指向 manifest）；`officeforge_deck` 的 `operation: "catalog"` 会列出每个包及其资产数量。

### Python 发现顺序

1. 配置项 `pythonPath`
2. 环境变量 `OFFICEFORGE_PYTHON`
3. 环境变量 `DSH_PYTHON`
4. `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe`
5. `$DSH_HOME/dsh-runtimes/*/dependencies/python`
6. `PATH` 上的 `python3` / `python`

## 架构

```
DSH 宿主
  └─ lib/index.js                 Node 宿主层
       ├─ Config schema           modules + features + 运行时选项
       ├─ 路径策略                ctx.fs.resolve / contains / processPath
       ├─ runEngine()             短生命周期 Python 子进程
       │    └─ lib/office_engine.py
       │         ├─ office_common.py      路径、原子写、占位符、feature gate
       │         ├─ office_docx.py        Word 读取/创建/编辑/模板/审计
       │         ├─ office_xlsx.py        Excel 读取/创建/编辑/模板/审计
       │         ├─ office_pptx.py        PowerPoint 读取/创建/编辑/模板/审计
       │         ├─ office_pdf.py         PDF 文本提取
       │         ├─ office_pdf_tools.py   PDF 合并/拆分/旋转/水印/表单/图片/OCR/加密
       │         ├─ office_backends.py    可选 COM / LibreOffice 后端
       │         ├─ office_com_worker.py  隔离的 COM Worker
       │         └─ office_migrations.py  spec schema 迁移
       ├─ officeKit 桥接          export / preview / recalculate
       └─ lib/client.js           插件页只读状态卡 + 环境检查 + 能力矩阵
```

Node 层只做四件事：配置裁剪、工作区路径策略、参数校验、子进程生命周期。所有文档解析与生成都在 Python 中完成；Python 引擎每次调用都是无状态的，每个工具只分发一个 action。

## 安全与隔离

- **只写工作区**：每个输入/输出路径同时经过宿主 `ctx.fs.resolve` / `ctx.fs.contains` 和词法 `PATH_OUTSIDE_WORKSPACE` 检查。
- **原子写入**：新文件写入同目录临时文件后 `os.replace`；默认拒绝覆盖，`overwrite: true` 才替换；可选生成同目录备份。
- **有界执行**：输入 50 MiB、输出 100 MiB、文本 200,000 字符、Excel 单表 5,000×300、PPT 300 页、batch 20 项。
- **无递归**：工具之间不调用；每个工具只向引擎发送一个 action；模板只展开一层循环；禁止嵌套 batch；表格文本替换使用显式 stack。
- **不注册宿主钩子**：只注入 `tools`、`fs`、`workingDirectory`；不注册 `systemPrompt`、LLM stream hook、tool middleware 或宿主 hook。
- **不冲突工具名**：所有工具都以 `officeforge_` 开头；内置 `office_*` 与 `run_code` 不受影响。
- **单工具隔离**：逐个注册；单个工具因宿主版本差异失败只记录 warning，不会中断插件或 profile 激活。
- **无损 JSON**：所有工具结果会归一化，缺省可选字段不会以 `undefined` 泄漏，避免宿主报 `value is not lossless JSON`。
- **插件页只读**：客户端 bundle 只注册 `plugins.detail.section` 只读状态卡，不写配置、不执行文档操作。

## 错误模型

所有失败都返回带 `code`、`message` 和可选 `detail` 的稳定 JSON 错误对象。常见错误码：

| 错误码 | 含义 |
|---|---|
| `INPUT_NOT_FOUND` | 源文件不存在或不在工作区。 |
| `PATH_OUTSIDE_WORKSPACE` | 路径越出允许的工作区。 |
| `UNSUPPORTED_FORMAT` | 文件扩展名不能被请求的工具处理。 |
| `OUTPUT_EXISTS` | 目标已存在且未请求 `overwrite`。 |
| `FEATURE_DISABLED` | 请求的能力在 `features` 中被关闭。 |
| `BACKEND_UNAVAILABLE` | 可选 COM/LibreOffice/OCR 后端未安装或未检测到。 |
| `SCHEMA_VERSION_UNSUPPORTED` | Spec 的 schema 版本高于引擎支持范围。 |
| `INVALID_PDF` / `PDF_PASSWORD_REQUIRED` / `PDF_PASSWORD_INVALID` | PDF 损坏、需要密码或密码错误。 |
| `ENGINE_BAD_OUTPUT` | Python 引擎返回无法无损序列化的值。 |

## 测试与性能基线

仓库包含十一套测试和一个基准套件。全部使用临时目录和 mock 宿主，不启动 DSH，也不写 profile。

| 套件 | 用途 |
|---|---|
| `tests/test_engine.py` | Python 引擎端到端：全格式、模板、审计、PDF、native、batch、copy。 |
| `tests/test_fuzz.py` | 畸形 JSON、路径逃逸、未知 action、嵌套 batch、稳定错误码。 |
| `tests/test_com_cleanup.py` | 可选 COM 进程回收测试；无 Office/pywin32 时跳过。 |
| `tests/test_plugin.mjs` | Node 插件层，使用 mock 宿主和 mock `officeKit`。 |
| `tests/test_plugin_config.mjs` | modules/features 归一化与类型裁剪。 |
| `tests/test_compatibility.mjs` | 真实宿主 schema 校验、工具名冲突、隔离静态检查。 |
| `tests/test_client.mjs` | 客户端 bundle manifest、lazy-CJS 形态、slot 注册。 |
| `tests/test_client_host.mjs` | 用真实 `dsh-client-modules` 驱动客户端 bundle。 |
| `tests/test_ppt_bridges.mjs` | 真实 dsh-ppt/studio/zbsph 桥接矩阵：24 主题、8 master、20 模板、82 版式、4 套 Vate 模板。 |
| `tests/test_dsh_profile.mjs` | 隔离 `DSH_HOME` 的一次性安装验证。 |
| `tests/run_regression.py` | 按 `tests/regression_baseline.json` 运行全部套件；`--quick` 跳过可选套件。 |
| `benchmarks/run_benchmarks.py` | 13 个性能与鲁棒性基准；阈值在 `benchmarks/baseline.json`。 |

### 本地验证

把 `PYTHON` 指向带 `python-docx`、`openpyxl`、`python-pptx` 的 Python 3.12。如果使用 DSH 捆绑运行时：

```bash
export DSH_HOME="${DSH_HOME:-$HOME/.dsh}"
export PYTHON="$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe"
```

Windows PowerShell：

```powershell
$env:DSH_HOME = "$env:USERPROFILE\.dsh"
$env:PYTHON = "$env:DSH_HOME\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
```

运行全部：

```bash
"$PYTHON" tests/run_regression.py
"$PYTHON" tests/run_regression.py --quick
```

单独运行：

```bash
"$PYTHON" tests/test_engine.py
"$PYTHON" tests/test_fuzz.py
"$PYTHON" tests/test_com_cleanup.py
node tests/test_plugin.mjs
node tests/test_plugin_config.mjs
node tests/test_compatibility.mjs
node tests/test_client.mjs
node tests/test_client_host.mjs
node tests/test_ppt_bridges.mjs
"$PYTHON" benchmarks/run_benchmarks.py
node tests/test_dsh_profile.mjs
```

`tests/test_client_host.mjs` 和 `tests/test_dsh_profile.mjs` 需要 DSH 安装；用 `DSH_APP_ROOT` 指向包含 `node_modules/@deepseek-ai/dsh` 的目录。找不到时会跳过。

## 兼容性与限制

- **DSH**：在 `0.2.2-rc.5` 上测试；`package.json` 声明 `engines.dsh >= 0.2.2-rc.5`。
- **Node**：Node 20+；DSH 宿主当前使用 Node 26。
- **Python**：CPython 3.12.14；`python-docx 1.2.0`、`openpyxl 3.1.5`、`python-pptx 1.0.2`。
- **Office COM**：仅 Windows + Microsoft Office + pywin32；不可用时 `officeforge_native` 返回 `BACKEND_UNAVAILABLE`，核心文件级工具仍可用。
- **LibreOffice**：可选，仅用于缺少 COM 时的旧格式转换。
- **OCR**：可选；需要 PyMuPDF、pytesseract、Tesseract 可执行文件和语言包。
- **预览**：使用宿主 `officeKit.renderImages`；像素尺寸和字体取决于宿主与已安装字体。
- **原生功能**：原生 PivotTable/PivotChart/Slicer 与原生 SmartArt 需要 COM；纯文件路径提供透视汇总和 SmartArt 风格的可编辑形状，不是完整原生对象。
- **Word 水印**：VML 水印在 Word 中可见；部分 LibreOffice 预览可能不显示。
- **宏**：不执行宏；在 openpyxl 支持范围内编辑 `.xlsm` 时保留 VBA 容器。

## 故障排查

| 现象 | 检查 |
|---|---|
| `officeforge_doctor` 显示 `BACKEND_UNAVAILABLE` | 可选后端未安装；核心工具仍可使用。 |
| 相对路径报 `INPUT_NOT_FOUND` | 路径按会话工作区解析；请使用工作区内路径。 |
| `PATH_OUTSIDE_WORKSPACE` | 目标路径越出工作区；写入工作区内或使用允许的宿主路径。 |
| `FEATURE_DISABLED` | 对应功能在 `features` 中关闭；在 bundle 配置中启用。 |
| `value is not lossless JSON` | 工具返回了缺省可选字段；升级到使用 `lossless()` 归一化结果的版本。 |
| Python import 错误 | 运行 `officeforge_doctor`；把 `pythonPath` 或 `OFFICEFORGE_PYTHON` 指向带 `python-docx`、`openpyxl`、`python-pptx` 的 Python。 |
| COM 输出失败 | 确认 Windows Office、pywin32 可用，且没有受保护视图弹窗阻止文件。 |

## 仓库结构

```
office-forge/
  package.json                 # DSH bundle 与 npm 元数据
  cordis.patch.yml             # 把 dsh-office-forge 插入 profile
  icon.svg
  locale/
    zh.json
    en.json
  resources/
    themes.json                # 31 套内置主题
    features.json              # 模块/功能注册表
    layouts.json               # 38 个 PowerPoint 版式
    templates.json             # 10 套 deck 模板
  lib/
    index.js                   # 工具注册、路径策略、Python 传输、officeKit 桥接
    client.js                  # 只读插件页状态卡、环境检查与能力矩阵
    office_engine.py           # JSON stdin/stdout 引擎入口
    office_common.py           # 路径、原子写、占位符、校验
    office_pdf.py              # PDF 文本提取
    office_pdf_tools.py        # PDF 文件操作
    office_backends.py         # 可选 COM / LibreOffice 后端
    office_com_worker.py       # 隔离 COM Worker
    office_migrations.py       # spec schema 迁移
    office_docx.py             # Word
    office_xlsx.py             # Excel
    office_pptx.py             # PowerPoint
    vendor/pypdf/              # 内置 pypdf
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

## 开发

插件是纯 ESM + Python 3.12，没有构建步骤。

```bash
git clone https://github.com/siyu0209/dsh-office-forge.git
cd dsh-office-forge
npm install --ignore-scripts
```

提交 PR 前请运行上表的验证命令。功能变更要同步 `tests/test_engine.py`，只有确认新的断言下限合理时才更新 `tests/regression_baseline.json`。

## 版本策略

语义化版本：patch 修复 bug 与兼容性；minor 增加能力；major 可能改变文档 spec 语义。每次发布记录在 `docs/CHANGELOG.md`，包含迁移规则。

## 许可证

MIT，见 [LICENSE](LICENSE)。
