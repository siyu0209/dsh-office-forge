# Office Forge 文档锻造台

Office Forge 是一个面向 DeepSeek Harness（DSH）的 Office 文档插件。它把 Word、Excel、PowerPoint、PDF 的读写、创建、编辑、模板填充、审计、批量处理、PDF 合并/拆分/旋转/水印、旧格式转换、视觉预览、公式重算、PDF 导出和可选 Microsoft Office COM 原生操作集中到一组十三个独立的 `officeforge_*` 工具里，所有文档逻辑由短生命周期的 Python 子进程执行，Node 层只负责路径授权、参数校验、子进程调度和宿主 `officeKit` 导出。

> 当前仓库包含插件源码与离线测试。安装到 DSH profile 是可选项，步骤见「安装」一节。

## 工具

| 工具 | 作用 |
|---|---|
| `officeforge_read` | 读取 `.docx` / `.xlsx` / `.pptx` / `.pdf`，Office 支持 `content`、`outline`、`search` 三种模式，输出 markdown / text / json；PDF 支持文本/markdown 提取。 |
| `officeforge_create` | 用声明式 spec 新建文档；docx 用 `title/subtitle/sections`，xlsx 用 `sheets`（或 `import` 从 CSV/TSV/JSON 导入），pptx 用 `slideSize/slides`，也可只给 `template`（board/project/pitch/academic/training/sales/report/product/strategy/minimal）生成整套成品 deck，或用 `master` 设置母版背景与页眉页脚。Excel 图表支持 `combo` 主/次坐标轴组合图；PPT 的 `layout` 支持 38 个版式库（kpi/timeline/cards/comparison/matrix/quote/image-* 等）。 |
| `officeforge_edit` | 对现有文档执行有界操作：docx `replace/append/addImage/addToc/addCaption/setStyle/updateFields/addComment/trackRevision/mailMerge/addFootnote/addBookmark/addCrossReference/addCitation/addBibliography/setColumns/addWatermark/addContentControl/saveAsTemplate/setCore`；xlsx `setCells/appendRows/replace/createSheet/deleteSheet/renameSheet/setColumnWidths/setFreeze/setAutoFilter/addChart/addConditionalFormatting/addDataValidation/addComment/mergeCells/addImage/addFormulaColumn/sortRange/removeDuplicates/fillBlanks/cleanRange/pivotSummary/addTable/protectSheet/protectWorkbook`；pptx `replace/appendSlides/setNotes/setNotesBatch/addSection/listSections/deleteSlide/moveSlide/setTransition/addShape/addTextBox/addConnector/addAnimation/addSmartArt/addComment/listComments/addMedia/applyMaster/applyTheme/setCore`。默认生成同目录备份。 |
| `officeforge_template` | 用 JSON 或 `.json/.csv/.tsv` 数据填充模板：`{{key}}` 占位符，Word/Excel 额外支持一层 `{{#each items}}...{{/each}}` 表格行/工作表行循环。 |
| `officeforge_audit` | 只读审计：Word 支持 `health`（标题跳级、重复标题、孤立标题、未清除占位符、超长段落、连续空段落）与 `profile`；Excel 支持 `health`（公式错误、同列公式模式不一致、公式列硬编码、空行断链、循环引用）、`profile`、`insight`、`diff`；PowerPoint 支持 `deck` 质量检查、`profile`、`themes` 主题列表、`layouts` 版式目录、`templates` 模板目录（后三者可省略 `path`）。 |
| `officeforge_doctor` | 只读运行时诊断：检查 Python 与 python-docx/openpyxl/python-pptx、内置 PDF 引擎、COM/LibreOffice 后端、Office 应用注册、`officeKit` 服务、客户端 bundle 以及当前 modules/features 配置；每项标注 required/optional、当前状态和缺失时的补救建议。安装后先运行它即可判断哪些能力可用。 |
| `officeforge_convert` | 旧格式 `.doc/.docm/.xls/.xlsm/.ppt/.pptm` 通过可选 COM/LibreOffice 后端转换；现代格式与 PDF 走宿主 `officeKit`。 |
| `officeforge_pdf` | 文件级 PDF 工具：`merge` 合并、`split` 拆分、`rotate` 旋转、`watermark` 水印、`pageNumbers` 页码、`formFields` 表单字段、`fillForm` 填表、`extractImages` 提取图片、`ocr` OCR、`encrypt`/`decrypt` 加密解密；使用包内 vendor 的 pypdf，不依赖外部 PDF 服务。 |
| `officeforge_native` | 可选 Microsoft Office COM 原生操作：`convert`、`exportPdf`、`acceptRevisions`、`recalculate`、`pivot`（原生 PivotTable）、`smartArt`（原生 SmartArt）；需要 Windows + Office + pywin32。 |
| `officeforge_preview` | 通过宿主 `officeKit.renderImages` 渲染页面/幻灯片/工作表区域为 PNG，不启动 Office UI。 |
| `officeforge_recalculate` | 通过宿主 `officeKit.recalculate` 做真实公式重算并保存新工作簿，不使用公式文本近似。 |
| `officeforge_batch` | 在一次 Python 调用里执行多个独立的 create/edit/template 任务；每项单独返回成功或失败，互不回滚、不允许嵌套。 |
| `officeforge_export` | 通过宿主 `officeKit` 将文档转换为 PDF 等格式；可选把结果复制回工作目录。 |

## 模块与功能开关

Office Forge 允许按文档族和高级功能做选择，配置写在 DSH profile 的 bundle config 中；所有开关默认开启，不配置时全部能力可用；`resources/features.json` 和 `resources/themes.json` 已外部化。

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

- 关闭 `modules.word` / `modules.excel` / `modules.ppt` 会分别移除对应文档类型和工具入口；关闭 `modules.pdf` 会移除 `officeforge_export`。
- 关闭某个 `features` 项后，引擎收到该能力请求会返回 `FEATURE_DISABLED`，不会静默降级或写入不完整结果。
- 高级能力：Word 图片/超链接/分页/目录/题注/样式/字段刷新/批注/修订/邮件合并/脚注/书签/交叉引用；Excel 图表/组合图次坐标轴/条件格式/数据验证/批注/图片/合并单元格/公式列/数组公式/排序/去重/空白填充/文本清洗/公式健康审计/透视摘要/CSV/TSV/JSON 导入/原生表格对象/工作表与工作簿保护；PowerPoint 图表/图片/切换/形状/文本框/连接线/主题预设/版面质检/动画/SmartArt 风格图/演讲者备注批量/节（section）。
- 生成文件写入中性的核心属性：`author`/`lastModifiedBy` 默认 `Office Forge`，`created`/`modified` 使用当前 UTC 时间，`revision` 从 1 开始；任意 create spec 都可传 `properties: {author, title, subject, keywords, comments, category}` 覆盖，docx 同时保留历史顶层 `author` 写法。
- 可选真实后端：`enableBackends` 默认开启；`officeforge_pdf` 始终可用（包内 pypdf），`officeforge_convert`/`officeforge_preview`/`officeforge_recalculate` 依赖宿主 `officeKit` 或本机 COM/LibreOffice。`officeforge_native` 需要 Windows + Microsoft Office + pywin32；缺后端时返回 `BACKEND_UNAVAILABLE`，不会影响 Native 工具。

## 快速示例

创建一个 Word 文档：

```json
{
  "path": "reports/weekly.docx",
  "type": "docx",
  "spec": {
    "title": "周报",
    "subtitle": "2026-11-21",
    "sections": [
      { "heading": "本周进展", "bullets": ["完成 A 模块", "验证 B 流程"] },
      { "heading": "下周计划", "numbered": ["修复 C", "上线 D"] }
    ]
  }
}
```

创建一个带图表、条件格式和图片的 Excel 工作簿：

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
        "charts": [{ "type": "bar", "dataRange": "B1:B4", "categoriesRange": "A2:A4", "anchor": "E2", "title": "收入趋势" }],
        "conditionalFormatting": [{ "range": "B2:B4", "type": "cellIs", "operator": "greaterThan", "formula": "150", "fill": "FFC7CE" }],
        "images": [{ "path": "logo.png", "anchor": "G2", "width": 120, "height": 40 }]
      }
    ]
  }
}
```

填充一个带表格循环的 Word 模板：

```json
{
  "template": "templates/people.docx",
  "output": "out/people.docx",
  "data": {
    "title": "人员名单",
    "people": [
      { "name": "张三", "dept": "研发" },
      { "name": "李四", "dept": "产品" }
    ]
  },
  "options": { "expandLoops": true, "missing": "keep" }
}
```

导出 PDF：

```json
{
  "path": "reports/weekly.docx",
  "output": "out/weekly.pdf",
  "format": "pdf"
}
```

审计 Excel 公式健康：

```json
{
  "path": "reports/metrics.xlsx",
  "type": "xlsx",
  "mode": "health"
}
```

检查 PowerPoint 版面并列出主题：

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

## 配置

插件配置通过 DSH profile 的 bundle config 覆盖，均有默认值：

| 字段 | 默认值 | 说明 |
|---|---|---|
| `pythonPath` | 自动发现 | 可选，指定已安装 `python-docx`、`openpyxl`、`python-pptx` 的 Python 可执行文件。 |
| `engineTimeoutMs` | `120000` | 单次 Python 引擎调用的超时（毫秒）。 |
| `maxBatchItems` | `20` | 单次 `officeforge_batch` 的最大任务数。 |
| `enableExport` | `true` | 是否注册 `officeforge_export`。 |
| `enableBatch` | `true` | 是否注册 `officeforge_batch`。 |
| `enableBackends` | `true` | 是否注册 `officeforge_convert` / `officeforge_preview` / `officeforge_recalculate`；缺少后端时这些工具会显式报错。 |
| `modules` | 全部 `true` | `word` / `excel` / `ppt` / `pdf` 四个文档族开关。 |
| `features` | 全部 `true` | Word / Excel / PowerPoint 的细粒度高级能力开关。 |

安装到 profile 时，可在对应的 `cordis.patch.yml` entry 里覆盖配置，例如只启用 Word 和 Excel、关闭 Excel 图表：

```yaml
- insert:
    - id: dsh-office-forge
      name: dsh-office-forge
      config:
        modules: { word: true, excel: true, ppt: false, pdf: true }
        features:
          excel: { charts: false, dataTools: true }
```

Python 自动发现顺序：`pythonPath` → `OFFICEFORGE_PYTHON` → `DSH_PYTHON` → `$DSH_HOME/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe` → 扫描 `$DSH_HOME/dsh-runtimes/*/dependencies/python` → `PATH`。当前机器的捆绑 Python 3.12.14 已带全部所需包。

## 边界与安全设计

- **只写工作目录**：所有输入和输出路径都经过 `ctx.fs.resolve`、`ctx.fs.contains` 和词法相对路径双重检查，越界返回 `PATH_OUTSIDE_WORKSPACE`。
- **原子写入**：所有新文件先写同目录临时文件，再用 `os.replace` 原子替换；已存在文件默认拒绝，必须显式 `overwrite`。
- **有界执行**：单文件输入上限 50 MiB，输出上限 100 MiB，文本最多 200,000 字符，Excel 最多 5,000 行 × 300 列，PPT 最多 300 页，批次最多 20 项。
- **无递归**：每个工具只调用 Python 引擎的单一 action；工具之间不互相调用，`officeforge_batch` 禁止 `action: "batch"`，不存在自调用链。
- **不碰破甲**：插件不注册 prompt 变量、不接管 LLM stream、不注册 tool middleware / hook，不会改变 `@adwmc/helm-d` 的 hcot、attack-scheduler、advisory 等行为。
- **不覆盖内置工具**：所有工具名以 `officeforge_` 开头，不与内置 `office_render/office_convert/office_recalculate` 重名。
- **单工具隔离**：`apply` 内逐个注册，某个工具注册失败只记录 warning，不会让整个插件或 profile 激活失败。
- **兼容性加固（v0.6.1）**：`officeforge_export` 使用根会话作为 `officeKit` 输出 owner，转换前后校验源文件版本，并在远程 workspace 上拒绝不可映射的 `officeKit` 路径；非法 profile 配置会回退默认值而不是阻断插件激活；Excel 图表数据使用唯一的隐藏工作表，不复用或覆盖用户已有工作表。
- **插件页状态卡与环境检查（v0.14.0）**：`lib/client.js` 只注册 `plugins.detail.section` 一个只读卡片，显示版本、工具清单与“检查环境”按钮；按钮调用只读 `GET /officeforge/doctor`，渲染 health/summary 与逐项 checks/remedy，不读写配置、不发起文档操作、不注册其他 slot。

## 兼容性与隔离（v0.14.0）

- 只注入 `tools`、`fs`、`workingDirectory` 三个宿主服务；不注册 `systemPrompt`、不接管 LLM stream、不注册 middleware/hook，`@adwmc/helm-d` 的 hcot、attack-scheduler、advisory 不受影响。
- 客户端 bundle 只注入 `slots`，只注册 `plugins.detail.section`（id `officeforge.status`，order 40）；不注册 `plugins.bundle.config`、不修改任何配置项，插件页开关与工具注册保持一致。
- 十三个工具名全部以 `officeforge_` 开头；与内置 `office_render` / `office_convert` / `office_recalculate`、`run_code` 以及常见 `word_*` / `excel_*` / `ppt_*` / `univer_*` 工具不重名。
- `tests/test_compatibility.mjs` 用真实 `assertSupportedJsonSchema` 校验十三个输出 schema，并验证模块全关时注册 0 工具、仅启用 Word 时类型枚举收窄、不会覆盖宿主已有工具。
- 所有文档操作在短生命周期 Python 子进程中完成；不修改进程环境、不写全局配置、不启动 DSH、不写入任何 profile。
- 单工具注册失败只记录 warning；配置解析失败回退默认值；批处理逐项隔离错误。多个插件同时安装时仍应避免并发写同一文件，Office Forge 不提供跨插件文件锁。

## 目录

```
office-forge/
  package.json           # DSH bundle 与 npm 包元数据
  cordis.patch.yml       # 向 profile 插入 dsh-office-forge
  icon.svg
  locale/
    zh.json              # { "meta": { "title", "description" } }
    en.json
  resources/
    themes.json          # 31 套可外部化主题（24 套来自 dsh-ppt + 7 套旧主题）
    features.json        # 模块/功能注册表
  lib/
    index.js             # DSH 工具注册、路径策略、Python 调度、officeKit 导出/预览/重算
    client.js            # 插件页只读状态卡 + 一键环境检查（plugins.detail.section）
    office_engine.py     # JSON stdin/stdout 引擎入口
    office_common.py     # 路径、原子写、占位符、数据文件、通用校验
    office_pdf.py        # PDF 文本提取（内置 vendor pypdf/fitz/pdfplumber/fallback）
    office_pdf_tools.py  # PDF 合并/拆分/旋转/水印
    office_backends.py   # 可选 COM / LibreOffice 旧格式转换、PDF 导出
    office_com_worker.py # 独立 pywin32 COM Worker（Word/Excel/PowerPoint）
    office_migrations.py # spec schema 迁移
    vendor/pypdf/        # 包内 PDF 读取/写入依赖
    office_docx.py       # Word 读/写/编辑/模板
    office_xlsx.py       # Excel 读/写/编辑/模板
    office_pptx.py       # PowerPoint 读/写/编辑/模板
  benchmarks/
    run_benchmarks.py   # performance/robustness baseline (12 real engine cases)
    baseline.json       # per-case maxSeconds / maxPeakMB thresholds
  tests/
    test_engine.py       # Python 引擎端到端测试（临时工作区）
    test_fuzz.py         # 畸形输入/路径逃逸鲁棒性测试
    test_com_cleanup.py  # COM 进程回收验证（需要 Office + pywin32）
    run_regression.py    # 一次跑完全部套件并对照回归基线
    regression_baseline.json # 各套件断言下限
    test_plugin.mjs      # Node 插件层 mock host 测试
    test_plugin_config.mjs # 模块/功能开关测试
    test_compatibility.mjs # 宿主契约、工具名冲突与隔离检查
    test_client.mjs      # 客户端 bundle 形态与 slot 注册
    test_client_host.mjs # 用真实 dsh-client-modules 驱动 client bundle
    test_dsh_profile.mjs # 隔离 DSH_HOME 中的一次性安装验证
  docs/
    BENCHMARKS.md        # performance and robustness baseline
    DESIGN.md            # 调研来源、架构、非递归约束与兼容性说明
    COMPATIBILITY.md     # v0.14.0 宿主契约、工具命名空间与隔离审查
    COMPATIBILITY-MATRIX.md # 插件能力/格式/后端兼容矩阵
    CHANGELOG.md         # 版本变更与迁移记录
    RESEARCH-v0.4.md     # v0.4 候选插件调研：吸收公式体检、PPT 主题与质检；不引入 COM/exceljs/LibreOffice
  README.md
  README.zh.md
  LICENSE
```

## 本地验证

在 `office-forge/` 目录下：

```bash
# 一次跑完全部套件（断言下限 + 完成标记双重校验）
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/run_regression.py

# 跳过需要 Office 与 DSH CLI 的两套
"C:/Users/zhoutao/DSH 官方 0.2.2/app/profile/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" tests/run_regression.py --quick

# 单独运行
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

十套测试（含 benchmark 套件）都只使用临时目录和 mock host，不会启动 DSH，也不会写入任何 profile。`test_dsh_profile.mjs` 会在系统临时目录建一个一次性 `DSH_HOME`，安装本地 tarball 到临时 profile 后立即删除，用于验证真实安装路径。`test_com_cleanup.py` 会真实调用本机 Office 导出 PDF，并比对导出前后的 `WINWORD/EXCEL/POWERPNT` 进程数，确保 COM 后端不留孤儿进程；没有 Office 或 pywin32 时该用例自动 SKIP。

## 安装

把打包后的 tarball（或本目录）作为本地 bundle 加入目标 profile 并启用 `dsh-office-forge`。本机构建环境已在 desktop profile 中安装该插件，可用 `officeforge_doctor` 查看实际生效配置。

## License

MIT
