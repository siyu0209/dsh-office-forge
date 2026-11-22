# Office Forge v0.4 调研记录

调研时间：2026-11-21。目标是在不安装任何插件、不修改 DSH profile 的前提下，从插件库与网络资料中寻找可被 Office Forge 深度吸收的能力，而不是整体拼接运行时。

## 候选与判断

| 候选 | 观察到的能力 | 判断 |
|---|---|---|
| `dsh-excel-chat` 0.43.0（npm，346 KB） | 约 25 个 Excel 工具：`excel_health_report`、`excel_validate_formulas`、`excel_find_errors`、`excel_repair_formulas`、`excel_profile`、`excel_insight`、`excel_diff_workbook`、图表/透视等。核心公式审计为纯 XML/exceljs；图表/PDF 依赖 Windows Excel COM。 | 采纳公式健康、profile、insight、diff 的文件级模型；不引入 exceljs、COM 或它的 `systemPrompt` 注册。 |
| `dsh-ppt` 0.33.0（npm，约 1.6 MB） | 12 个工具：主题、模板、创建、编辑、检查、渲染、导入、吸收、引用；24 套主题、6 套母版骨架、35+ 版式。依赖 `@office-kit/pptx`、`fflate`、`csv-parse`、`jsep`，可选 LibreOffice 渲染。 | 采纳 8 套内置主题、`deck` 质量检查和 `applyTheme` 的文件级子集；不引入其大依赖链、模板库或 LibreOffice。 |
| `dsh-ppt-studio` 1.2.7 | 22 个工具、17 个技能，强调逐页制作、布局检查、Office 真渲染复核、预设隔离。 | 作为版式/质检思路参考；运行时依赖过重，不嵌入。 |
| `dsh-office-toolkit` 0.3.33 | 8 个工具，纯 JS，支持 xls/ods/pdf 读取、多表 join/透视、表格画像。 | 作为 Excel 工具面参考；本次只吸收审计/画像子集。 |
| `dsh-office-kit` 0.4.8、`@huiliyi37/dsh-office` 0.2.4、`dsh-plugin-wps-office-next` 0.9.0 | 通用 Office 读写或 WPS COM 路线。 | 备选，不引入外部二进制/COM。 |
| `@copylee/dsh-office` 1.2.2 | Windows COM 直接操作已打开的 Word/Excel/PPT；题注、脚注、引用、样式、页面设置、目录、更新字段、透视、主题、SmartArt、成品版式。 | v0.3 已吸收文件级可验证的 Word 题注/样式/字段、Excel 公式列/数据工具、PPT 文本框/连接线；v0.4 继续参考其主题与质检清单，不引入 COM。 |

## 设计结论

1. **不整体嵌入**：`dsh-excel-chat` 与 `dsh-ppt` 都带有自己的工具注册、提示词或大依赖链；直接整体安装会与现有 `officeforge_*` 工具重复，并可能触碰 prompt/LLM 注册边界。
2. **只吸收可文件级验证的能力**：Excel 公式体检/profile/insight/diff 可用 `openpyxl` 在内存中完成；PPT 主题与版面检查可用 `python-pptx` 完成。两者都不会启动 Office、LibreOffice 或网络服务。
3. **保持一个职责边界**：新增的 `officeforge_audit` 是只读工具，不修改原文件；写入仍由 `officeforge_create/edit/template` 负责。
4. **不破坏破甲**：没有新增 prompt 变量、LLM stream hook、tool middleware 或全局副作用；所有注册仍在 `ctx.effect()` 内完成。

## 本版落地

- `lib/office_xlsx.py`：`_health_xlsx`、`profile_xlsx`、`insight_xlsx`、`diff_xlsx`、`audit_xlsx`。
- `lib/office_pptx.py`：`_THEMES`、`_apply_theme`、`_deck_check`、`audit_pptx`。
- `lib/office_engine.py`：`action: "audit"` 调度与 capability 探测。
- `lib/index.js`：`officeforge_audit` 工具、`excel.audit` / `ppt.audit` / `ppt.themes` 开关。
- `tests/test_engine.py`、`tests/test_plugin.mjs`、`tests/test_plugin_config.mjs`：覆盖审计、主题、开关关闭与工具注册。

## 边界

`officeforge_audit` 不计算 Excel 公式结果，不做 Excel COM 级图表重建，也不做 PPT 动画/SmartArt/母版级渲染；它只报告文件结构、公式文本模式与 PPT 文本/主题状态。`dsh-excel-chat` 的 `systemPrompt` 和 `dsh-ppt` 的模板/渲染引擎均未复制。
