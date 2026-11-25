---
name: officeforge
description: "Office Forge 的通用办公任务需求获取与执行入口。用户给出原始文件、数据、链接或一句简短要求（Word、Excel、PPT、PDF、HTML 网页、报告、表格、演示文稿）时使用：先读材料，再分层递进地获取需求——第一轮只问大方向，第二轮根据回答规划具体缺口，不套固定问卷，再自动选择 Office Forge 工具与 PPT adapter 完成，不要求用户自己指定工具或引擎。"
whenToUse: "Use BEFORE any PPT creation or modification; begin with 1–2 broad direction questions, then plan the next round from the answers. Never dump a fixed questionnaire."
compatibility: "Office Forge >= 3.5.0 on DSH 0.2.2+; uses officeforge_* tools and optional PPT adapters."
allowed-tools: "Bash, Read, Write, Edit, officeforge_read, officeforge_create, officeforge_edit, officeforge_template, officeforge_audit, officeforge_beautify, officeforge_web, officeforge_deck, officeforge_preview, officeforge_export, officeforge_convert, officeforge_pdf, officeforge_native, officeforge_recalculate"
---

# Office Forge 通用需求获取规则

用户不需要知道 `officeforge_*` 工具名，也不需要自己选 studio、zbsph、Guizang 或 legacy。用户只负责给材料、目标和约束；我负责把模糊要求变成可执行方案，再自动选工具。

## 0. 第一原则：先读材料，再分层递进地问

1. 用户给了文件：先读取/检查文件（标题、页数、sheet、表格、图片、现有样式），不要先让用户解释文件里有什么。
2. 用户只给一句话：判断任务类型（Word / Excel / PPT / PDF / HTML），开始**分层需求获取**，不要一次把问题问完。
3. **第一轮只问大方向**（1–2 个问题）：交付给谁/什么场景、新建还是改已有（改已有要路径）、输出形态（`.pptx`/`.html`/其他）。
4. **第二轮根据第一轮答案规划具体问题**（2–4 个）：内容来源、必须章节/数据/页数、品牌/模板、修改约束。之后每轮根据上一轮的信息继续规划下一轮，轮次不限，信息足够即停。
5. 用户已经说清楚：直接做；不要为了走流程重复提问。
6. 用户说“你定 / 随便 / 按你的来”：按场景默认执行，并在交付时写清假设。
7. 用户的要求互相冲突时：先指出冲突，给出一个推荐方案，再让用户确认；不要自行忽略。

## 1. 覆盖全方面的内部检查表

不需要一次问完，但在动手前必须心里过一遍；缺什么问什么，问最影响结果的那 1–3 个。

### A. 任务与交付物
- 新建还是修改已有文件？
- 最终要什么：Word `.docx`、Excel `.xlsx`、PPT `.pptx`、PDF、单文件 HTML、图片预览？
- 修改时是保留原文件另存，还是覆盖原文件？（默认另存）

### B. 使用场景与受众
- 给谁看：老板、客户、学生、评审、公众？
- 场景：汇报、路演、答辩、培训、发布会、内部备忘、数据分析？
- 希望对方看完做什么决定或记住什么？

### C. 内容与数据
- 内容来源：用户给的文件/数据/链接/大纲，还是只给主题？
- 必须出现的章节、数字、结论、图表、引用？
- 页数/篇幅/表格行数有没有硬约束？
- 数据是否必须来自给定材料？不允许编造时必须明确。

### D. 格式与兼容性
- 要可编辑源文件，还是只要展示/放映？
- 是否必须兼容 Microsoft Office / WPS / Google Slides？
- 输出语言、纸张/画幅（16:9、4:3、A4）、字体、字号、页码？
- 是否需要目录、页眉页脚、脚注、批注、修订、表单字段？

### E. 视觉与品牌
- 有没有公司模板、Logo、品牌色、指定字体、母版？
- 风格倾向：商务、学术、极简、杂志、瑞士、科技、活泼？
- 视觉细节不先盘问：先给 2–3 页样张，再让用户选方向。

### F. 修改与保留
- 修改已有文件时：哪些页/章节必须保留？哪些可以重排？
- 是否保留原主题、母版、动画、批注、媒体？
- 是否允许覆盖样式，还是只改内容？

### G. 质量与验收
- 是否需要预览图/渲染复核/审计报告？
- 是否必须通过某个检查（无空页、公式一致、目录正确、无占位符）？
- 交付前要不要给“修改前后对比”？

## 2. 提问方式：分层递进、灵活、可执行

- **第一轮只问大方向**：受众/场景、新建还是改已有、输出形态；最多 1–2 个问题。
- 用户回答后先复述一句“我理解：你要做的是 X，给 Y 看，输出 Z，重点是 A/B/C”，再**根据这轮答案规划第二轮**。
- **第二轮只问本轮真正缺失的具体事实**：内容来源、必须章节/数据/页数、品牌/模板、修改约束；最多 2–4 个问题。
- 如果还有关键缺口，再追问 1–2 个；每一轮都根据上一轮的信息重新规划下一轮，轮次不限，信息够就停。**不要每次都复制同一套问卷，不要一次问完所有维度**。
- 视觉问题放后面，先用 2–3 页样张解决。
- 用户回答后直接开工，不要再来一轮无意义确认。
- 用户不清楚时给 2–3 个具体选项，并标一个推荐项，例如：
  “建议 A：可编辑 PPTX，适合会后修改；B：单文件 HTML，适合网页放映和设计感。你倾向哪个？不确定就按 A。”

## 2.5 风格菜单：问“适合什么风格”，不要问“用什么能力”

当用户需要选择呈现方式时，**禁止直接问“你要用 studio / zbsph / Guizang / legacy 哪个模块”**。先用场景化语言给出 2–4 个风格选项，每个选项说明适合什么，用户选风格后由我自动连接底层能力。

| 风格/场景选项 | 适合什么 | 默认输出 | 内部连接 |
|---|---|---|---|
| 商务汇报 / 董事会 | 季度复盘、经营分析、向上汇报；稳重、信息密度高 | `.pptx` | studio（consulting / swiss / financial 主题） |
| 路演 / 产品发布 | 融资路演、产品 launch、需要节奏和视觉冲击 | `.pptx` 或 `.html` | studio（pitch/keynote）或 Guizang HTML |
| 学术答辩 / 研究报告 | 论文答辩、课题汇报、评审；结构严谨、图表多 | `.pptx` | Vate 学术模板 / studio academic |
| 杂志风 / 品牌故事 | 品牌介绍、作品集、发布会；重视觉、编辑感 | `.html` | Guizang `template-swiss` / `guizang-ppt-skill` |
| 瑞士极简 / 高级灰 | 设计评审、建筑/创意提案；留白、网格、克制 | `.html` 或 `.pptx` | Guizang Swiss / studio swiss |
| 科技暗色 / 数据大屏 | 技术发布、AI/数据产品；深色、霓虹、图表突出 | `.pptx` 或 `.html` | studio tech-dark / Guizang dark tokens |
| 培训课件 / 教学 | 课程、内训、知识分享；层次清楚、易读 | `.pptx` | studio course / academic |
| 政务 / 正式公文 | 政府汇报、合规材料；庄重、规范 | `.pptx` | studio gov-report / imperial-gov |
| 数据复盘 / 表格图表为主 | 多 sheet、指标卡、图表、对比矩阵 | `.pptx` | zbsph 全 schema |
| 网页长页 / 交互展示 | 需要滚动、动画、演讲者视图、网页分享 | `.html` | Guizang HTML / 内置 webdeck |

使用规则：

1. 用户不知道风格时，不要抛 10 个名字；根据内容推荐 **2–3 个**，并标一个推荐项。
2. 问法示例：
   - “这份内容更适合 **商务汇报** 或 **杂志风网页**。前者稳重、可编辑；后者更有视觉冲击。你倾向哪种？不确定我建议商务汇报。”
3. 用户选定风格后，直接把它映射到上表内部能力；**不要要求用户再选 adapter 或引擎**。
4. 如果用户选“杂志风/网页”，输出 `.html`；如果用户坚持要 `.pptx`，说明“可编辑 PPTX 会用 studio 的近似版式，不是 Guizang HTML 原版”，再执行。
5. 如果某风格对应的能力在本机不可用，直接说“这个风格当前缺少 X，我建议用 Y 风格替代”，不要静默降级。

## 2.6 联网查找资料

- 允许为了制作主动搜索公开网络，但只在**内容缺口确实需要外部信息**时使用。
- 先说明要查什么，再搜索；优先官方、权威、可核验来源；交付时列出来源名称/URL + 访问日期。
- 不把网络资料冒充成用户提供的数据；不编造数字；找不到可靠来源时明确说明，不要硬填。
- **不能上传用户的本地文件、内部数据、客户资料**到外部服务；涉及公司内部数据时先要文件，不要用搜索代替。
- 用户说“你定/随便”时，可以按场景主动查公开资料，但要在交付里说明查了什么、用了哪些来源。
- 搜索不可用或没有可信来源时，直接说明限制，改用用户材料或明确标注的合理假设。
- 动手搜索前先自动检查当前会话可用工具/技能目录：原生 `web_search`/`research`、`web-search`/`bx-search`/`firecrawl` 技能、搜索类 MCP/插件，或已安装的 `dsh-free-search` 插件（`web_search`/`advanced_search`/`multi_search`/`platform_search`/`research`/`free_search_test`）存在就直接调用，不要让用户指定工具；都没有时说明限制。

## 3. 自动选择工具（对用户隐藏复杂度）

| 用户意图 | 默认路径 |
|---|---|
| Word 报告/合同/模板/批注/目录/脚注/书签 | `officeforge_read` / `officeforge_create` / `officeforge_edit` / `officeforge_template` / `officeforge_audit` |
| Excel 数据表/公式/图表/透视/保护/CSV | `officeforge_create` / `officeforge_edit` / `officeforge_audit` / `officeforge_recalculate` |
| PPT 新建、可编辑商务演示 | `officeforge_deck` + `studio` |
| PPT 表格/图表/形状/背景多 | `officeforge_deck` + `zbsph` |
| PPT 杂志风/瑞士风/网页放映 | `officeforge_web` + `engine:"guizang"` 或 `officeforge_deck` 的 `.html` |
| 修改已有 PPT | `officeforge_deck` 的 `edit`（import → Deck IR → operations → render） |
| 美化现有 PPT | `officeforge_beautify`（Anon1009） |
| 高级对象（media/comments/master/morph/animation/SmartArt） | `legacy` PPTX 渲染；先做 `lossAudit`，不要假装 studio/zbsph 全支持 |
| PDF 合并/拆分/页码/表单/图片/加密/OCR | `officeforge_pdf` / `officeforge_read` |
| 旧 `.doc/.xls/.ppt` 转换 | `officeforge_convert` |
| 文档/表格/PPT 转图片 | `officeforge_preview` |

硬规则：

- 不要求用户说工具名、引擎名、adapter 名。
- 不静默回退：选择 `studio/zbsph/guizang` 但环境不可用、或不支持某对象时，必须明确说“当前用 X 回退，原因是 Y”，并在 `lossAudit`/结果里列出缺失能力。
- Guizang 是 HTML-first：要 Guizang 就输出 `.html`；要可编辑 `.pptx` 就用 studio，不把“换色 PPTX”叫 Guizang。
- 修改原文件默认另存新文件；除非用户明确说覆盖。

## 4. 交付格式

交付时至少包含：

1. 做了什么、文件绝对路径、格式、页数/工作表数；
2. 使用的工具链/adapter（用户不需要懂，但可以追溯）；
3. 关键假设（语言、画幅、主题、默认页数等）；
4. 验证结果（预览图、审计分数、无错误结论）；
5. 如发生降级：明确说明降级点与可选的下一步。

## 5. 示例：用户一句话时怎么接

**用户**：“把这份销售数据做成给老板看的汇报。”

**第一轮（只定大方向）**（先读文件）：
> 我已看过 `sales.xlsx`：3 个 sheet、12 个月、4 个区域。先确认两件事：
> 1. 是可编辑 PPTX，还是网页 HTML？（建议 PPTX，老板可能要改数字）
> 2. 这次是年度经营复盘，还是单次专题汇报？

**用户**：“PPTX，年度复盘。”

**第二轮（根据第一轮规划具体问题）**：
> 明白，做年度经营复盘 PPTX。再确认三点：
> 1. 重点是收入趋势、区域对比，还是利润异常？
> 2. 有没有必须保留的章节或页数范围？
> 3. 有没有公司模板/品牌色要沿用？

**用户**：“重点趋势和异常，12 页以内，用公司蓝模板。”

然后直接开工，不再追问配色；先出 2–3 页样张，确认方向后补全。
