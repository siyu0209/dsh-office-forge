---
name: officeforge-ppt
description: "Office Forge 的 PPT 制作与修改入口。用户要做 PPT、幻灯片、演示文稿、deck、汇报、路演、答辩、课件，或要修改已有 .pptx 时使用；先分层递进地获取需求：第一轮只问受众/场景、新做还是已有、输出格式，第二轮再按第一轮答案规划内容/数据/页数/品牌等具体问题，不套固定问卷，再由助手选择 Guizang HTML、studio、zbsph、legacy 或编辑路径，最后用 officeforge_deck / officeforge_create / officeforge_edit 完成并验证。"
whenToUse: "Use BEFORE any PPT creation or modification; start with 1–2 broad direction questions, then plan the next round from those answers. Do not dump a fixed questionnaire."
compatibility: "Office Forge >= 3.4.0 on DSH 0.2.2+; uses the plugin's Office Forge tools and optional dsh-ppt-studio adapter."
allowed-tools: "Bash, Read, Write, Edit, officeforge_deck, officeforge_create, officeforge_edit, officeforge_read, officeforge_preview, officeforge_audit"
---

# Office Forge PPT 制作规则

当本插件可用，用户要做或改 PPT 时，优先按本规则执行。目标不是先问一堆问题，而是**先通过 1–3 个关键问题确定意图，再由我选择能力**，不要求用户说“用 Guizang/用 studio”。

## 0. 强制入口：分层递进地问，不要一次问完

用户只说“做个 PPT / 改个 PPT / 做个汇报”时，**不要直接开工，也不要先报引擎名称**。按层次递进：

### 第一轮：只定大方向（1–2 个问题）

先问最粗的三件事，用户答到什么程度就停到什么程度：

1. **给谁看、什么场景**（汇报/路演/答辩/分享/发布会），希望观众记住什么。
2. **新做还是改已有**；改已有必须问文件路径或附件。
3. **输出 `.pptx` 还是 `.html`**；不确定时给建议，不替用户硬选。

这一轮**不要**追问章节、数据、页数、配色、字体、模板等细节。

### 第二轮：根据第一轮答案，规划具体问题（2–4 个）

复述一句“我理解你要的是 X，输出 Y，重点 Z”，然后只问本轮真正缺失、且不问就会做偏的事实：

- 内容来源：已有文档/数据/大纲，还是只给主题；
- 必须包含的章节、关键数字、页数范围；
- 公司模板、品牌色、字体、Logo、必须沿用的母版；
- 修改约束（哪些页不能动、是否保留原主题）。

### 后续轮次：按缺口继续，轮次不限

每轮回答后，根据这一轮新得到的信息规划**下一轮**要问什么；还有关键事实缺失就继续追问 1–2 个，信息足够就停，不设固定轮数上限。不要复制同一套问卷。配色、字体、图标、动画、页码等纯视觉细节不盘问，先出 2–3 页样张让用户看方向。

用户说“你定/随便”时，按场景默认并说明假设；用户已给足信息时直接做，不要为了走流程再问。

## 1. 先判断信息缺口

先在心里过一遍下面五类关键事实，只问真正缺失、且不问就会做偏的 1–2 个问题：

1. **新做还是修改已有**：新做，还是改某个 `.pptx` 文件？修改时先要文件路径或附件。
2. **内容来源**：用户已给文档/数据/大纲，还是只给一个主题？只给主题时，不要凭记忆编造具体数字；需要时先说明会按公开资料整理，或请用户给材料。
3. **必须包含什么**：必须有的章节、关键数据、页数范围、受众、使用场景。
4. **输出格式**：`.pptx`、`.html`（网页放映），还是两者都要。
5. **品牌/模板**：有没有公司模板、品牌色、字体、Logo 或必须沿用的母版。

### 不要先问的内容

配色、字体、图标、动画、页码、留白、卡片样式这类**只影响视觉**的细节，不要先盘问。先按场景生成 2–3 页样张，让用户看到方向后再定。

## 2. 问法

- **不固定问卷**。根据当前缺口决定问什么，不要每次都把同一组问题复制一遍。
- 一次只问 **1–2 个真正关键的问题**；其余用合理默认值推进。
- 用户说“你定”“随便”“按你的来”时：直接按场景默认制作，并在交付说明里写清假设（例如“默认 16:9、中文、商务汇报、consulting 主题”）。
- 用户已给足信息时：**直接做**，不要为了走流程再问。

## 2.5 风格菜单（问风格，不问模块）

不要问“你要用 studio / zbsph / Guizang / legacy 哪个”。用下面的场景化选项问：

| 风格 | 适合什么 | 输出 | 内部连接 |
|---|---|---|---|
| 商务汇报 / 董事会 | 季度复盘、经营分析、汇报 | `.pptx` | studio consulting/swiss |
| 路演 / 产品发布 | 融资、产品 launch、发布会 | `.pptx` 或 `.html` | studio pitch/keynote 或 Guizang |
| 学术答辩 / 研究报告 | 论文、课题、评审 | `.pptx` | academic 模板 |
| 杂志风 / 品牌故事 | 品牌、作品集、发布会 | `.html` | Guizang template-swiss |
| 瑞士极简 / 高级灰 | 设计评审、创意提案 | `.html` 或 `.pptx` | Guizang Swiss / studio swiss |
| 科技暗色 / 数据大屏 | 技术发布、AI/数据产品 | `.pptx` 或 `.html` | studio tech-dark / Guizang |
| 培训课件 / 教学 | 课程、内训、知识分享 | `.pptx` | studio course |
| 政务 / 正式公文 | 政府汇报、合规材料 | `.pptx` | studio gov-report |
| 数据复盘 / 表格图表 | 多 sheet、指标、图表 | `.pptx` | zbsph |
| 网页长页 / 交互展示 | 滚动、动画、演讲者视图 | `.html` | Guizang / webdeck |

规则：用户不知道风格时，按内容推荐 2–3 个并标一个推荐；用户选风格后自动映射，不要求用户再选引擎；选了“杂志风/网页”默认 `.html`，坚持 `.pptx` 时说明只是近似版式。

## 2.6 联网查找资料

- 允许为了制作主动搜索公开网络，但只在**内容缺口确实需要外部信息**时使用。
- 先说明要查什么，再搜索；优先官方、权威、可核验来源；在交付说明或备注里列出**来源名称/URL + 访问日期**。
- 不把网络资料冒充成用户提供的数据；不编造数字；找不到可靠来源时明确说“未找到公开数据”，不要硬填。
- **不能上传用户的本地文件、内部数据、客户资料**到外部服务；如果内容涉及公司内部数据，先向用户要文件，不要用搜索代替。
- 用户说“你定/随便”时，可以按场景主动查公开资料，但要在交付里说明查了什么、用了哪些来源。
- 搜索不可用或没有可信来源时，直接说明限制，改用用户材料或明确标注的合理假设。
- 动手搜索前先自动检查当前会话可用工具/技能目录：原生 `web_search`/`research`、`web-search`/`bx-search`/`firecrawl` 技能、搜索类 MCP/插件，或已安装的 `dsh-free-search` 插件（`web_search`/`advanced_search`/`multi_search`/`platform_search`/`research`/`free_search_test`）存在就直接调用，不要让用户指定工具；都没有时说明限制。

## 3. 能力选择（由我决定，不要求用户说引擎名）

先读 `officeforge_deck` 的 `capabilities`，再按“用户要什么 + 输出形态”选：

| 用户意图 | 先选能力 | 说明 |
|---|---|---|
| 网页 PPT、杂志风、瑞士风、发布会感、需要横向翻页/演讲者视图 | `officeforge_web` + `engine:"guizang"`；复杂叙事再加载 `guizang-ppt-skill` | 先用 `officeforge_web` 的 Guizang 引擎直接生成单文件 HTML（使用插件内真实 `template-swiss.html` 和原套视觉系统）；需要自定义版式/演讲者细节时，再按 `guizang-ppt-skill` 的 `SKILL.md`、`assets/template*.html`、`references/layouts*.md` 补。**这才是 Guizang 的原套能力，不是 PPTX 换色** |
| 要可编辑 `.pptx`，普通商务/汇报，使用 dsh-ppt 设计系统 | `officeforge_deck` + `studio` | 真实调用 dsh-ppt-studio；24 主题/8 master/20 模板/82 版式 |
| `.pptx` 里表格/图表/图片多，数据驱动 | `officeforge_deck` + `zbsph` | zbsph 全 schema；但 media/comments/master/morph/animation/SmartArt 仍按 `lossAudit` 明确告知，不静默回退 |
| 已有 `.pptx` 要改内容/局部版式 | `officeforge_deck` 的 `edit` | `import → Deck IR → operations → render`；保留原主题/master，未改页不动 |
| 已有 `.pptx` 要美化/重排 | `officeforge_beautify` | 默认 Anon1009；先 audit 再 beautify，评分下降要解释 |
| 要单文件 HTML/网页展示，但不是 Guizang 风格 | `officeforge_deck` + `composer`/`officeforge_web` | composer 是内置轻量 web deck；若要完整 composer 上游能力，先检查 adapter 是否可用 |
| 需要 media/comments/master/morph/animation/SmartArt/transition | `legacy` | 只有 legacy 原生对象能完整承载；必须显式告知能力归属，不假装 studio/zbsph 支持 |

硬规则：

- **Guizang 不再走 `officeforge_deck engine: "guizang"` 的 native PPTX 路径**；该路径只是色板近似，不能称为 Guizang 满血。用户要 Guizang → 用 `officeforge_web(engine:"guizang")` 生成真实模板 HTML；需要高级自定义时再加载 `guizang-ppt-skill`。
- 用户没有明确要 `.pptx` 时，若需求是分享/发布会/设计感，优先建议 HTML；用户选 `.pptx` 时按上表走 studio/zbsph/legacy，并明确差异。
- 修改已有 `.pptx` 时，不要先改主题色冒充“重做”；先 `import`/`audit`，按 `lossAudit` 告知哪些对象会丢失，再选择 legacy 或局部重排。
- 生成后必须验证：HTML 用浏览器/预览检查；PPTX 用 `officeforge_preview` 或 `officeforge_deck verify` 真实渲染，不能只看文件存在。

## 4. 交付

- 给出绝对路径、格式、页数、使用的主题/模板、adapter；如有假设，一并说明。
- 如果某项能力不可用，明确说“当前环境缺少 X，已使用 Y 回退”，不要把回退包装成完整能力。
- 不覆盖用户原文件，除非用户明确要求；修改默认另存新文件，并保留原件。
