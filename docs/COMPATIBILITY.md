# Office Forge v0.14.0 兼容性与隔离审查

本文件记录 Office Forge 与 DSH 宿主、现有 profile 插件及 Office 同类插件共存时的边界。目标是：不污染 profile、不覆盖内置工具、不注册 prompt/LLM/middleware hook、不改变 `@adwmc/helm-d` 的破甲行为。

## 1. 宿主契约核对

| 项目 | 结论 |
|---|---|
| `ToolsService.register(definition)` | 返回 exact disposer；插件在 `ctx.effect` 中保存并在卸载时调用。 |
| `defineTool` 参数/输出 schema | 十三个工具的 `parameters` 与 `output.schema` 均通过真实 `assertSupportedJsonSchema`；输出 schema 根级不使用 `required`。 |
| `dsh.bundle.patch` | `package.json` 使用字符串 `./cordis.patch.yml`，`loadOverlayPatches` 解析结果为单个 `insert` entry。 |
| `officeKit` 导出 | 使用 `subagents.rootOf(agent).session.id` 作为输出 owner；源文件转换前后校验 `stat.version`；远程 workspace 不可映射时返回 `OFFICEKIT_REMOTE_WORKSPACE`。 |
| 配置异常 | `normalizeConfig()` 捕获非法 profile 配置并回退 `Config({})`，不会阻断激活。 |
| `dsh.client` 客户端 bundle | `package.json` 声明 `dsh.client.platform = "web"` 与 `exports["./client"] = "./lib/client.js"`；bundle 用宿主 lazy-CJS 形态 `window.__ModuleLoader__.load({id, factory})`，factory 返回 `module.exports`；`inject = ["slots"]`，只注册 `plugins.detail.section` 一个只读卡片。 |

## 2. 工具命名空间

Office Forge 只注册以下十三个工具：

`officeforge_read`、`officeforge_create`、`officeforge_edit`、`officeforge_template`、`officeforge_audit`、`officeforge_doctor`、`officeforge_batch`、`officeforge_export`、`officeforge_convert`、`officeforge_pdf`、`officeforge_native`、`officeforge_preview`、`officeforge_recalculate`。

| 插件/来源 | 相关工具命名 | 与 Office Forge 的关系 |
|---|---|---|
| DSH 内置 Office | `office_render` / `office_convert` / `office_recalculate` | 不重名，不覆盖；Office Forge 不调用它们。 |
| DSH 核心 | `run_code` | 不重名；保留给宿主。 |
| `@adwmc/helm-d` | 破甲 hooks（hcot/attack-scheduler/advisory） | Office Forge 不注册 prompt/LLM stream/middleware，不触碰。 |
| `dsh-rea` | `mcp__rea__*` | 不同命名空间。 |
| `dsh-ops` | `ops_*` | 不同命名空间。 |
| `billion-context` | `acp_*` | 不同命名空间。 |
| `dsh-free-search` / `@liustack/modsearch` | 搜索类工具 | 不同命名空间。 |
| `dsh-context` | 上下文管理工具 | 不同命名空间。 |
| `dsh-mnemon` / `@openviking/*` | `mnemon_*` / 记忆类工具 | 不同命名空间。 |
| `@changfenhuang/dsh-genui` | `render_ui` | 不同命名空间。 |
| `dsh-plugin-image-workbench` | 图片工作台工具 | 不同命名空间。 |
| `dshmarket` | `plugin_manager` | 不同命名空间。 |
| `ds-harness-remote` | `remote_*` | 不同命名空间。 |
| `dsh-excel-chat`（若安装） | `excel_*` | 不重名；能力有重叠，默认不同时调用。 |
| `dsh-ppt`（若安装） | `ppt_*` | 不重名；能力有重叠，默认不同时调用。 |
| `dsh-univer-office`（若安装） | `univer_*` | 不重名。 |
| `dsh-office-tools`（若安装） | `word_*` / `excel_*` / `ppt_*` | 不重名。 |

## 3. 并发与文件安全

- 每个工具只启动一个短生命周期 Python 子进程；不修改 `process.env`、不写全局配置、不启动 DSH、不写 profile。
- 输出文件使用同目录临时文件 + `os.replace`；默认拒绝覆盖，显式 `overwrite` 才替换，并可生成同目录备份。
- `officeforge_batch` 的每个 item 独立 try/catch；坏路径只进入 skipped，不阻断同批其他 item；禁止嵌套 batch。
- Excel 图表数据写入唯一的隐藏工作表（`_officeforge_chart_data`、`_officeforge_chart_data_1`...），不复用用户已有表。
- 多个插件同时操作同一文件时没有跨插件文件锁；这是宿主级限制，调用方应避免并发写同一路径。Office Forge 不声称提供分布式锁。

## 4. 可执行验证

```bash
node tests/test_compatibility.mjs
```

测试覆盖：真实 schema 校验、十三工具注册、内置工具名不覆盖、模块全关时仅保留 `officeforge_doctor`、仅 Word 时类型枚举收窄、无 `systemPrompt`/LLM stream/middleware/hcot/attack-scheduler/advisory 代码。输出 `COMPATIBILITY TESTS PASSED`。

同时执行：

- `tests/test_engine.py` → `ALL TESTS PASSED`
- `tests/test_plugin.mjs` → `PLUGIN TESTS PASSED`（十三工具、doctor、convert/preview/recalculate）
- `tests/test_plugin_config.mjs` → `PLUGIN CONFIG TESTS PASSED`
- `tests/test_client.mjs` → `CLIENT UI TESTS PASSED`（manifest、lazy-CJS 形态、slot 注册）
- `tests/test_client_host.mjs` → `CLIENT HOST MACHINERY TESTS PASSED`（真实 `dsh-client-modules` 驱动）
- `tests/test_dsh_profile.mjs` → `DSH PROFILE TESTS PASSED`（临时 DSH_HOME 一次性安装）
- `python -m py_compile lib/*.py`
- `node --check lib/index.js lib/client.js`
- `npm pack` 后解包 `node --check` / `py_compile` → `TARBALL_SYNTAX_OK`
