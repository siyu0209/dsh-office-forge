// Office Forge - a document workshop for DeepSeek Harness.
//
// Design contract:
//   * every tool is independent; no tool calls another tool;
//   * all document mutation lives in a short-lived Python child process;
//   * tools are registered one-by-one and a failed registration never blocks
//     the rest of the plugin or the host profile;
//   * no prompt variables, no LLM stream hooks, no tool middleware - the
//     existing "破甲" plugin keeps its own hooks untouched.

import { spawn } from "node:child_process";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { basename, dirname, extname, isAbsolute, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import z from "@deepseek-ai/schemastery";
import { defineTool } from "@deepseek-ai/dsh-tools";

const name = "dsh-office-forge";
const inject = ["tools", "fs", "workingDirectory"];

const ENGINE = fileURLToPath(new URL("./office_engine.py", import.meta.url));
const PYTHON_PROBE = "import docx, openpyxl, pptx; print('officeforge-ok')";
const MAX_ENGINE_STDOUT = 32 * 1024 * 1024;
const MAX_ENGINE_STDERR = 2 * 1024 * 1024;
const MAX_EDIT_OPERATIONS = 200;

const RESOURCE_ROOT = fileURLToPath(new URL("../resources/", import.meta.url));
const PLUGIN_ROOT = fileURLToPath(new URL("../", import.meta.url));
const VENDOR_PDF = join(PLUGIN_ROOT, "lib", "vendor", "pypdf");
const CLIENT_BUNDLE = join(PLUGIN_ROOT, "lib", "client.js");

const PLUGIN_VERSION = (() => {
	try {
		return JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")).version || "unknown";
	} catch {
		return "unknown";
	}
})();

function loadResource(name, fallback) {
	try {
		return JSON.parse(readFileSync(join(RESOURCE_ROOT, name), "utf8"));
	} catch {
		return fallback;
	}
}

const FEATURE_REGISTRY = loadResource("features.json", {
	modules: ["word", "excel", "ppt", "pdf"],
	features: {
		word: ["images", "hyperlinks", "pageBreaks", "toc", "captions", "styles", "fields", "comments", "revisions", "mailMerge", "footnotes", "bookmarks", "crossReferences", "citations", "columns", "watermark", "contentControls", "templates", "audit"],
		excel: ["charts", "conditionalFormatting", "dataValidation", "comments", "images", "mergedCells", "formulas", "dataTools", "audit", "pivots", "tables", "protection"],
		ppt: ["charts", "images", "transitions", "shapes", "themes", "audit", "animations", "smartart", "notes", "sections", "layouts", "templates", "master", "comments", "media"]
	}
});

const DEFAULT_MODULES = Object.fromEntries((FEATURE_REGISTRY.modules || ["word", "excel", "ppt", "pdf"]).map((key) => [key, true]));
const DEFAULT_FEATURES = Object.fromEntries(Object.entries(FEATURE_REGISTRY.features || {}).map(([module, names]) => [module, Object.fromEntries(names.map((key) => [key, true]))]));

function moduleSchema() {
	const shape = {};
	for (const key of FEATURE_REGISTRY.modules || ["word", "excel", "ppt", "pdf"]) {
		shape[key] = z.boolean().default(true).description(`Enable ${key} document tools.`);
	}
	return z.object(shape).default(DEFAULT_MODULES).description("Choose which document families are registered.");
}

function featureSchema(module) {
	const shape = {};
	for (const key of FEATURE_REGISTRY.features?.[module] || []) {
		shape[key] = z.boolean().default(true).description(`${module}.${key} feature switch.`);
	}
	return z.object(shape).default(DEFAULT_FEATURES[module] || {}).description(`${module} advanced feature switches.`);
}

const Config = z.object({
	pythonPath: z.string().description("Optional absolute path to a Python executable that already has python-docx, openpyxl and python-pptx installed."),
	comPythonPath: z.string().description("Optional absolute path to a system Python with pywin32 and Microsoft Office COM for officeforge_native/legacy conversion."),
	engineTimeoutMs: z.number().default(120000).description("Per-call Python engine timeout in milliseconds."),
	maxBatchItems: z.number().default(20).description("Maximum operations in one officeforge_batch call."),
	enableExport: z.boolean().default(true).description("Register officeforge_export (PDF conversion through the host officeKit service)."),
	enableBackends: z.boolean().default(true).description("Register officeforge_convert, officeforge_preview and officeforge_recalculate when their host services are available."),
	enableBatch: z.boolean().default(true).description("Register officeforge_batch."),
	modules: moduleSchema(),
	features: z.object({
		word: featureSchema("word"),
		excel: featureSchema("excel"),
		ppt: featureSchema("ppt")
	}).default(DEFAULT_FEATURES).description("Fine-grained advanced feature switches. Disabled features are refused by the engine.")
});

const EXTENSIONS = { docx: [".docx", ".dotx"], xlsx: [".xlsx"], pptx: [".pptx"] };
const READ_EXTENSIONS = {
	docx: [".docx", ".docm"],
	xlsx: [".xlsx", ".xlsm"],
	pptx: [".pptx", ".pptm"],
	pdf: [".pdf"]
};

function enabledTypes(config) {
	const types = [];
	if (config.modules?.word) types.push("docx");
	if (config.modules?.excel) types.push("xlsx");
	if (config.modules?.ppt) types.push("pptx");
	return types;
}

function readTypes(config) {
	const types = enabledTypes(config);
	if (config.modules?.pdf) types.push("pdf");
	return types;
}

function allowedExtensionsFor(config) {
	return enabledTypes(config).flatMap((type) => EXTENSIONS[type]);
}

function readExtensionsFor(config) {
	return readTypes(config).flatMap((type) => READ_EXTENSIONS[type] ?? []);
}

function featuresFor(config) {
	return config.features ?? DEFAULT_FEATURES;
}

function auditTypes(config) {
	const features = featuresFor(config);
	return enabledTypes(config).filter((type) => (type === "docx" && features.word?.audit) || (type === "xlsx" && features.excel?.audit) || (type === "pptx" && features.ppt?.audit));
}

// ---------------------------------------------------------------------------
// Python discovery
// ---------------------------------------------------------------------------

class EngineFailure extends Error {
	constructor(code, message, detail) {
		super(message);
		this.name = "EngineFailure";
		this.code = code;
		this.detail = detail;
	}
}

let cachedPython;
let pythonProbePromise;

function executableCandidates(config) {
	const candidates = [];
	const push = (value) => {
		if (typeof value === "string" && value.trim() !== "" && !candidates.includes(value)) candidates.push(value);
	};
	push(config?.pythonPath);
	push(process.env.OFFICEFORGE_PYTHON);
	push(process.env.DSH_PYTHON);
	const dshHome = process.env.DSH_HOME || process.env.DSH_PROFILE_HOME;
	if (dshHome) {
		const runtimes = join(dshHome, "dsh-runtimes");
		push(join(runtimes, "dsh-primary-runtime", "dependencies", "python", process.platform === "win32" ? "python.exe" : "python"));
		try {
			if (existsSync(runtimes) && statSync(runtimes).isDirectory()) {
				for (const entry of readdirSync(runtimes, { withFileTypes: true })) {
					if (!entry.isDirectory()) continue;
					push(join(runtimes, entry.name, "dependencies", "python", process.platform === "win32" ? "python.exe" : "python"));
					push(join(runtimes, entry.name, "dependencies", "python", "bin", "python"));
				}
			}
		} catch {
			// A missing DSH_HOME scan is not fatal; PATH candidates remain.
		}
	}
	push(process.platform === "win32" ? "python.exe" : "python3");
	push("python");
	return candidates;
}

function probePython(candidate) {
	return new Promise((resolveProbe) => {
		let settled = false;
		const finish = (value) => {
			if (settled) return;
			settled = true;
			resolveProbe(value);
		};
		let child;
		try {
			child = spawn(candidate, ["-c", PYTHON_PROBE], {
				windowsHide: true,
				stdio: ["ignore", "pipe", "pipe"],
				env: { ...process.env, PYTHONUTF8: "1", PYTHONIOENCODING: "utf-8" }
			});
		} catch {
			finish(null);
			return;
		}
		let output = "";
		const timer = setTimeout(() => {
			try { child.kill(); } catch { /* ignore */ }
			finish(null);
		}, 15000);
		child.stdout.on("data", (chunk) => { output += chunk.toString("utf8"); });
		child.stderr.on("data", () => { /* probing is silent */ });
		child.on("error", () => { clearTimeout(timer); finish(null); });
		child.on("close", (code) => {
			clearTimeout(timer);
			finish(code === 0 && output.includes("officeforge-ok") ? candidate : null);
		});
	});
}

async function resolvePython(config) {
	if (cachedPython) return cachedPython;
	if (pythonProbePromise) return pythonProbePromise;
	pythonProbePromise = (async () => {
		const candidates = executableCandidates(config);
		for (const candidate of candidates) {
			const found = await probePython(candidate);
			if (found !== null) {
				cachedPython = found;
				return found;
			}
		}
		throw new EngineFailure(
			"PYTHON_UNAVAILABLE",
			"No Python interpreter with python-docx, openpyxl and python-pptx was found. Set the plugin's pythonPath config or the OFFICEFORGE_PYTHON environment variable.",
			{ candidates }
		);
	})().finally(() => { pythonProbePromise = undefined; });
	return pythonProbePromise;
}

// ---------------------------------------------------------------------------
// Filesystem helpers
// ---------------------------------------------------------------------------

async function sessionCwd(ctx, exec) {
	try {
		if (ctx.workingDirectory?.ensure && exec.agent) {
			const ensured = await ctx.workingDirectory.ensure(exec.agent, exec.signal);
			if (typeof ensured === "string" && ensured !== "") return resolve(ensured);
		}
	} catch {
		// Fall through to the session header; the service may be absent in a
		// minimal host while tools still execute against the session cwd.
	}
	const cwd = exec.agent?.session?.header?.cwd;
	if (typeof cwd !== "string" || cwd === "") {
		throw new EngineFailure("NO_WORKSPACE", "Office Forge requires an active session with a working directory.");
	}
	return resolve(cwd);
}

async function effectiveRoot(ctx, exec) {
	const root = await sessionCwd(ctx, exec);
	if (ctx.fs?.sandboxMode === undefined) return root;
	try {
		const policy = typeof ctx.get === "function" ? ctx.get("sandboxPolicy") : undefined;
		const resolved = policy?.resolve?.(exec.agent ? { session: exec.agent.session } : {});
		if (resolved?.workspaceRoot) return resolve(resolved.workspaceRoot);
	} catch {
		// Keep the session root; the ctx.fs.contains check below is authoritative.
	}
	return root;
}

function assertLexical(root, candidate) {
	const rel = relative(root, candidate);
	if (rel === "" || rel.startsWith("..") || isAbsolute(rel)) {
		throw new EngineFailure("PATH_OUTSIDE_WORKSPACE", `path "${candidate}" escapes the session workspace "${root}"`);
	}
}

async function resolveTarget(ctx, exec, rawPath, allowedExtensions, mustExist) {
	if (typeof rawPath !== "string" || rawPath.trim() === "") {
		throw new EngineFailure("INVALID_ARGUMENT", "path must be a non-empty string");
	}
	const root = await effectiveRoot(ctx, exec);
	const candidate = resolve(isAbsolute(rawPath) ? rawPath : join(root, rawPath));
	assertLexical(root, candidate);
	const extension = extname(candidate).toLowerCase();
	if (allowedExtensions.length > 0 && !allowedExtensions.includes(extension)) {
		throw new EngineFailure("UNSUPPORTED_FORMAT", `expected ${allowedExtensions.join(" or ")} file, got "${extension || "(none)"}"`);
	}
	const rootTarget = await ctx.fs.resolve(".", { cwd: root, signal: exec.signal });
	const target = await ctx.fs.resolve(rawPath, { cwd: root, signal: exec.signal });
	if (!ctx.fs.contains(rootTarget, target)) {
		throw new EngineFailure("PATH_OUTSIDE_WORKSPACE", `path "${rawPath}" escapes the session workspace "${root}"`);
	}
	const absolute = ctx.fs.processPath(target);
	if (mustExist) {
		const info = await ctx.fs.stat(target, exec.signal);
		if (info === undefined) throw new EngineFailure("INPUT_NOT_FOUND", `"${absolute}" does not exist`);
		if (info.type !== "file") throw new EngineFailure("INVALID_ARGUMENT", `"${absolute}" is not a regular file`);
	}
	const rel = relative(root, absolute);
	return { target, absolute, display: rel !== "" && !rel.startsWith("..") && !isAbsolute(rel) ? rel : absolute, root };
}

// ---------------------------------------------------------------------------
// Engine transport
// ---------------------------------------------------------------------------

function runEngine(config, payload, exec, workspace, extraEnv = {}) {
	return new Promise(async (resolveRun, rejectRun) => {
		let python;
		try {
			python = await resolvePython(config);
		} catch (error) {
			rejectRun(error);
			return;
		}
		let settled = false;
		const finish = (fn, value) => {
			if (settled) return;
			settled = true;
			fn(value);
		};
		const child = spawn(python, [ENGINE], {
			cwd: workspace,
			windowsHide: true,
			stdio: ["pipe", "pipe", "pipe"],
			env: {
				...process.env,
				PYTHONUTF8: "1",
				PYTHONIOENCODING: "utf-8",
				OFFICEFORGE_WORKSPACE: workspace,
				...(config.comPythonPath ? { OFFICEFORGE_COM_PYTHON: config.comPythonPath } : {}),
				...extraEnv
			}
		});
		const stdout = [];
		const stderr = [];
		let stdoutBytes = 0;
		let stderrBytes = 0;
		const timer = setTimeout(() => {
			try { child.kill(); } catch { /* ignore */ }
			finish(rejectRun, new EngineFailure("ENGINE_TIMEOUT", `Office engine exceeded ${config.engineTimeoutMs} ms`));
		}, Math.max(1000, config.engineTimeoutMs));
		const onAbort = () => {
			try { child.kill(); } catch { /* ignore */ }
			finish(rejectRun, new EngineFailure("ABORTED", "Office operation was cancelled"));
		};
		if (exec.signal) {
			if (exec.signal.aborted) {
				onAbort();
				return;
			}
			exec.signal.addEventListener("abort", onAbort, { once: true });
		}
		child.stdout.on("data", (chunk) => {
			stdoutBytes += chunk.length;
			if (stdoutBytes <= MAX_ENGINE_STDOUT) stdout.push(chunk);
		});
		child.stderr.on("data", (chunk) => {
			stderrBytes += chunk.length;
			if (stderrBytes <= MAX_ENGINE_STDERR) stderr.push(chunk);
		});
		child.on("error", (error) => {
			clearTimeout(timer);
			if (exec.signal) exec.signal.removeEventListener("abort", onAbort);
			finish(rejectRun, new EngineFailure("ENGINE_SPAWN_FAILED", `cannot start Office engine: ${error.message}`));
		});
		child.on("close", (code) => {
			clearTimeout(timer);
			if (exec.signal) exec.signal.removeEventListener("abort", onAbort);
			const raw = Buffer.concat(stdout).toString("utf8").trim();
			const errorText = Buffer.concat(stderr).toString("utf8").trim();
			if (raw === "") {
				finish(rejectRun, new EngineFailure("ENGINE_EMPTY", `Office engine returned no output (exit ${code ?? "unknown"})${errorText ? `: ${errorText.slice(-2000)}` : ""}`));
				return;
			}
			let parsed;
			try {
				parsed = JSON.parse(raw);
			} catch (error) {
				finish(rejectRun, new EngineFailure("ENGINE_BAD_JSON", `Office engine returned invalid JSON: ${error.message}`, { raw: raw.slice(0, 2000), stderr: errorText.slice(-2000) }));
				return;
			}
			if (parsed?.ok === true) {
				const normalized = lossless(parsed.result);
				if (normalized === undefined) {
					finish(rejectRun, new EngineFailure("ENGINE_BAD_OUTPUT", "Office engine returned a result that is not JSON data", { raw: raw.slice(0, 2000) }));
					return;
				}
				finish(resolveRun, normalized);
				return;
			}
			const failure = parsed?.error ?? {};
			finish(rejectRun, new EngineFailure(failure.code ?? "ENGINE_ERROR", failure.message ?? "Office engine failed", failure.detail));
		});
		try {
			child.stdin.write(JSON.stringify(payload));
			child.stdin.end();
		} catch (error) {
			finish(rejectRun, new EngineFailure("ENGINE_WRITE_FAILED", `cannot send request to Office engine: ${error.message}`));
		}
	});
}

// ---------------------------------------------------------------------------
// Tool schemas
// ---------------------------------------------------------------------------

const FILE_RESULT_SCHEMA = {
	type: "object",
	additionalProperties: true,
	properties: {
		type: { type: "string" },
		path: { type: "string" },
		bytes: { type: "integer" }
	}
};

function textBlock(text) {
	return [{ type: "text", text: String(text) }];
}

function renderRead(_args, value) {
	if (value && typeof value.text === "string") {
		const suffix = value.truncated ? `\n\n[truncated; ${value.totalChars} characters total]` : "";
		return textBlock(value.text + suffix);
	}
	return textBlock(JSON.stringify(value, null, 2));
}

function renderJson(_args, value) {
	const summary = value && typeof value === "object" ? JSON.stringify(value, null, 2) : String(value);
	return textBlock(summary);
}

/**
 * Strip `undefined` members and non-finite numbers from a tool result.
 *
 * The host validates every tool result with `snapshotJsonValue`, which rejects
 * `undefined` object members, sparse arrays, and non-finite numbers outright
 * instead of dropping them the way `JSON.stringify` would. Optional fields that
 * a backend happens to omit (a preview image with no `page`, a conversion with
 * no `pageCount`) therefore fail the whole tool call with
 * "value is not lossless JSON". Normalising here keeps every declared output
 * field honest without inventing placeholder values.
 */
function lossless(value) {
	if (value === null) return null;
	if (typeof value === "number") return Number.isFinite(value) ? value : null;
	if (typeof value === "string" || typeof value === "boolean") return value;
	if (typeof value !== "object") return undefined;
	if (Array.isArray(value)) {
		const items = [];
		for (const entry of value) {
			const cleaned = lossless(entry);
			if (cleaned !== undefined) items.push(cleaned);
		}
		return items;
	}
	const proto = Object.getPrototypeOf(value);
	if (proto !== Object.prototype && proto !== null) return undefined;
	const result = {};
	for (const [key, entry] of Object.entries(value)) {
		const cleaned = lossless(entry);
		if (cleaned !== undefined) result[key] = cleaned;
	}
	return result;
}

// ---------------------------------------------------------------------------
// Tool registration
// ---------------------------------------------------------------------------

function registerRead(ctx, config) {
	const types = readTypes(config);
	if (types.length === 0) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_read",
		description: "Read a Word (.docx), Excel (.xlsx), PowerPoint (.pptx) or PDF (.pdf) document from the session workspace. Returns readable text/markdown, a structural outline, or search hits. Use officeforge_create / officeforge_edit for changes.",
		parameters: {
			path: { type: "string", required: true, description: "Document path relative to the session workspace or an absolute path inside it." },
			type: { type: "string", enum: types, description: "Document type. Omit to infer from the file extension." },
			mode: { type: "string", enum: ["content", "outline", "search"], description: "content (default), outline (headings/sheets/slides), or search." },
			format: { type: "string", enum: ["markdown", "text", "json"], description: "Output format for content mode. Defaults to markdown." },
			search: { type: "string", description: "Literal search text for mode=search." },
			matchCase: { type: "boolean", description: "Case-sensitive search. Defaults to false." },
			sheet: { type: "string", description: "Excel only: read one worksheet by name." },
			maxRows: { type: "integer", description: "Excel: max rows per sheet (default 200)." },
			maxCols: { type: "integer", description: "Excel: max columns per sheet (default 100)." },
			maxSlides: { type: "integer", description: "PowerPoint: max slides (default 200)." },
			pages: { type: "array", items: { type: "integer" }, description: "PDF: page numbers to read (1-based). Omit for all pages." },
			maxPages: { type: "integer", description: "PDF: maximum pages to read (default 50)." },
			password: { type: "string", description: "PDF: password for an encrypted file." },
			maxChars: { type: "integer", description: "Maximum characters returned (default 200000)." },
			includeFormulas: { type: "boolean", description: "Excel: include formulas instead of cached values." },
			includeMetadata: { type: "boolean", description: "Word: include core document properties." },
			headerRow: { type: "boolean", description: "Excel markdown: treat the first row as a header. Defaults to true." }
		},
		output: { schema: { type: "object", additionalProperties: true }, render: renderRead },
		presentCall: (args) => ({ card: "generic", title: `Read ${args.path}`, kind: "read", locations: [{ path: args.path }] }),
		async execute(args, exec) {
			const target = await resolveTarget(ctx, exec, args.path, readExtensionsFor(config), true);
			const result = await runEngine(config, {
				action: "read",
				path: target.absolute,
				type: args.type,
				args: {
					mode: args.mode,
					format: args.format,
					search: args.search,
					matchCase: args.matchCase,
					sheet: args.sheet,
					maxRows: args.maxRows,
					maxCols: args.maxCols,
					maxSlides: args.maxSlides,
					pages: args.pages,
					maxPages: args.maxPages,
					password: args.password,
					maxChars: args.maxChars,
					includeFormulas: args.includeFormulas,
					includeMetadata: args.includeMetadata,
					headerRow: args.headerRow
				}
			}, exec, target.root);
			return { ...result, path: target.display };
		}
	}));
}

function registerCreate(ctx, config) {
	const types = enabledTypes(config);
	if (types.length === 0) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_create",
		description: "Create a new Word (.docx), Excel (.xlsx) or PowerPoint (.pptx) document from a declarative spec. The spec shape depends on the type: docx uses title/subtitle/sections; xlsx uses sheets, or import {source, format, sheet, header} to load CSV/TSV/JSON (xlsx charts support type 'combo' with primary/secondary axes); pptx uses slideSize/slides. Existing files are refused unless overwrite is true.",
		parameters: {
			path: { type: "string", required: true, description: "Output path ending in .docx, .xlsx or .pptx." },
			type: { type: "string", enum: types, description: "Document type. Omit to infer from the extension." },
			spec: { type: "object", additionalProperties: true, required: true, description: "Declarative document specification." },
			overwrite: { type: "boolean", description: "Replace an existing file. Defaults to false." }
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Create ${args.path}`, kind: "edit", locations: [{ path: args.path }] }),
		async execute(args, exec) {
			const target = await resolveTarget(ctx, exec, args.path, allowedExtensionsFor(config), false);
			const result = await runEngine(config, {
				action: "create",
				path: target.absolute,
				type: args.type,
				spec: args.spec,
				features: featuresFor(config),
				overwrite: args.overwrite ?? false
			}, exec, target.root);
			return { ...result, path: target.display };
		}
	}));
}

function registerEdit(ctx, config) {
	const types = enabledTypes(config);
	if (types.length === 0) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_edit",
		description: "Edit an existing Word, Excel or PowerPoint document with a bounded list of operations. docx supports replace/append/addImage/addToc/addCaption/setStyle/updateFields/addComment/trackRevision/mailMerge/addFootnote/addBookmark/addCrossReference/addCitation/addBibliography/setColumns/addWatermark/addContentControl/saveAsTemplate/setCore; xlsx supports setCells/appendRows/replace/createSheet/deleteSheet/renameSheet/setColumnWidths/setFreeze/setAutoFilter/addChart/addConditionalFormatting/addDataValidation/addComment/mergeCells/addImage/addFormulaColumn/sortRange/removeDuplicates/fillBlanks/cleanRange/pivotSummary/addTable/protectSheet/protectWorkbook; pptx supports replace/appendSlides/setNotes/setNotesBatch/addSection/listSections/deleteSlide/moveSlide/setTransition/addShape/addTextBox/addConnector/addAnimation/addSmartArt/addComment/listComments/addMedia/applyMaster/applyTheme/setCore. A backup is made by default.",
		parameters: {
			path: { type: "string", required: true, description: "Document path to edit." },
			type: { type: "string", enum: types, description: "Document type. Omit to infer from the extension." },
			operations: { type: "array", required: true, items: { type: "object", additionalProperties: true }, description: "Ordered edit operations (1-200 per call)." },
			backup: { type: "boolean", description: "Write a sibling backup before saving. Defaults to true." }
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Edit ${args.path}`, kind: "edit", locations: [{ path: args.path }] }),
		async execute(args, exec) {
			if (!Array.isArray(args.operations) || args.operations.length === 0) {
				throw new EngineFailure("INVALID_ARGUMENT", "operations must contain at least one edit operation");
			}
			if (args.operations.length > MAX_EDIT_OPERATIONS) {
				throw new EngineFailure("LIMIT_EXCEEDED", `too many edit operations (max ${MAX_EDIT_OPERATIONS})`);
			}
			const target = await resolveTarget(ctx, exec, args.path, allowedExtensionsFor(config), true);
			const result = await runEngine(config, {
				action: "edit",
				path: target.absolute,
				type: args.type,
				args: { operations: args.operations, backup: args.backup ?? true },
				features: featuresFor(config)
			}, exec, target.root);
			return { ...result, path: target.display };
		}
	}));
}

function registerTemplate(ctx, config) {
	const types = enabledTypes(config);
	if (types.length === 0) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_template",
		description: "Fill a Word, Excel or PowerPoint template with JSON data. Replaces {{key}} placeholders; Word and Excel additionally expand one-level {{#each items}} ... {{/each}} table/row loops. Supports missing=keep|empty|error and atomic output.",
		parameters: {
			template: { type: "string", required: true, description: "Template document path." },
			output: { type: "string", required: true, description: "Output document path." },
			type: { type: "string", enum: types, description: "Document type. Omit to infer from the template extension." },
			data: { type: "object", additionalProperties: true, description: "Inline JSON data for placeholders." },
			dataPath: { type: "string", description: "Optional .json/.csv/.tsv data file inside the workspace; ignored when data is supplied." },
			options: {
				type: "object",
				additionalProperties: true,
				properties: {
					expandLoops: { type: "boolean", description: "Expand {{#each}} loops. Defaults to true." },
					missing: { type: "string", enum: ["keep", "empty", "error"], description: "Missing placeholder behavior. Defaults to keep." },
					overwrite: { type: "boolean", description: "Replace an existing output. Defaults to false." },
					backup: { type: "boolean", description: "Back up an existing output. Defaults to false." }
				}
			}
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Fill template ${args.template}`, kind: "edit", locations: [{ path: args.output }] }),
		async execute(args, exec) {
			const template = await resolveTarget(ctx, exec, args.template, allowedExtensionsFor(config), true);
			const output = await resolveTarget(ctx, exec, args.output, allowedExtensionsFor(config), false);
			if (args.data === undefined && typeof args.dataPath !== "string") {
				throw new EngineFailure("INVALID_ARGUMENT", "officeforge_template requires either data or dataPath");
			}
			let dataPath;
			if (args.data === undefined && args.dataPath) {
				const resolvedData = await resolveTarget(ctx, exec, args.dataPath, [".json", ".csv", ".tsv"], true);
				dataPath = resolvedData.absolute;
			}
			const result = await runEngine(config, {
				action: "template",
				template: template.absolute,
				output: output.absolute,
				type: args.type,
				data: args.data,
				dataPath,
				options: args.options ?? {}
			}, exec, template.root);
			return { ...result, path: output.display };
		}
	}));
}

function registerBatch(ctx, config) {
  const types = enabledTypes(config);
  if (types.length === 0) return undefined;
  return ctx.tools.register(defineTool({
    name: "officeforge_batch",
    description: "Run several independent create/edit/template operations in one Python engine call. Batch items never nest and never call another tool. Each item reports success or failure independently; one bad item does not roll back the others.",
    parameters: {
      items: {
        type: "array",
        required: true,
        items: { type: "object", additionalProperties: true },
        description: "Items shaped like {id, action:'create'|'edit'|'template', path, type, spec?, args?, template?, output?, data?, options?}."
      }
    },
    output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
    presentCall: () => ({ card: "generic", title: "Office Forge batch", kind: "edit", locations: [] }),
    async execute(args, exec) {
      const items = Array.isArray(args.items) ? args.items : [];
      if (items.length === 0) throw new EngineFailure("INVALID_ARGUMENT", "items must contain at least one operation");
      if (items.length > config.maxBatchItems) throw new EngineFailure("LIMIT_EXCEEDED", `too many batch items (max ${config.maxBatchItems})`);
      const cwd = await effectiveRoot(ctx, exec);
      const prepared = [];
      const skipped = [];
      for (const [index, rawItem] of items.entries()) {
        const item = rawItem && typeof rawItem === "object" ? rawItem : {};
        const itemId = String(item.id ?? index + 1);
        try {
          if (item.action === "batch") throw new EngineFailure("INVALID_OPERATION", "nested batch is not allowed");
          if (item.action === "create") {
            const target = await resolveTarget(ctx, exec, item.path, allowedExtensionsFor(config), false);
            prepared.push({ id: itemId, action: "create", path: target.absolute, type: item.type, spec: item.spec, overwrite: item.overwrite ?? false });
          } else if (item.action === "edit") {
            const target = await resolveTarget(ctx, exec, item.path, allowedExtensionsFor(config), true);
            prepared.push({ id: itemId, action: "edit", path: target.absolute, type: item.type, args: item.args ?? {} });
          } else if (item.action === "template") {
            const template = await resolveTarget(ctx, exec, item.template, allowedExtensionsFor(config), true);
            const output = await resolveTarget(ctx, exec, item.output, allowedExtensionsFor(config), false);
            prepared.push({ id: itemId, action: "template", template: template.absolute, output: output.absolute, type: item.type, data: item.data, dataPath: item.dataPath, options: item.options ?? {} });
          } else {
            skipped.push({ id: itemId, ok: false, error: { code: "UNSUPPORTED_OPERATION", message: "batch supports create, edit and template only" } });
          }
        } catch (error) {
          // A bad item must not prevent its siblings from running.
          skipped.push({
            id: itemId,
            ok: false,
            error: {
              code: error instanceof EngineFailure ? error.code : "INVALID_ARGUMENT",
              message: error instanceof Error ? error.message : String(error)
            }
          });
        }
      }
      if (prepared.length === 0) return lossless({ items: skipped, succeeded: 0, failed: skipped.length });
      const result = await runEngine(config, { action: "batch", items: prepared, features: featuresFor(config) }, exec, cwd);
      return lossless({
        ...result,
        succeeded: Number(result.succeeded ?? 0),
        failed: Number(result.failed ?? 0) + skipped.length,
        items: [...skipped, ...(result.items ?? [])]
      });
    }
  }));
}

function registerAudit(ctx, config) {
	const types = auditTypes(config);
	if (types.length === 0) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_audit",
		description: "Analyze a Word document, Excel workbook, or PowerPoint deck without rewriting it. docx modes: health (empty documents, heading jumps, duplicate headings, orphan headings, unresolved placeholders, long paragraphs, empty-paragraph runs), profile (paragraph/heading/table/image/section counts, style usage, page setup, metadata). xlsx modes: health (formula errors, inconsistent formula patterns, hardcoded values in formula columns, blank row breaks, circular references), profile (sheets, formulas, charts, merged ranges, validations, comments), insight (numeric column statistics and text value frequencies), diff (compare with comparePath). pptx modes: deck (empty/missing-title slides, dense text, long titles, small fonts, inconsistent title fonts), profile, themes (list built-in theme presets).",
		parameters: {
			path: { type: "string", description: "Document, workbook, or presentation path. Optional for pptx mode themes." },
			type: { type: "string", enum: types, description: "Document type. Omit to infer from the extension." },
			mode: { type: "string", enum: ["health", "profile", "insight", "diff", "deck", "themes", "layouts", "templates"], description: "Analysis mode. docx/xlsx default to health; pptx defaults to deck. pptx also exposes themes/layouts/templates catalogs without a path." },
			comparePath: { type: "string", description: "Second workbook path for xlsx diff mode." },
			maxIssues: { type: "integer", description: "Maximum issues to return. Defaults to 200." },
			maxRows: { type: "integer", description: "Maximum rows per sheet for insight mode. Defaults to 2000." },
			maxChanges: { type: "integer", description: "Maximum changed cells to return for diff mode. Defaults to 200." }
		},
		output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Audit ${args.path}`, kind: "read", locations: [{ path: args.path }] }),
		async execute(args, exec) {
			const catalogMode = ["themes", "layouts", "templates"].includes(args.mode);
			const catalogOnly = catalogMode && (!args.path || String(args.path).trim() === "");
			let target;
			if (catalogOnly) {
				const root = await ctx.workingDirectory.ensure(exec.agent, exec.signal).catch(() => exec.agent?.session?.header?.cwd);
				target = { absolute: undefined, display: undefined, root };
			} else {
				target = await resolveTarget(ctx, exec, args.path, allowedExtensionsFor(config), true);
			}
			let comparePath;
			if (typeof args.comparePath === "string" && args.comparePath.trim() !== "") {
				const compared = await resolveTarget(ctx, exec, args.comparePath, allowedExtensionsFor(config), true);
				comparePath = compared.absolute;
			}
			const result = await runEngine(config, {
				action: "audit",
				path: target.absolute,
				type: args.type,
				mode: args.mode,
				comparePath,
				features: featuresFor(config),
				args: { maxIssues: args.maxIssues, maxRows: args.maxRows, maxChanges: args.maxChanges }
			}, exec, target.root);
			return { ...result, ...(target.display ? { path: target.display } : {}) };
		}
	}));
}

function buildDoctorReport(engine, engineError, config, kit) {
	const backends = engine?.backends ?? {};
	const pythonInfo = engine?.engine ?? {};
	const modules = pythonInfo.modules ?? {};
	const requiredModules = ["python-docx", "openpyxl", "python-pptx"];
	const missingModules = requiredModules.filter((moduleName) => !modules[moduleName] || modules[moduleName] === "unknown");
	const moduleDetail = requiredModules.map((moduleName) => `${moduleName}=${modules[moduleName] ?? "missing"}`).join(", ");
	const checks = [];
	const addCheck = (checkName, required, ok, detail, remedy, statusOverride) => {
		checks.push({
			name: checkName,
			required: Boolean(required),
			ok: Boolean(ok),
			status: statusOverride ?? (ok ? "ok" : (required ? "missing" : "warn")),
			detail,
			...(remedy ? { remedy } : {})
		});
	};

	if (engineError) {
		addCheck("python", true, false, `Python engine failed to start: ${engineError.message ?? String(engineError)}`,
			"Set pythonPath in the plugin config or OFFICEFORGE_PYTHON to a Python that has python-docx, openpyxl and python-pptx.");
	} else {
		addCheck("python", true, Boolean(pythonInfo.executable), pythonInfo.executable ?? "not found",
			"Set pythonPath in the plugin config or OFFICEFORGE_PYTHON to a Python that has python-docx, openpyxl and python-pptx.");
	}
	addCheck("pythonModules", true, missingModules.length === 0, moduleDetail,
		`Install the missing Python packages (${missingModules.join(", ") || "python-docx, openpyxl, python-pptx"}) into the DSH runtime or point pythonPath at a suitable interpreter.`);
	const pdfOk = Boolean(engine?.pdfVendor ?? existsSync(VENDOR_PDF));
	addCheck("pdfVendor", Boolean(config.modules?.pdf), pdfOk,
		engine?.pdfVendorPath ?? (existsSync(VENDOR_PDF) ? VENDOR_PDF : "lib/vendor/pypdf missing"),
		"Reinstall the plugin package; lib/vendor/pypdf is shipped inside the tarball.");
	const ocr = backends.ocr ?? {};
	const ocrOk = Boolean(ocr.available);
	const ocrDetail = ocrOk
		? `pytesseract=${ocr.version ?? "ok"}${Array.isArray(ocr.languages) && ocr.languages.length ? ` | languages=${ocr.languages.join(",")}` : ""} | tesseract=${ocr.tesseractPath ?? "found"}`
		: `pytesseract=${ocr.pytesseract ? "ok" : "missing"}, fitz=${ocr.fitz ? "ok" : "missing"}, tesseract=${ocr.tesseract ? "ok" : "missing"}`;
	addCheck("ocr", false, ocrOk, ocrDetail,
		"Install PyMuPDF and pytesseract into the DSH Python runtime, install Tesseract OCR with the desired language packs, and put tesseract.exe on PATH or set TESSERACT_CMD.");
	const com = backends.com ?? {};
	const comOffice = com.office ?? {};
	const comOfficeOk = Object.values(comOffice).every(Boolean);
	if (!config.enableBackends) {
		addCheck("com", false, true, "disabled by config.enableBackends=false", null, "disabled");
	} else {
		addCheck("com", false, Boolean(com.available),
			com.available ? `${com.python ?? "python"}${Object.keys(comOffice).length ? ` | office: ${JSON.stringify(comOffice)}` : ""}` : (com.python ?? "no Python with pywin32"),
			"Install pywin32 into a system Python and Microsoft Office, or set comPythonPath / OFFICEFORGE_COM_PYTHON; native and legacy conversion stay unavailable without it.");
		if (com.available && Object.keys(comOffice).length > 0) {
			addCheck("officeComApps", false, comOfficeOk, JSON.stringify(comOffice),
				"Install the missing Microsoft Office application, or avoid operations that need that application.");
		}
	}
	const libre = backends.libreoffice ?? {};
	addCheck("libreoffice", false, Boolean(libre.available), libre.path ?? "not found",
		"Install LibreOffice for headless legacy conversion and PDF export; COM can substitute when Office is installed.", config.enableBackends ? undefined : "disabled");
	addCheck("officeKit", Boolean(config.enableExport || config.enableBackends), Boolean(kit), kit ? "available" : "not loaded",
		"Enable the host officeKit service (dsh office-kit) or disable officeforge_export / preview / recalculate.");
	addCheck("client", false, existsSync(CLIENT_BUNDLE), CLIENT_BUNDLE,
		"Reinstall the plugin package; lib/client.js carries the read-only plugin-page status card.");
	const missingRequired = checks.filter((check) => check.required && !check.ok).length;
	const warnings = checks.filter((check) => !check.required && !check.ok && check.status !== "disabled").length;
	return lossless({
		ok: missingRequired === 0,
		health: missingRequired > 0 ? "error" : (warnings > 0 ? "degraded" : "ok"),
		summary: { ok: checks.filter((check) => check.ok).length, warn: warnings, missing: missingRequired },
		plugin: { name, version: PLUGIN_VERSION },
		python: pythonInfo,
		pdfVendor: pdfOk,
		backends,
		officeKit: Boolean(kit),
		config: {
			modules: config.modules,
			features: config.features,
			enableBackends: config.enableBackends,
			enableExport: config.enableExport,
			enableBatch: config.enableBatch
		},
		checks,
		...(engineError ? { error: { code: engineError.code ?? "DOCTOR_FAILED", message: engineError.message ?? String(engineError), detail: engineError.detail } } : {})
	});
}

function registerDoctor(ctx, config) {
	return ctx.tools.register(defineTool({
		name: "officeforge_doctor",
		description: "Diagnose the Office Forge runtime before using it: Python engine version and modules, bundled PDF runtime, OCR backend availability, COM/LibreOffice availability, host officeKit service, client status card, and the active module/feature switches. Each check reports whether the component is required, its detail, and a remedy when it is missing. This tool never reads or writes documents.",
		parameters: {},
		output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
		presentCall: () => ({ card: "generic", title: "Office Forge doctor", kind: "read", locations: [] }),
		async execute(_args, exec) {
			const root = await effectiveRoot(ctx, exec);
			const kit = getService(ctx, "officeKit");
			let engine;
			let engineError;
			try {
				engine = await runEngine(config, { action: "doctor" }, exec, root);
			} catch (error) {
				engineError = error;
			}
			return buildDoctorReport(engine, engineError, config, kit);
		}
	}));
}

function getService(ctx, name) {
	try {
		return typeof ctx.get === "function" ? ctx.get(name) : undefined;
	} catch {
		return undefined;
	}
}

function outputOwner(ctx, agent) {
	// Keep a subagent's converted file alive for the lifetime of the root
	// conversation, matching the host office tools. When the subagent service
	// is absent or denies the lookup, the current session owns its output.
	const subagents = getService(ctx, "subagents");
	if (!subagents || typeof subagents.rootOf !== "function") return agent.session.id;
	try {
		return subagents.rootOf(agent).session.id;
	} catch (error) {
		if (error && error.code === "UNAUTHORIZED") return agent.session.id;
		throw error;
	}
}

function assertSourceUnchanged(ctx, target, info, signal) {
	return Promise.resolve(ctx.fs.stat(target, signal)).then((current) => {
		if (!current || current.type !== "file" || current.version !== info.version) {
			throw new EngineFailure("SOURCE_CHANGED", "source document changed while it was being converted");
		}
	});
}

async function prepareOfficeSource(ctx, exec, target) {
	const info = await ctx.fs.stat(target.target, exec.signal);
	if (!info || info.type !== "file") throw new EngineFailure("INPUT_NOT_FOUND", "source document is not a regular file");
	await ctx.fs.readByteRange(target.target, { offset: 0, length: 1 }, exec.signal);
	await assertSourceUnchanged(ctx, target.target, info, exec.signal);
	return {
		info,
		source: {
			key: JSON.stringify([exec.agent.session.id, ctx.fs.processPath(target.target)]),
			version: info.version,
			bytes: info.size,
			async read(signal, maxBytes) {
				const bytes = await ctx.fs.readBytes(target.target, signal, maxBytes);
				await assertSourceUnchanged(ctx, target.target, info, signal);
				return { bytes, version: info.version };
			}
		}
	};
}

async function copyOfficeKitOutput(ctx, config, exec, converted, outputPath, overwrite) {
	let kitPath = converted.path;
	if (typeof ctx.fs.processPathFromHostPath === "function") {
		const mapped = ctx.fs.processPathFromHostPath(converted.path);
		if (mapped === undefined) throw new EngineFailure("OFFICEKIT_REMOTE_WORKSPACE", "converted files are not available when the workspace is on a remote host");
		kitPath = mapped;
	}
	if (!outputPath) return { path: kitPath, display: kitPath };
	const outputTarget = await resolveTarget(ctx, exec, outputPath, [extname(outputPath).toLowerCase()], false);
	await runEngine(config, {
		action: "copy",
		source: converted.path,
		target: outputTarget.absolute,
		overwrite: overwrite ?? false
	}, exec, outputTarget.root, { OFFICEFORGE_OFFICEKIT_DIR: getService(ctx, "officeKit")?.outputDirectory ?? dirname(converted.path) });
	return { path: outputTarget.absolute, display: outputTarget.display };
}

function registerConvert(ctx, config) {
	if (!config.enableBackends) return undefined;
	if (!config.modules?.word && !config.modules?.excel && !config.modules?.ppt && !config.modules?.pdf) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_convert",
		description: "Convert a legacy .doc/.xls/.ppt file (or a modern Office/PDF file) to another format through the host officeKit service. This is the compatibility bridge for legacy formats; no local Office installation is required by the plugin itself.",
		parameters: {
			path: { type: "string", required: true, description: "Source document path." },
			format: { type: "string", enum: ["docx", "xlsx", "pptx", "pdf"], description: "Target format. Defaults to the modern equivalent of the source extension." },
			output: { type: "string", description: "Optional workspace output path; defaults to the officeKit output store." },
			overwrite: { type: "boolean", description: "Replace an existing output file. Defaults to false." }
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Convert ${args.path}`, kind: "edit", locations: [{ path: args.output ?? args.path }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_convert runs only in an Agent step");
			const kit = getService(ctx, "officeKit");
			if (!kit) throw new EngineFailure("OFFICEKIT_UNAVAILABLE", "The host officeKit service is unavailable.");
			const sourceTarget = await resolveTarget(ctx, exec, args.path, [".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx", ".pptm"], true);
			const prepared = await prepareOfficeSource(ctx, exec, sourceTarget);
			const sourceExtension = extname(sourceTarget.absolute).toLowerCase().replace(/^\./, "");
			const defaultFormat = ({ doc: "docx", docm: "docx", xls: "xlsx", xlsm: "xlsx", ppt: "pptx", pptm: "pptx" })[sourceExtension] ?? sourceExtension;
			const format = String(args.format ?? defaultFormat).toLowerCase();
			const stem = basename(sourceTarget.display).replace(/\.[^.]+$/u, "");
			const converted = await kit.convertTo({
				extension: sourceExtension,
				format,
				source: prepared.source,
				output: { owner: outputOwner(ctx, exec.agent), name: `${stem}.${format}` }
			}, exec.signal);
			const copied = await copyOfficeKitOutput(ctx, config, exec, converted, args.output, args.overwrite);
			return lossless({
				type: format,
				path: copied.display,
				bytes: converted.bytes,
				pageCount: converted.pageCount,
				missingFonts: converted.missingFonts ?? [],
				sourceFormat: sourceExtension,
				backend: "officeKit",
				officeKitPath: converted.path
			});
		}
	}));
}

function registerPreview(ctx, config) {
	if (!config.enableBackends) return undefined;
	if (!config.modules?.word && !config.modules?.excel && !config.modules?.ppt && !config.modules?.pdf) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_preview",
		description: "Render pages, slides or worksheet areas to PNG images through the host officeKit service. This is a file-level visual preview path, not a live editor.",
		parameters: {
			path: { type: "string", required: true, description: "Office or PDF document path." },
			outputDir: { type: "string", description: "Workspace directory for PNG files. Defaults to officeforge-previews." },
			pages: { type: "array", items: { type: "integer" }, description: "Page/slide numbers to render (1-based)." },
			sheet: { type: "string", description: "Excel worksheet name." },
			range: { type: "string", description: "Excel range, for example A1:F20." },
			dpi: { type: "integer", description: "Render DPI. Defaults to the host officeKit setting." },
			maxImages: { type: "integer", description: "Maximum number of images. Defaults to 12." },
			maxPixels: { type: "integer", description: "Per-image pixel limit." },
			maxDimension: { type: "integer", description: "Per-image dimension limit." },
			maxBytes: { type: "integer", description: "Per-image byte limit." }
		},
		output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Preview ${args.path}`, kind: "read", locations: [{ path: args.path }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_preview runs only in an Agent step");
			const kit = getService(ctx, "officeKit");
			if (!kit) throw new EngineFailure("OFFICEKIT_UNAVAILABLE", "The host officeKit service is unavailable.");
			const sourceTarget = await resolveTarget(ctx, exec, args.path, [".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx", ".pptm", ".pdf"], true);
			const prepared = await prepareOfficeSource(ctx, exec, sourceTarget);
			const root = await effectiveRoot(ctx, exec);
			const dirRaw = args.outputDir ?? "officeforge-previews";
			const dirCandidate = resolve(isAbsolute(dirRaw) ? dirRaw : join(root, dirRaw));
			assertLexical(root, dirCandidate);
			const dirTarget = await ctx.fs.resolve(dirRaw, { cwd: root, signal: exec.signal });
			const rootTarget = await ctx.fs.resolve(".", { cwd: root, signal: exec.signal });
			if (!ctx.fs.contains(rootTarget, dirTarget)) throw new EngineFailure("PATH_OUTSIDE_WORKSPACE", "preview output directory escapes the session workspace");
			const dirPath = ctx.fs.processPath(dirTarget);
			await mkdir(dirPath, { recursive: true });
			const sourceExtension = extname(sourceTarget.absolute).toLowerCase().replace(/^\./, "");
			const rendered = await kit.renderImages({
				extension: sourceExtension,
				source: prepared.source,
				pages: args.pages,
				sheet: args.sheet,
				range: args.range,
				dpi: args.dpi,
				maxImages: args.maxImages ?? 12,
				maxPixels: args.maxPixels,
				maxDimension: args.maxDimension,
				maxBytes: args.maxBytes
			}, exec.signal);
			const images = [];
			const stem = basename(sourceTarget.display).replace(/\.[^.]+$/u, "");
			const usedNames = new Set();
			for (let index = 0; index < rendered.images.length; index += 1) {
				const image = rendered.images[index];
				// Two documents with the same stem but different formats (report.docx
				// and report.pptx) must not overwrite each other's previews, so the
				// source extension is part of the name. Non-ASCII sheet names were
				// previously flattened to underscores; keep Unicode letters and digits
				// instead, and fall back to the ordinal when nothing usable remains.
				// Compose the suffix from every locator the renderer returned:
				// page for documents/decks, sheet+range for spreadsheets. Using only
				// one of them made "same sheet, different range" previews overwrite
				// each other.
				const suffixParts = [];
				if (image.page !== undefined && image.page !== null) suffixParts.push(String(image.page));
				if (image.sheet !== undefined && image.sheet !== null) suffixParts.push(String(image.sheet));
				if (image.range !== undefined && image.range !== null) suffixParts.push(String(image.range).split(":").join("-"));
				const rawSuffix = suffixParts.length > 0 ? suffixParts.join("-") : `${index + 1}`;
				const safeSuffix = String(rawSuffix).replace(/[^\p{L}\p{N}_-]+/gu, "_").replace(/^_+|_+$/gu, "");
				const suffix = safeSuffix === "" ? `${index + 1}` : safeSuffix;
				let name = `${stem}-${sourceExtension}-${suffix}.png`;
				for (let attempt = 2; usedNames.has(name); attempt += 1) {
					name = `${stem}-${sourceExtension}-${suffix}-${attempt}.png`;
				}
				usedNames.add(name);
				const fullPath = join(dirPath, name);
				await writeFile(fullPath, Buffer.from(image.png));
				images.push({ path: relative(root, fullPath).replace(/\\/g, "/"), width: image.width, height: image.height, page: image.page, sheet: image.sheet, range: image.range });
			}
			return lossless({ type: "preview", path: args.path, pageCount: rendered.pageCount, dpi: rendered.dpi, missingFonts: rendered.missingFonts ?? [], images });
		}
	}));
}

function registerRecalculate(ctx, config) {
	if (!config.enableBackends || !config.modules?.excel) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_recalculate",
		description: "Recalculate every formula in an Excel workbook through the host officeKit service and save a new workbook with refreshed cached results. This uses the host's real spreadsheet engine rather than a formula-text approximation.",
		parameters: {
			path: { type: "string", required: true, description: "Source workbook path (.xlsx/.xlsm)." },
			output: { type: "string", description: "Optional workspace output path; defaults to the officeKit output store." },
			overwrite: { type: "boolean", description: "Replace an existing output file. Defaults to false." }
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Recalculate ${args.path}`, kind: "edit", locations: [{ path: args.output ?? args.path }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_recalculate runs only in an Agent step");
			const kit = getService(ctx, "officeKit");
			if (!kit) throw new EngineFailure("OFFICEKIT_UNAVAILABLE", "The host officeKit service is unavailable.");
			const sourceTarget = await resolveTarget(ctx, exec, args.path, [".xlsx", ".xlsm", ".xls"], true);
			const prepared = await prepareOfficeSource(ctx, exec, sourceTarget);
			const extension = extname(sourceTarget.absolute).toLowerCase().replace(/^\./, "");
			const stem = basename(sourceTarget.display).replace(/\.[^.]+$/u, "");
			const converted = await kit.recalculate({
				extension,
				source: prepared.source,
				format: "xlsx",
				output: { owner: outputOwner(ctx, exec.agent), name: `${stem}-recalculated.xlsx` }
			}, exec.signal);
			const copied = await copyOfficeKitOutput(ctx, config, exec, converted, args.output, args.overwrite);
			return lossless({
				type: "xlsx",
				path: copied.display,
				bytes: converted.bytes,
				missingFonts: converted.missingFonts ?? [],
				backend: "officeKit",
				officeKitPath: converted.path
			});
		}
	}));
}

function registerPdf(ctx, config) {
	if (!config.modules?.pdf) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_pdf",
		description: "File-level PDF toolkit: merge, split, rotate, watermark, page numbers, form field listing/filling, image extraction, optional OCR, and password encrypt/decrypt. Uses the bundled pure-Python pypdf runtime; OCR additionally requires PyMuPDF + pytesseract + the tesseract executable.",
		parameters: {
			operation: { type: "string", required: true, enum: ["merge", "split", "rotate", "watermark", "pageNumbers", "formFields", "fillForm", "extractImages", "ocr", "encrypt", "decrypt"], description: "PDF operation to perform." },
			inputs: { type: "array", items: { type: "string" }, description: "Input PDF paths for merge." },
			path: { type: "string", description: "Input PDF path." },
			output: { type: "string", description: "Output PDF path for merge, rotate, watermark, pageNumbers, fillForm, encrypt or decrypt." },
			outputDir: { type: "string", description: "Output directory for split or extractImages." },
			pages: { type: "array", items: { type: "integer" }, description: "Optional 1-based page numbers." },
			everyPage: { type: "boolean", description: "For split: one output file per page." },
			angle: { type: "integer", description: "Rotation angle in degrees, multiple of 90. Defaults to 90." },
			watermarkPath: { type: "string", description: "Watermark PDF path for watermark." },
			format: { type: "string", description: "Page number template, e.g. '{page} / {total}'." },
			position: { type: "string", description: "Page number position: bottom-center, bottom-right, bottom-left, top-center, top-right, top-left." },
			start: { type: "integer", description: "First page number. Defaults to 1." },
			fontSize: { type: "number", description: "Page number font size. Defaults to 10." },
			margin: { type: "number", description: "Page number margin in points. Defaults to 24." },
			fields: { type: "object", additionalProperties: true, description: "For fillForm: object mapping form field names to values." },
			flatten: { type: "boolean", description: "For fillForm: flatten fields after filling." },
			language: { type: "string", description: "OCR language, e.g. eng or eng+chi_sim. Defaults to eng." },
			dpi: { type: "integer", description: "OCR render DPI. Defaults to 200." },
			password: { type: "string", description: "PDF password for encrypted input." },
			userPassword: { type: "string", description: "For encrypt: user password." },
			ownerPassword: { type: "string", description: "For encrypt: owner password." },
			overwrite: { type: "boolean", description: "Replace existing output files. Defaults to false." }
		},
		output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `PDF ${args.operation}`, kind: "edit", locations: [{ path: args.output ?? args.outputDir ?? args.path ?? (args.inputs && args.inputs[0]) ?? "." }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_pdf runs only in an Agent step");
			const root = await effectiveRoot(ctx, exec);
			const operation = String(args.operation).toLowerCase();
			const payload = { action: "pdfop", operation, overwrite: args.overwrite ?? false, pages: args.pages, everyPage: args.everyPage, angle: args.angle, password: args.password };
			if (operation === "merge") {
				if (!Array.isArray(args.inputs) || args.inputs.length === 0) throw new EngineFailure("INVALID_ARGUMENT", "inputs must be a non-empty array for merge");
				payload.inputs = [];
				for (const item of args.inputs) {
					const target = await resolveTarget(ctx, exec, item, [".pdf"], true);
					payload.inputs.push(target.absolute);
				}
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				payload.output = output.absolute;
			} else if (operation === "split") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const directory = await resolveTarget(ctx, exec, args.outputDir ?? "officeforge-pdf", [], false);
				payload.path = source.absolute;
				payload.outputDir = directory.absolute;
			} else if (operation === "rotate" || operation === "watermark") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				payload.path = source.absolute;
				payload.output = output.absolute;
				if (operation === "watermark") {
					const overlay = await resolveTarget(ctx, exec, args.watermarkPath, [".pdf"], true);
					payload.watermarkPath = overlay.absolute;
				}
			} else if (operation === "pagenumbers") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				payload.path = source.absolute;
				payload.output = output.absolute;
				payload.format = args.format;
				payload.position = args.position;
				payload.start = args.start;
				payload.fontSize = args.fontSize;
				payload.margin = args.margin;
			} else if (operation === "formfields") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				payload.path = source.absolute;
			} else if (operation === "fillform") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				if (!args.fields || typeof args.fields !== "object" || Array.isArray(args.fields)) throw new EngineFailure("INVALID_ARGUMENT", "fields must be an object for fillForm");
				payload.path = source.absolute;
				payload.output = output.absolute;
				payload.fields = args.fields;
				payload.flatten = args.flatten ?? false;
			} else if (operation === "extractimages") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const directory = await resolveTarget(ctx, exec, args.outputDir ?? "officeforge-pdf-images", [], false);
				payload.path = source.absolute;
				payload.outputDir = directory.absolute;
			} else if (operation === "ocr") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				payload.path = source.absolute;
				payload.language = args.language;
				payload.dpi = args.dpi;
			} else if (operation === "encrypt") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				payload.path = source.absolute;
				payload.output = output.absolute;
				payload.userPassword = args.userPassword ?? "";
				payload.ownerPassword = args.ownerPassword ?? "";
			} else if (operation === "decrypt") {
				const source = await resolveTarget(ctx, exec, args.path, [".pdf"], true);
				const output = await resolveTarget(ctx, exec, args.output, [".pdf"], false);
				payload.path = source.absolute;
				payload.output = output.absolute;
			} else {
				throw new EngineFailure("UNSUPPORTED_OPERATION", `unsupported PDF operation: ${operation}`);
			}
			return runEngine(config, payload, exec, root);
		}
	}));
}

function registerNative(ctx, config) {
	if (!config.enableBackends) return undefined;
	if (!config.modules?.word && !config.modules?.excel && !config.modules?.ppt) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_native",
		description: "Use the installed Microsoft Office COM backend for operations that need a real Office application: convert, export PDF, accept tracked revisions, recalculate formulas, create a native PivotTable, or add native SmartArt. Requires Windows with Office and a Python that has pywin32; returns BACKEND_UNAVAILABLE otherwise.",
		parameters: {
			operation: { type: "string", required: true, enum: ["convert", "exportPdf", "acceptRevisions", "recalculate", "pivot", "smartArt"], description: "Native Office operation." },
			path: { type: "string", required: true, description: "Source Office document path." },
			output: { type: "string", required: true, description: "Output file path." },
			spec: { type: "object", additionalProperties: true, description: "Operation-specific options, for example pivot source/data fields or SmartArt layout/text." },
			timeout: { type: "integer", description: "COM operation timeout in seconds. Defaults to 120." },
			overwrite: { type: "boolean", description: "Replace an existing output file. Defaults to false." }
		},
		output: { schema: { type: "object", additionalProperties: true }, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Native ${args.operation}`, kind: "edit", locations: [{ path: args.output }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_native runs only in an Agent step");
			const root = await effectiveRoot(ctx, exec);
			const source = await resolveTarget(ctx, exec, args.path, [".doc", ".docm", ".docx", ".xls", ".xlsm", ".xlsx", ".ppt", ".pptm", ".pptx"], true);
			const output = await resolveTarget(ctx, exec, args.output, [], false);
			return runEngine(config, {
				action: "native",
				operation: args.operation,
				path: source.absolute,
				output: output.absolute,
				spec: args.spec ?? {},
				timeout: args.timeout,
				overwrite: args.overwrite ?? false
			}, exec, root);
		}
	}));
}

function registerExport(ctx, config) {
	if (!config.modules?.pdf) return undefined;
	return ctx.tools.register(defineTool({
		name: "officeforge_export",
		description: "Convert a Word, Excel or PowerPoint document to PDF (or another office format) through the host officeKit service. Optionally copy the converted file into the session workspace.",
		parameters: {
			path: { type: "string", required: true, description: "Source document path." },
			output: { type: "string", description: "Optional workspace output path; defaults to the officeKit output directory." },
			format: { type: "string", enum: ["pdf", "docx", "xlsx", "pptx", "odt", "ods", "odp", "csv", "txt"], description: "Target format. Defaults to pdf." },
			overwrite: { type: "boolean", description: "Replace an existing output file. Defaults to false." }
		},
		output: { schema: FILE_RESULT_SCHEMA, render: renderJson },
		presentCall: (args) => ({ card: "generic", title: `Export ${args.path}`, kind: "edit", locations: [{ path: args.output ?? args.path }] }),
		async execute(args, exec) {
			if (!exec.agent) throw new EngineFailure("NO_SESSION", "officeforge_export runs only in an Agent step");
			const kit = getService(ctx, "officeKit");
			if (!kit) throw new EngineFailure("OFFICEKIT_UNAVAILABLE", "The host officeKit service is unavailable.");
			const sourceTarget = await resolveTarget(ctx, exec, args.path, allowedExtensionsFor(config), true);
			const info = await ctx.fs.stat(sourceTarget.target, exec.signal);
			if (!info || info.type !== "file") throw new EngineFailure("INPUT_NOT_FOUND", "source document is not a regular file");
			await ctx.fs.readByteRange(sourceTarget.target, { offset: 0, length: 1 }, exec.signal);
			await assertSourceUnchanged(ctx, sourceTarget.target, info, exec.signal);
			const source = {
				key: JSON.stringify([exec.agent.session.id, ctx.fs.processPath(sourceTarget.target)]),
				version: info.version,
				bytes: info.size,
				async read(signal, maxBytes) {
					const bytes = await ctx.fs.readBytes(sourceTarget.target, signal, maxBytes);
					await assertSourceUnchanged(ctx, sourceTarget.target, info, signal);
					return { bytes, version: info.version };
				}
			};
			const format = String(args.format ?? "pdf").toLowerCase();
			const extension = extname(sourceTarget.absolute).toLowerCase().replace(/^\./, "");
			const stem = basename(sourceTarget.display).replace(/\.[^.]+$/u, "");
			const converted = await kit.convertTo({
				extension,
				format,
				source,
				output: { owner: outputOwner(ctx, exec.agent), name: `${stem}.${format}` }
			}, exec.signal);
			let kitPath = converted.path;
			if (typeof ctx.fs.processPathFromHostPath === "function") {
				const mapped = ctx.fs.processPathFromHostPath(converted.path);
				if (mapped === undefined) throw new EngineFailure("OFFICEKIT_REMOTE_WORKSPACE", "converted files are not available when the workspace is on a remote host");
				kitPath = mapped;
			}
			let workspacePath = kitPath;
			if (args.output) {
				const outputTarget = await resolveTarget(ctx, exec, args.output, [`.${format}`], false);
				await runEngine(config, {
					action: "copy",
					source: converted.path,
					target: outputTarget.absolute,
					overwrite: args.overwrite ?? false
				}, exec, outputTarget.root, { OFFICEFORGE_OFFICEKIT_DIR: kit.outputDirectory ?? dirname(converted.path) });
				workspacePath = outputTarget.display;
			}
			return lossless({
				type: format,
				path: workspacePath,
				bytes: converted.bytes,
				pageCount: converted.pageCount,
				missingFonts: converted.missingFonts ?? [],
				officeKitPath: converted.path
			});
		}
	}));
}

// ---------------------------------------------------------------------------
// Plugin entry
// ---------------------------------------------------------------------------

function normalizeConfig(configInput) {
	try {
		return Config(configInput ?? {});
	} catch (error) {
		// A malformed profile override must not abort plugin activation. Defaults
		// keep the plugin usable; the warning is visible in the host log.
		const message = `[dsh-office-forge] invalid config, using defaults: ${error instanceof Error ? error.message : String(error)}`;
		if (typeof console !== "undefined" && typeof console.warn === "function") console.warn(message);
		return Config({});
	}
}

function apply(ctx, configInput) {
	const config = normalizeConfig(configInput);
	ctx.effect(() => {
		const disposers = [];
		const safeRegister = (label, register) => {
			try {
				const disposer = register(ctx, config);
				if (typeof disposer === "function") disposers.push(disposer);
			} catch (error) {
				// One incompatible tool must not abort activation or the profile.
				const message = `[dsh-office-forge] ${label} was not registered: ${error instanceof Error ? error.message : String(error)}`;
				if (typeof ctx.logger?.warn === "function") ctx.logger.warn(message);
				else console.warn(message);
			}
		};
		safeRegister("officeforge_read", registerRead);
		safeRegister("officeforge_create", registerCreate);
		safeRegister("officeforge_edit", registerEdit);
		safeRegister("officeforge_template", registerTemplate);
		safeRegister("officeforge_audit", registerAudit);
		safeRegister("officeforge_doctor", registerDoctor);
		safeRegister("officeforge_convert", registerConvert);
		safeRegister("officeforge_preview", registerPreview);
		safeRegister("officeforge_recalculate", registerRecalculate);
		safeRegister("officeforge_pdf", registerPdf);
		safeRegister("officeforge_native", registerNative);
		if (config.enableBatch) safeRegister("officeforge_batch", registerBatch);
		if (config.enableExport) safeRegister("officeforge_export", registerExport);
		return () => {
			for (const dispose of disposers.reverse()) {
				try { dispose(); } catch { /* a failing disposer must not block the rest */ }
			}
		};
	});
	// The plugin page's "check environment" button reads this read-only route.
	// Keep it optional: headless hosts may not expose a webServer service.
	if (typeof ctx.inject === "function") {
		ctx.inject(["webServer"], (webCtx) => {
			webCtx.effect(() => webCtx.webServer.register({
				kind: "exact",
				path: "/officeforge/doctor",
				handler: async (req, res) => {
					if (req.method !== "GET") {
						res.writeHead(405, { "content-type": "application/json; charset=utf-8" });
						res.end(JSON.stringify({ ok: false, error: { code: "METHOD_NOT_ALLOWED", message: "use GET" } }));
						return;
					}
					let engine;
					let engineError;
					try {
						engine = await runEngine(config, { action: "doctor" }, { signal: undefined }, PLUGIN_ROOT);
					} catch (error) {
						engineError = error;
					}
					const report = buildDoctorReport(engine, engineError, config, getService(ctx, "officeKit"));
					res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
					res.end(JSON.stringify(report));
				}
			}), "dsh-office-forge: /officeforge/doctor route");
		});
	}
}

export { Config, apply, inject, name };
