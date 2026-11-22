// Compatibility and isolation checks for the DSH host contract.
// Uses the real @deepseek-ai/dsh-tools schema validator through defineTool,
// a mock registry that rejects duplicate/reserved names like the host, and
// no real DSH profile.
import { mkdtempSync, rmSync, readFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { assertSupportedJsonSchema } from "@deepseek-ai/dsh-tools";
import { apply, Config, inject, name } from "../lib/index.js";
import { pythonPath } from "./test_env.mjs";

const PYTHON = pythonPath();

const workspace = mkdtempSync(join(tmpdir(), "officeforge-compat-"));
const registered = new Map();
const warnings = [];
const existing = new Set(["office_render", "office_convert", "office_recalculate", "run_code"]);

function makeCtx(store = registered) {
  return {
    logger: { warn: (message) => warnings.push(String(message)) },
    effect(fn) {
      const dispose = fn();
      return () => { if (typeof dispose === "function") dispose(); };
    },
    tools: {
      register(tool) {
        if (!tool || typeof tool.name !== "string") throw new Error("tool must have a name");
        if (existing.has(tool.name)) throw new Error(`reserved/existing tool ${tool.name}`);
        if (store.has(tool.name)) throw new Error(`tool "${tool.name}" is already registered`);
        assertSupportedJsonSchema(tool.output.schema);
        store.set(tool.name, tool);
        return () => store.delete(tool.name);
      },
    },
    fs: {
      async resolve(raw) { return { targetKey: raw, displayPath: raw, absolute: raw }; },
      async stat() { return { version: "compat-v1", type: "file", size: 1 }; },
      async readByteRange() {},
      async readBytes() { return new Uint8Array([1]); },
      contains() { return true; },
      processPath(target) { return target.targetKey; },
    },
    workingDirectory: { async ensure() { return workspace; } },
    get() { return undefined; },
  };
}

function check(condition, label) {
  if (!condition) throw new Error(label);
  console.log(`  PASS  ${label}`);
}

const expected = ["officeforge_read", "officeforge_create", "officeforge_edit", "officeforge_template", "officeforge_audit", "officeforge_doctor", "officeforge_batch", "officeforge_export", "officeforge_convert", "officeforge_preview", "officeforge_recalculate", "officeforge_pdf", "officeforge_native"];

try {
  check(inject.join(",") === "tools,fs,workingDirectory", "inject contract");
  check(typeof Config === "function" && name === "dsh-office-forge", "entry identity");

  apply(makeCtx(), { pythonPath: PYTHON });
  check(warnings.length === 0, "no warnings on default activation");
  check(registered.size === expected.length, `exactly ${expected.length} tools registered`);
  for (const toolName of expected) check(registered.has(toolName), `registered ${toolName}`);

  // Every tool schema must pass the host validator; duplicate names must fail.
  for (const tool of registered.values()) {
    let ok = true;
    try { assertSupportedJsonSchema(tool.output.schema); } catch { ok = false; }
    check(ok, `schema valid for ${tool.name}`);
  }

  // The plugin must not collide with built-in office tools.
  for (const builtin of ["office_render", "office_convert", "office_recalculate"]) {
    check(!registered.has(builtin), `does not shadow ${builtin}`);
  }
  check(!registered.has("run_code"), "does not touch reserved run_code");

  // Disabling every document module must leave only the always-on doctor tool.
  const emptyRegistered = new Map();
  apply(makeCtx(emptyRegistered), { pythonPath: PYTHON, modules: { word: false, excel: false, ppt: false, pdf: false } });
  check(emptyRegistered.size === 1 && emptyRegistered.has("officeforge_doctor"), "all document modules off keeps only doctor");

  // Disabling Excel must remove xlsx from the read enum.
  const oneRegistered = new Map();
  apply(makeCtx(oneRegistered), { pythonPath: PYTHON, modules: { word: true, excel: false, ppt: false, pdf: false } });
  const read = oneRegistered.get("officeforge_read");
  check(JSON.stringify(read.parameters.properties.type.enum) === JSON.stringify(["docx"]), "type enum follows modules");
  check(oneRegistered.has("officeforge_audit"), "audit available when only Word is enabled");
  check(!oneRegistered.has("officeforge_export"), "export omitted when pdf is disabled");

  // Static isolation: no prompt/LLM/middleware registration in source.
  const source = readFileSync(join(process.cwd(), "lib/index.js"), "utf8");
  const codeOnly = source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
  const forbidden = ["systemPrompt", "system-prompt", "middleware", "llm.stream", "llm/stream", "on(\"llm", "on('llm", "hcot", "attack-scheduler", "advisory"];
  for (const term of forbidden) check(!codeOnly.includes(term), `code avoids ${term}`);

  console.log("COMPATIBILITY TESTS PASSED");
} finally {
  try { rmSync(workspace, { recursive: true, force: true }); } catch { /* ignore */ }
}
