// Verify module/feature configuration is normalized and enforced at the Node
// boundary without touching a real DSH profile.
import { mkdtempSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { Config, apply } from "../lib/index.js";
import { pythonPath } from "./test_env.mjs";

const PYTHON = pythonPath();

const workspace = mkdtempSync(join(tmpdir(), "officeforge-config-"));
const registered = new Map();
const warnings = [];
const services = new Map();
const ctx = {
  logger: { warn: (message) => warnings.push(String(message)) },
  effect(fn) {
    const dispose = fn();
    return () => { if (typeof dispose === "function") dispose(); };
  },
  tools: {
    register(tool) {
      if (!tool || typeof tool.name !== "string") throw new Error("tool must have a name");
      registered.set(tool.name, tool);
      return () => registered.delete(tool.name);
    },
  },
  fs: {
    async resolve(raw) {
      const absolute = resolve(workspace, raw);
      return { targetKey: absolute, displayPath: absolute, absolute };
    },
    async stat(target) {
      try {
        const info = statSync(target.absolute);
        return { version: "test-v1", type: info.isFile() ? "file" : "directory", size: info.size };
      } catch { return undefined; }
    },
    async readByteRange() {},
    contains() { return true; },
    processPath(target) { return target.absolute; },
  },
  workingDirectory: { async ensure() { return workspace; } },
  get() { return services.get("officeKit"); },
};
const exec = { agent: { session: { id: "config-session", header: { cwd: workspace } } }, signal: new AbortController().signal };

function check(condition, label) {
  if (!condition) throw new Error(label);
  console.log(`  PASS  ${label}`);
}

try {
  const normalized = Config({ modules: { word: true, excel: false, ppt: false, pdf: false }, features: { word: { toc: false } } });
  check(normalized.modules.excel === false && normalized.modules.word === true, "nested module defaults normalize");
  check(normalized.features.word.toc === false && normalized.features.word.images === true, "nested feature defaults normalize");
  check(normalized.features.word.captions === true && normalized.features.word.styles === true && normalized.features.word.fields === true, "v0.3 Word feature defaults normalize");
  check(normalized.features.excel.formulas === true && normalized.features.excel.dataTools === true, "v0.3 Excel feature defaults normalize");
  check(normalized.features.excel.audit === true && normalized.features.ppt.themes === true && normalized.features.ppt.audit === true, "v0.4 audit/theme feature defaults normalize");
  check(normalized.features.word.audit === true, "v0.8 Word audit feature default normalizes");
  check(normalized.features.word.comments === true && normalized.features.word.revisions === true && normalized.features.word.mailMerge === true, "v0.5 Word review feature defaults normalize");
  check(normalized.features.excel.pivots === true && normalized.features.ppt.animations === true && normalized.features.ppt.smartart === true, "v0.5 pivot/animation/SmartArt feature defaults normalize");

  apply(ctx, {
    pythonPath: PYTHON,
    modules: { word: true, excel: false, ppt: false, pdf: false },
    features: { word: { toc: false } },
  });
  check(warnings.length === 0, `no config registration warnings (${warnings.join(" | ")})`);
  check(registered.has("officeforge_read") && registered.has("officeforge_create"), "word tools registered");
  check(!registered.has("officeforge_export"), "pdf export tool skipped when pdf module is off");
  check(registered.has("officeforge_audit"), "audit tool available for Word-only configuration");
  const readTool = registered.get("officeforge_read");
  check(JSON.stringify(readTool.parameters.properties.type.enum) === JSON.stringify(["docx"]), "read type enum is word-only");

  const failure = await registered.get("officeforge_create").execute({
    path: "blocked-toc.docx",
    type: "docx",
    overwrite: true,
    spec: { toc: true, sections: [{ paragraphs: ["blocked"] }] },
  }, exec).then(() => null, (error) => error);
  check(failure && failure.code === "FEATURE_DISABLED", "disabled Word TOC is refused by engine");

  console.log("PLUGIN CONFIG TESTS PASSED");
} finally {
  try { rmSync(workspace, { recursive: true, force: true }); } catch { /* ignore */ }
}
