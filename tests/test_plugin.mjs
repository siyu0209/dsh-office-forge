// Load the plugin entry in a minimal host mock, register every tool, then
// execute one create + one read through the real Python engine.
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { apply, Config, inject, name } from "../lib/index.js";
import { pythonPath } from "./test_env.mjs";

const PYTHON = pythonPath();

const workspace = mkdtempSync(join(tmpdir(), "officeforge-plugin-"));
const kitRoot = join(workspace, "kit-out");
mkdirSync(kitRoot, { recursive: true });
const registered = new Map();
const warnings = [];
const services = new Map();
const kit = {
  outputDirectory: kitRoot,
  async convertTo(request) {
    const destination = join(kitRoot, request.output.name);
    copyFileSync(join(workspace, "plugin-report.docx"), destination);
    return { path: destination, bytes: statSync(destination).size, pageCount: 1, missingFonts: [] };
  },
  async renderImages() {
    // Deliberately omit page/sheet/range on the second image and pageCount/dpi
    // on the response: real officeKit builds omit whichever field does not apply.
    return { images: [{ png: Buffer.from("fake-png"), width: 1, height: 1, page: 1 }, { png: Buffer.from("fake-png"), width: 2, height: 2 }], missingFonts: [] };
  },
  async recalculate(request) {
    const destination = join(kitRoot, request.output.name);
    copyFileSync(join(workspace, "plugin-audit.xlsx"), destination);
    return { path: destination, bytes: statSync(destination).size, missingFonts: [] };
  },
};
services.set("officeKit", kit);

let routeRegistration;
const ctx = {
  logger: { warn: (message) => warnings.push(String(message)) },
  effect(fn) {
    const dispose = fn();
    return () => { if (typeof dispose === "function") dispose(); };
  },
  inject(names, callback) {
    if (!names.includes("webServer")) return;
    callback({
      effect(fn) {
        const dispose = fn();
        return () => { if (typeof dispose === "function") dispose(); };
      },
      webServer: {
        register(route) {
          routeRegistration = route;
          return () => {};
        },
      },
    });
  },
  tools: {
    register(tool) {
      if (!tool || typeof tool.name !== "string") throw new Error("tool must have a name");
      if (registered.has(tool.name)) throw new Error(`duplicate tool ${tool.name}`);
      if (!tool.output || typeof tool.output.render !== "function") throw new Error(`tool ${tool.name} missing output.render`);
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
      } catch {
        return undefined;
      }
    },
    async readByteRange() {},
    async readBytes(target) { return readFileSync(target.absolute); },
    contains() { return true; },
    processPath(target) { return target.absolute; },
  },
  workingDirectory: { async ensure() { return workspace; } },
  get(key) { return services.get(key); },
};

const exec = {
  agent: { session: { id: "test-session", header: { cwd: workspace } } },
  signal: new AbortController().signal,
};

function check(condition, label) {
  if (!condition) throw new Error(label);
  console.log(`  PASS  ${label}`);
}

try {
  console.log(`[plugin] ${name}`);
  check(inject.join(",") === "tools,fs,workingDirectory", "inject contract is tools/fs/workingDirectory");
  check(typeof Config === "function", "Config is exported");
  apply(ctx, { pythonPath: PYTHON });
  check(warnings.length === 0, `no registration warnings (${warnings.join(" | ")})`);
  check(routeRegistration?.path === "/officeforge/doctor", "doctor HTTP route is registered");
  const routeResponse = {
    statusCode: 0,
    headers: null,
    body: "",
    writeHead(status, headers) { this.statusCode = status; this.headers = headers; },
    end(body) { this.body = body ?? ""; },
  };
  await routeRegistration.handler({ method: "POST" }, routeResponse);
  check(routeResponse.statusCode === 405 && JSON.parse(routeResponse.body).error.code === "METHOD_NOT_ALLOWED", "doctor route rejects non-GET requests");
  const getResponse = {
    statusCode: 0,
    headers: null,
    body: "",
    writeHead(status, headers) { this.statusCode = status; this.headers = headers; },
    end(body) { this.body = body ?? ""; },
  };
  await routeRegistration.handler({ method: "GET" }, getResponse);
  const getBody = JSON.parse(getResponse.body);
  check(getResponse.statusCode === 200 && Array.isArray(getBody.checks) && getBody.checks.length >= 5, `doctor route returns live checks over HTTP (${getResponse.statusCode}: ${getBody.error?.message ?? "ok"})`);
  const expected = ["officeforge_read", "officeforge_create", "officeforge_edit", "officeforge_template", "officeforge_audit", "officeforge_doctor", "officeforge_batch", "officeforge_export", "officeforge_convert", "officeforge_preview", "officeforge_recalculate", "officeforge_pdf", "officeforge_native"];
  for (const toolName of expected) check(registered.has(toolName), `${toolName} registered`);
  check(!registered.has("officeforge_missing"), "no phantom tool registered");
  check(registered.size === expected.length, `exactly ${expected.length} tools registered`);

  const doctor = await registered.get("officeforge_doctor").execute({}, exec);
  check(doctor.ok === true && doctor.plugin?.name === "dsh-office-forge", "doctor reports plugin identity");
  check(doctor.pdfVendor === true, "doctor reports bundled PDF runtime");
  check(Array.isArray(doctor.checks) && doctor.checks.length >= 5, "doctor reports environment checks");
  check(doctor.health === "ok" || doctor.health === "degraded", `doctor reports health ${doctor.health}`);
  check(doctor.checks.some((item) => item.name === "pythonModules" && item.ok === true), "doctor reports Python modules");
  check(doctor.checks.every((item) => typeof item.required === "boolean" && typeof item.status === "string"), "doctor checks carry required/status fields");
  check(doctor.summary && typeof doctor.summary.ok === "number", "doctor reports check summary");
  const ocrCheck = doctor.checks.find((item) => item.name === "ocr");
  check(Boolean(ocrCheck) && typeof ocrCheck.required === "boolean" && (ocrCheck.status === "ok" || ocrCheck.status === "warn"), "doctor reports OCR backend");

  const created = await registered.get("officeforge_create").execute({
    path: "plugin-report.docx",
    type: "docx",
    overwrite: true,
    spec: {
      title: "插件层测试",
      sections: [{ heading: "概览", paragraphs: ["从 Node 层调用 Python 引擎。"] }],
    },
  }, exec);
  check(created.bytes > 1000, `create returned ${created.bytes} bytes`);
  check(created.path === "plugin-report.docx", "create returns workspace-relative path");

  const read = await registered.get("officeforge_read").execute({
    path: "plugin-report.docx",
    mode: "content",
    format: "markdown",
  }, exec);
  check(read.text.includes("插件层测试") && read.text.includes("从 Node 层调用"), "read returns created content");

  const workbookResult = await registered.get("officeforge_create").execute({
    path: "plugin-audit.xlsx",
    type: "xlsx",
    overwrite: true,
    spec: { sheets: [{ name: "Data", columns: ["Item", "Qty", "Amount"], rows: [["A", 10, "=B2*2"], ["B", 20, 30], ["C", 30, "=B4*2"], ["D", 40, "=B5*3"]] }] },
  }, exec);
  check(workbookResult.bytes > 1000, `audit workbook created (${workbookResult.bytes} bytes)`);
  const audit = await registered.get("officeforge_audit").execute({
    path: "plugin-audit.xlsx",
    type: "xlsx",
    mode: "health",
  }, exec);
  check(audit.issueCount >= 1 && audit.summary.hardcoded_value >= 1, "audit returns workbook health findings");
  const themes = await registered.get("officeforge_audit").execute({ type: "pptx", mode: "themes" }, exec);
  check(themes.themes.length >= 24, "audit lists at least 24 PowerPoint themes without a path");

  const bad = await registered.get("officeforge_read").execute({
    path: "../escape.docx",
    mode: "content",
  }, exec).then(() => null, (error) => error);
  check(bad && bad.code === "PATH_OUTSIDE_WORKSPACE", "path escape is rejected");

  const exported = await registered.get("officeforge_export").execute({
    path: "plugin-report.docx",
    output: "plugin-report.pdf",
  }, exec);
  check(exported.bytes > 1000, `export copied ${exported.bytes} bytes`);
  check(exported.path === "plugin-report.pdf", "export returns workspace-relative path");
  check(exported.officeKitPath.startsWith(kitRoot), "export keeps officeKit source path");

  const converted = await registered.get("officeforge_convert").execute({
    path: "plugin-report.docx",
    format: "pdf",
    output: "plugin-convert.pdf",
    overwrite: true,
  }, exec);
  check(converted.path === "plugin-convert.pdf" && converted.backend === "officeKit", "convert bridges through officeKit");

  const preview = await registered.get("officeforge_preview").execute({
    path: "plugin-report.docx",
    outputDir: "previews",
    pages: [1],
  }, exec);
  check(preview.images.length === 2 && preview.images[0].path.startsWith("previews/"), "preview writes PNGs into the workspace");
  check(preview.images.every((image) => image.path.includes("plugin-report-docx-")), "preview names include the source extension to avoid cross-format collisions");
  check(new Set(preview.images.map((image) => image.path)).size === preview.images.length, "preview image names are unique within one call");
  check(!("page" in preview.images[1]) && !("sheet" in preview.images[1]), "preview drops absent optional image fields instead of forwarding undefined");
  check(!("pageCount" in preview) && !("dpi" in preview), "preview drops absent optional response fields instead of forwarding undefined");

  // A rendered image may omit page/sheet/range, and a conversion may omit
  // pageCount. The host rejects `undefined` members as non-lossless JSON, so the
  // tool must drop absent optional fields instead of forwarding them.
  const hostLossless = (value) => {
    if (value === null) return true;
    if (typeof value === "number") return Number.isFinite(value);
    if (typeof value === "string" || typeof value === "boolean") return true;
    if (typeof value !== "object") return false;
    if (Array.isArray(value)) return value.every(hostLossless);
    const proto = Object.getPrototypeOf(value);
    if (proto !== Object.prototype && proto !== null) return false;
    return Object.values(value).every(hostLossless);
  };
  check(hostLossless(preview), "preview output survives the host lossless JSON gate");
  check(hostLossless(converted), "convert output survives the host lossless JSON gate");
  check(hostLossless(await registered.get("officeforge_doctor").execute({}, exec)), "doctor output survives the host lossless JSON gate");

  const recalculated = await registered.get("officeforge_recalculate").execute({
    path: "plugin-audit.xlsx",
    output: "plugin-recalc.xlsx",
    overwrite: true,
  }, exec);
  check(recalculated.path === "plugin-recalc.xlsx" && recalculated.backend === "officeKit", "recalculate refreshes an xlsx through officeKit");
  check(hostLossless(recalculated), "recalculate output survives the host lossless JSON gate");

  const batch = await registered.get("officeforge_batch").execute({
    items: [
      { id: "good", action: "create", path: "batch-good.docx", type: "docx", overwrite: true, spec: { title: "批量成功", sections: [] } },
      { id: "bad", action: "create", path: "../batch-bad.docx", type: "docx", spec: { title: "越界" } },
    ],
  }, exec);
  check(batch.succeeded === 1 && batch.failed === 1, "batch isolates a bad path from a good item");
  check(batch.items.some((item) => item.id === "bad" && item.ok === false), "batch reports the bad item");

  console.log("PLUGIN TESTS PASSED");
} finally {
  try { rmSync(workspace, { recursive: true, force: true }); } catch { /* ignore */ }
}
