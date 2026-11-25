#!/usr/bin/env node
/**
 * Office Forge -> dsh-ppt-studio (zbsph) bridge.
 *
 * Generates a small PPTD project (deck.yaml + pages/*.yaml) from Office Forge's
 * flat slide spec, then uses dsh-ppt-studio's own resolveDeck/exportPptx
 * functions.  This is intentionally a text-first adapter: table/chart/image/
 * media/comments/master/morph/animation/SmartArt remain on the legacy
 * python-pptx path so nothing is silently dropped.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value === undefined || value === null) return [];
  return [value];
}

function textOf(value) {
  if (typeof value === "string") return value;
  if (value === undefined || value === null) return "";
  return String(value);
}

function q(value) {
  return JSON.stringify(textOf(value));
}

function safeName(value) {
  return String(value).replace(/[^\p{L}\p{N}._-]+/gu, "_").replace(/^_+|_+$/gu, "") || "image";
}

function tableElement(value) {
  if (!value) return null;
  if (Array.isArray(value)) {
    const rows = value.map((row) => asArray(row).map((cell) => String(cell)));
    if (!rows.length) return null;
    const cols = rows.shift();
    return { elementId: "table", elementType: "table", cols, rows, header: true };
  }
  if (typeof value !== "object") return null;
  const rows = asArray(value.rows || value.data).map((row) => asArray(row).map((cell) => String(cell)));
  let cols = value.headers || value.header || value.columns;
  let header = value.header;
  if (!Array.isArray(cols) && rows.length) {
    cols = rows.shift();
    if (header === undefined) header = true;
  }
  if (!Array.isArray(cols)) return null;
  return {
    elementId: String(value.elementId || "table"),
    elementType: "table",
    cols: cols.map((cell) => String(cell)),
    rows,
    header: header === undefined ? false : Boolean(header)
  };
}

function chartElement(value) {
  if (!value || typeof value !== "object") return null;
  const raw = value.chart && typeof value.chart === "object" ? value.chart : value;
  const rawType = String(raw.type || raw.kind || raw.chartType || "bar").toLowerCase();
  const type = ({ bar: "bar", column: "bar", column2d: "bar", line: "line", pie: "pie" })[rawType];
  if (!type) return null;
  let data;
  if (raw.data && Array.isArray(raw.data.cols) && Array.isArray(raw.data.rows)) {
    data = { cols: raw.data.cols, rows: raw.data.rows };
  } else {
    const categories = raw.categories || raw.labels || raw.x;
    const series = raw.series || raw.data;
    if (!Array.isArray(categories) || !Array.isArray(series)) return null;
    const names = series.map((item, index) => String(item?.name || item?.label || `Series ${index + 1}`));
    const rows = categories.map((category, index) => [
      String(category),
      ...series.map((item) => {
        const values = item?.values || item?.data || item?.y || [];
        const itemValue = Array.isArray(values) ? values[index] : 0;
        return typeof itemValue === "number" ? itemValue : Number(itemValue || 0);
      })
    ]);
    data = { cols: ["Category", ...names], rows };
  }
  const chart = { type, data };
  for (const key of ["series", "colors", "render", "legend", "labels", "axes"]) {
    if (raw[key] !== undefined) chart[key] = raw[key];
  }
  return { elementId: String(raw.elementId || "chart"), elementType: "chart", chart };
}

function copyBackground(value, deckDir, workspace) {
  const src = String(value?.src || value?.path || value?.file || "");
  if (!src) return null;
  const absolute = path.isAbsolute(src) ? src : path.resolve(workspace, src);
  if (!fs.existsSync(absolute)) return null;
  const mediaDir = path.join(deckDir, "media");
  fs.mkdirSync(mediaDir, { recursive: true });
  const name = safeName(path.basename(absolute));
  fs.copyFileSync(absolute, path.join(mediaDir, name));
  return { ...value, src: `media/${name}` };
}

function copyImage(value, deckDir, workspace, index) {
  const src = typeof value === "string" ? value : String(value?.src || value?.path || value?.file || "");
  if (!src) return null;
  const absolute = path.isAbsolute(src) ? src : path.resolve(workspace, src);
  if (!fs.existsSync(absolute)) return null;
  const mediaDir = path.join(deckDir, "media");
  fs.mkdirSync(mediaDir, { recursive: true });
  const base = safeName(path.basename(absolute));
  let name = base;
  let suffix = 2;
  while (fs.existsSync(path.join(mediaDir, name))) {
    const ext = path.extname(base);
    name = `${path.basename(base, ext)}-${suffix++}${ext}`;
  }
  fs.copyFileSync(absolute, path.join(mediaDir, name));
  return { elementId: `image-${index + 1}`, elementType: "image", src: `media/${name}`, fit: String(value?.fit || "cover") };
}

const SHAPE_KINDS = new Set([
  "rect", "roundRect", "ellipse", "triangle", "custGeom",
  "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow",
  "pentagon", "hexagon", "chevron", "parallelogram", "diamond", "octagon", "star5",
  "notchedRightArrow", "flowchartProcess", "flowchartDecision", "flowchartData", "flowchartTerminator"
]);

function boundsOf(value, fallback) {
  if (Array.isArray(value.bounds) && value.bounds.length === 4) return value.bounds.map(Number);
  if ([value.x, value.y, value.width ?? value.w, value.height ?? value.h].every((v) => v !== undefined)) {
    return [Number(value.x), Number(value.y), Number(value.width ?? value.w), Number(value.height ?? value.h)];
  }
  return fallback;
}

function shapeElement(value, index) {
  if (!value || typeof value !== "object") return null;
  const kind = String(value.kind || value.type || "rect");
  if (!SHAPE_KINDS.has(kind)) return null;
  const out = {
    elementId: String(value.elementId || `shape-${index + 1}`),
    elementType: "shape",
    bounds: boundsOf(value, [60, 150, 240, 120]),
    kind
  };
  for (const key of ["fill", "line", "rotation", "path", "role", "contains", "badgeOf", "roleReason"]) {
    if (value[key] !== undefined) out[key] = value[key];
  }
  return out;
}

function lineElement(value, index) {
  if (!value || typeof value !== "object") return null;
  const points = value.points || value.line?.points;
  if (!Array.isArray(points) || points.length < 2) return null;
  const out = {
    elementId: String(value.elementId || `line-${index + 1}`),
    elementType: "line",
    points: points.map((point) => Array.isArray(point) ? [Number(point[0]), Number(point[1])] : point)
  };
  for (const key of ["arrow", "line", "dash", "attach", "role", "contains", "badgeOf", "roleReason"]) {
    if (value[key] !== undefined) out[key] = value[key];
  }
  return out;
}

function yamlElement(element) {
  // JSON is a strict subset of YAML 1.2 flow syntax.  Emitting JSON values keeps
  // every valid PPTD field (nested content/fill/chart/path/role...) lossless
  // instead of hand-writing a partial subset of the schema.
  const lines = [`  - elementId: ${q(element.elementId)}`, `    elementType: ${String(element.elementType)}`];
  for (const [key, value] of Object.entries(element)) {
    if (key === "elementId" || key === "elementType" || value === undefined) continue;
    lines.push(`    ${key}: ${JSON.stringify(value)}`);
  }
  return lines;
}

function hasUnsupported(slide) {
  const keys = ["media", "audio", "video",
    "comments", "comment", "master", "transition",
    "animation", "smartArt", "smartart", "morph"];
  for (const key of keys) {
    const value = slide[key];
    if (value !== undefined && value !== null && !(Array.isArray(value) && value.length === 0)) return key;
  }
  return null;
}

function makePage(slide, index, deckDir, workspace) {
  const title = textOf(slide.title);
  const subtitle = textOf(slide.subtitle);
  const bullets = asArray(slide.bullets).map(textOf).filter(Boolean);
  const paragraphs = asArray(slide.paragraphs).map(textOf).filter(Boolean);
  const layout = String(slide.layout || "").toLowerCase();
  const isCover = layout === "title" || layout === "cover" || layout === "title-subtitle";
  const elements = [];
  if (title) {
    elements.push({
      elementId: "title",
      elementType: "text",
      bounds: [60, isCover ? 170 : 40, 840, 80],
      content: { text: title, style: "$title" }
    });
  }
  if (subtitle) {
    elements.push({
      elementId: "subtitle",
      elementType: "text",
      bounds: [80, 260, 800, 60],
      content: { text: subtitle, style: "$body" }
    });
  }
  const bodyLines = [];
  if (bullets.length) bodyLines.push(...bullets.map((line) => `• ${line}`));
  if (paragraphs.length) bodyLines.push(...paragraphs);
  if (bodyLines.length) {
    elements.push({
      elementId: "body",
      elementType: "text",
      bounds: [80, isCover ? 330 : 150, 800, isCover ? 100 : 330],
      content: { text: bodyLines.join("\n"), style: "$body" }
    });
  }
  const table = tableElement(slide.table);
  if (table) elements.push({ ...table, bounds: [60, 150, 840, 320] });
  const chart = chartElement(slide.chart);
  if (chart) elements.push({ ...chart, bounds: [60, 150, 840, 340] });
  asArray(slide.images || slide.image).forEach((value, imageIndex) => {
    const image = copyImage(value, deckDir, workspace, imageIndex);
    if (image) elements.push({ ...image, bounds: [520, 140 + imageIndex * 12, 360, 260] });
  });
  asArray(slide.shapes || slide.shape).forEach((value, shapeIndex) => {
    const shape = shapeElement(value, shapeIndex);
    if (shape) elements.push(shape);
  });
  asArray(slide.lines || slide.line).forEach((value, lineIndex) => {
    const line = lineElement(value, lineIndex);
    if (line) elements.push(line);
  });
  if (!elements.length) {
    elements.push({
      elementId: "empty",
      elementType: "text",
      bounds: [80, 200, 800, 80],
      content: { text: "", style: "$body" }
    });
  }
  let rawBackground = slide.background ?? slide.backgrounds ?? "#FFFFFF";
  if (rawBackground && typeof rawBackground === "object" && rawBackground.type === "image" && rawBackground.src) {
    const copied = copyBackground(rawBackground, deckDir, workspace);
    if (copied) rawBackground = copied;
  }
  const lines = [
    `pageType: ${isCover ? "cover" : "content"}`,
    `background: ${JSON.stringify(rawBackground)}`
  ];
  for (const key of ["safeArea", "expectedOverlaps", "expectedOutOfSafeArea", "overlapMode"]) {
    if (slide[key] !== undefined) lines.push(`${key}: ${JSON.stringify(slide[key])}`);
  }
  lines.push(`elements:`);
  for (const element of elements) {
    lines.push(...yamlElement(element));
  }
  const notes = textOf(slide.notes);
  if (notes) {
    lines.push(`notes: ${q(notes)}`);
  }
  return lines.join("\n") + "\n";
}

async function main() {
  const input = await new Promise((resolve, reject) => {
    let data = "";
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (chunk) => { data += chunk; });
    process.stdin.on("end", () => resolve(data));
    process.stdin.on("error", reject);
  });
  const request = JSON.parse(input);
  const spec = request.spec || {};
  const output = request.output;
  const workspace = request.workspace || process.cwd();
  if (!output) throw Object.assign(new Error("output is required"), { code: "ZBSPH_ERROR" });
  const slides = asArray(spec.slides);
  if (!slides.length) throw Object.assign(new Error("deck must contain at least one slide"), { code: "ZBSPH_ERROR" });
  const unsupported = slides.map(hasUnsupported).find(Boolean);
  if (unsupported) {
    throw Object.assign(new Error(`zbsp adapter does not support ${unsupported}`), { code: "ZBSPH_UNSUPPORTED" });
  }

  const deckDir = fs.mkdtempSync(path.join(workspace, ".officeforge-zbsp-"));
  try {
    const pagesDir = path.join(deckDir, "pages");
    fs.mkdirSync(pagesDir, { recursive: true });
    const pageRefs = [];
    slides.forEach((slide, index) => {
      const pageName = `pages/${String(index + 1).padStart(2, "0")}.yaml`;
      fs.writeFileSync(path.join(deckDir, pageName), makePage(slide, index, deckDir, workspace), "utf8");
      pageRefs.push(pageName);
    });
    const title = textOf(spec.title || spec.deckTitle || "Office Forge deck");
    const studioPkg = require.resolve("dsh-ppt-studio/package.json");
    const studioRoot = path.dirname(studioPkg);
    const templateId = String(spec.template || "").trim();
    let templateTheme = null;
    if (templateId) {
      try {
        const yaml = createRequire(path.join(studioRoot, "package.json"))("yaml");
        const templatePath = path.join(studioRoot, "templates", templateId, "deck.yaml");
        if (fs.existsSync(templatePath)) {
          const parsed = yaml.parse(fs.readFileSync(templatePath, "utf8"));
          if (parsed && parsed.theme && typeof parsed.theme === "object") templateTheme = parsed.theme;
        }
      } catch {
        // Template lookup is best-effort; unknown ids are reported by the caller's catalog.
      }
    }
    const deckTheme = (spec.deckTheme && typeof spec.deckTheme === "object")
      ? spec.deckTheme
      : (templateTheme || ((spec.theme && typeof spec.theme === "object") ? spec.theme : null));
    const themeLines = deckTheme
      ? [`theme: ${JSON.stringify(deckTheme)}`]
      : [
          "theme:",
          "  colors:",
          '    ink: "#1F2937"',
          '    primary: "#2563EB"',
          '    muted: "#6B7280"',
          "  textStyles:",
          "    title:",
          "      fontSize: 32",
          "      color: \"$ink\"",
          "      bold: true",
          "    body:",
          "      fontSize: 16",
          "      color: \"$ink\""
        ];
    const deckLines = [
      "version: 1",
      `title: ${q(title)}`,
      "size: [960, 540]",
      ...themeLines,
      "pages:",
      ...pageRefs.map((ref) => `  - ${q(ref)}`),
      ""
    ];
    fs.writeFileSync(path.join(deckDir, "deck.yaml"), deckLines.join("\n"), "utf8");

    const schemaUrl = pathToFileURL(path.join(studioRoot, "lib", "pptd", "schema.js")).href;
    const exportUrl = pathToFileURL(path.join(studioRoot, "lib", "pptd", "export-pptx.js")).href;
    const { resolveDeck } = await import(schemaUrl);
    const { exportPptx } = await import(exportUrl);
    const ctx = await resolveDeck(deckDir);
    const result = await exportPptx(ctx, { out: output, engine: "pptd" });
    process.stdout.write(JSON.stringify({
      ok: true,
      pptxPath: output,
      slides: result.slides ?? slides.length,
      engine: "zbsph"
    }) + "\n");
  } finally {
    fs.rmSync(deckDir, { recursive: true, force: true });
  }
}

main().catch((error) => {
  process.stdout.write(JSON.stringify({
    ok: false,
    error: { code: error.code || "ZBSPH_ERROR", message: String(error.message || error) }
  }) + "\n");
  process.exitCode = 1;
});
