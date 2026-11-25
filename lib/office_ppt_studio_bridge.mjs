#!/usr/bin/env node
/**
 * Office Forge -> dsh-ppt studio bridge.
 *
 * Reads one JSON request from stdin and writes one JSON response to stdout.
 * It only handles the file-level, deliberately small subset of slide specs
 * that map cleanly onto dsh-ppt's deck renderer.  Anything that needs
 * Office-specific objects (comments, media, masters, morph, animations,
 * SmartArt, pictures, tables or charts) is rejected with STUDIO_UNSUPPORTED
 * so the Python backend can fall back to the legacy renderer instead of
 * silently dropping content.
 */
import fs from "node:fs";
import path from "node:path";
import { buildDeckAsync, listMasterPresets } from "dsh-ppt/deck-advanced";

const THEME_MAP = {
  forge: "consulting",
  academic: "academic",
  minimal: "swiss",
  ocean: "cobalt-cream",
  forest: "editorial-forest",
  sunset: "velvet",
  midnight: "linear-dark",
  crimson: "luxury-noir",
  swiss: "swiss",
  velvet: "velvet",
  data: "data",
  soft: "soft",
  bold: "bold",
  consulting: "consulting",
  "cobalt-cream": "cobalt-cream",
  "exec-jilan": "exec-jilan",
  "exec-palace": "exec-palace",
  "editorial-forest": "editorial-forest",
  cartesian: "cartesian",
  "course-sage": "course-sage",
  "linear-dark": "linear-dark",
  "keynote-light": "keynote-light",
  "financial-ft": "financial-ft",
  "medical-clean": "medical-clean",
  "luxury-noir": "luxury-noir",
  "bcg-emerald": "bcg-emerald",
  "economist-brief": "economist-brief",
  "venture-daylight": "venture-daylight",
  "stripe-craft": "stripe-craft",
  "aurora-stage": "aurora-stage",
  "obsidian-amber": "obsidian-amber"
};

const LAYOUT_MAP = {
  title: "cover",
  cover: "cover",
  "title-subtitle": "cover",
  section: "section",
  "section-number": "section",
  closing: "closing",
  content: "bullets",
  "title-content": "bullets",
  "title-only": "bullets",
  bullets: "bullets",
  "bullets-two-column": "bullets",
  "two_content": "bullets",
  "two-content": "bullets",
  "two-col": "bullets",
  "two-column": "bullets",
  "three-column": "bullets",
  "four-quadrant": "bullets",
  kpi: "kpi",
  stats: "stats",
  timeline: "timeline",
  roadmap: "roadmap",
  process: "process",
  "process-chevron": "process",
  comparison: "comparison",
  "before-after": "before-after",
  matrix: "matrix",
  swot: "matrix",
  quote: "quote",
  statement: "statement",
  "image-left": "image-left",
  "image-right": "image-right",
  "full-image": "full-image",
  gallery: "gallery",
  table: "table",
  chart: "chart",
  "chart-split": "chart-split",
  cards: "cards",
  "cards-three": "cards",
  team: "team",
  org: "org",
  funnel: "funnel",
  pyramid: "pyramid",
  cycle: "cycle",
  checklist: "checklist",
  dashboard: "dashboard"
};

// dsh-ppt's own catalog includes layouts we do not need to rename.  Pass those
// through unchanged so the studio renderer can use its full 80+ layout set.
const STUDIO_LAYOUTS = new Set([
  "3d-stack", "ablation-compare", "agenda", "architecture", "before-after", "bento", "browser",
  "bubble-matrix", "bullets", "callout", "cards", "chart", "chart-split", "chevron", "china-map",
  "closing", "code", "comparison", "cover", "device", "dumbbell", "exec-summary", "flowchart",
  "flywheel", "formula-callout", "funnel", "gallery", "gallery-grid", "gantt", "geo-map",
  "harvey", "heat-matrix", "heatmap", "image", "image-left", "image-right", "iso-layers",
  "iso-pillars", "iso-stack", "isometric", "issue-action", "kpi", "kpi-grid", "layers",
  "magic-quadrant", "map", "marimekko", "matrix", "matrix-2x2", "mekko", "mermaid", "mobile",
  "mockup", "org", "pillars", "porter", "process", "quadrant", "quote", "radar", "roadmap",
  "sankey", "scorecard", "section", "sequence", "sequence-diagram", "slope", "snippet",
  "spider", "statement", "steps", "swimlane", "swot", "swot-matrix", "table", "takeaways",
  "tam", "team", "timeline", "value-chain", "waterfall", "world-map"
]);

const MASTER_PRESETS = new Map(
  listMasterPresets("zh").map((preset) => [String(preset.id), preset])
);

const UNSUPPORTED_KEYS = [
  "shapes",
  "shape",
  "transition",
  "animation",
  "smartArt",
  "smartart",
  "comments",
  "comment",
  "media",
  "audio",
  "video",
  "background"
];

const CHART_KIND_ALIASES = {
  bar: "bar",
  column: "column",
  line: "line",
  area: "area",
  pie: "pie",
  doughnut: "doughnut",
  donut: "doughnut",
  scatter: "scatter",
  combo: "combo",
  waterfall: "waterfall"
};

function mediaSource(value) {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return "";
  return String(value.src || value.path || value.file || value.url || "");
}

function mapImage(value) {
  if (!value || typeof value !== "object") return null;
  const src = mediaSource(value);
  if (!src) return null;
  const image = { src };
  if (value.alt !== undefined) image.alt = String(value.alt);
  if (value.fit !== undefined) image.fit = String(value.fit);
  if (value.focalPoint !== undefined) image.focalPoint = value.focalPoint;
  return image;
}

function mapImages(value) {
  return asArray(value).map(mapImage).filter(Boolean);
}

function mapTable(value) {
  if (Array.isArray(value)) return value.map((row) => asArray(row).map((cell) => String(cell)));
  if (!value || typeof value !== "object") return [];
  const rows = asArray(value.rows || value.data).map((row) => asArray(row).map((cell) => String(cell)));
  const header = value.headers || value.header || value.columns;
  if (Array.isArray(header) && rows.length) return [header.map((cell) => String(cell)), ...rows];
  if (Array.isArray(header)) return [header.map((cell) => String(cell))];
  return rows;
}

function mapChart(value) {
  if (!value || typeof value !== "object") return null;
  const rawKind = String(value.kind || value.type || value.chartType || "column").toLowerCase();
  const kind = CHART_KIND_ALIASES[rawKind];
  if (!kind) throw new Error(`STUDIO_UNSUPPORTED:chart-${rawKind}`);
  const categories = value.categories || value.labels || value.x;
  const series = value.series || value.data;
  if (!Array.isArray(categories) || !Array.isArray(series)) {
    throw new Error("STUDIO_UNSUPPORTED:chart-shape");
  }
  return {
    kind,
    categories: categories.map((item) => String(item)),
    series: series.map((item, index) => {
      const source = item && typeof item === "object" ? item : { values: item };
      const values = source.values || source.data || source.y;
      if (!Array.isArray(values)) throw new Error("STUDIO_UNSUPPORTED:chart-series");
      const mapped = {
        name: String(source.name || source.label || `系列 ${index + 1}`),
        values: values.map((v) => (typeof v === "number" ? v : Number(v)))
      };
      if (source.chartKind) mapped.chartKind = String(source.chartKind);
      if (source.secondaryAxis === true) mapped.secondaryAxis = true;
      if (Array.isArray(source.xValues)) mapped.xValues = source.xValues;
      return mapped;
    })
  };
}

function applyBlocks(slide, mapped) {
  const blocks = asArray(slide.blocks);
  if (!blocks.length) return;
  const bullets = [];
  const paragraphs = [];
  const items = [];
  for (const block of blocks) {
    if (!block || typeof block !== "object") continue;
    const type = String(block.type || "").toLowerCase();
    if (type === "title") mapped.title = String(block.text || block.value || mapped.title || "");
    else if (type === "subtitle") mapped.subtitle = String(block.text || block.value || "");
    else if (type === "bullets") {
      const values = asArray(block.items || block.bullets || block.values);
      bullets.push(...values.map(textOf).filter(Boolean));
    } else if (type === "paragraphs" || type === "paragraph" || type === "text") {
      const values = asArray(block.items || block.paragraphs || block.values);
      paragraphs.push(...values.map(textOf).filter(Boolean));
    } else if (type === "kpi" || type === "kpis" || type === "metrics" || type === "stats" || type === "timeline" || type === "matrix") {
      for (const item of asArray(block.items || block.values || block.metrics)) {
        if (!item || typeof item !== "object") continue;
        items.push({
          title: String(item.label || item.title || ""),
          value: String(item.value !== undefined ? item.value : ""),
          text: String(item.delta || item.note || item.text || "")
        });
      }
    } else if (type === "table") {
      const rows = mapTable(block.table || block);
      if (rows.length) {
        mapped.rows = rows;
        mapped.layout = "table";
      }
    } else if (type === "chart") {
      const chart = mapChart(block.chart || block);
      if (chart) {
        mapped.chart = chart;
        mapped.layout = mapped.layout === "chart-split" ? "chart-split" : "chart";
      }
    } else if (type === "image" || type === "images") {
      const image = mapImage(block.image || block);
      if (image) mapped.image = image;
      const images = mapImages(block.images || block.items);
      if (images.length) mapped.images = images;
      if (!mapped.layout || mapped.layout === "content") mapped.layout = "image";
    } else if (type === "quote") {
      mapped.text = String(block.text || block.quote || mapped.text || "");
    } else if (type === "note" || type === "notes") {
      mapped.notes = String(block.text || block.notes || mapped.notes || "");
    }
  }
  if (bullets.length) mapped.bullets = [...(mapped.bullets || []), ...bullets];
  if (paragraphs.length) mapped.bullets = [...(mapped.bullets || []), ...paragraphs];
  if (items.length) mapped.items = [...(mapped.items || []), ...items];
}

function respond(value) {
  process.stdout.write(JSON.stringify(value));
  return value;
}

function fail(code, message, detail) {
  return respond({
    ok: false,
    error: { code, message, ...(detail ? { detail } : {}) }
  });
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function textOf(value) {
  if (typeof value === "string") return value;
  if (value && typeof value === "object" && typeof value.text === "string") return value.text;
  return "";
}

function mapSlide(raw) {
  const slide = raw && typeof raw === "object" ? raw : {};
  for (const key of UNSUPPORTED_KEYS) {
    if (slide[key] !== undefined && slide[key] !== null && slide[key] !== "") {
      throw new Error(`STUDIO_UNSUPPORTED:${key}`);
    }
  }
  const rawLayout = String(slide.layout || "content");
  const layout = STUDIO_LAYOUTS.has(rawLayout) ? rawLayout : (LAYOUT_MAP[rawLayout] || "bullets");
  const bullets = asArray(slide.bullets)
    .map(textOf)
    .filter(Boolean);
  const paragraphs = asArray(slide.paragraphs)
    .map(textOf)
    .filter(Boolean);
  const body = bullets.length ? bullets : paragraphs;
  const mapped = {
    layout,
    title: String(slide.title || ""),
    subtitle: String(slide.subtitle || ""),
    kicker: String(slide.kicker || "")
  };
  if (["bullets-two-column", "two_content", "two-content", "two-col", "two-column"].includes(rawLayout)) {
    mapped.variant = "two-column";
  }
  if (body.length) mapped.bullets = body;
  const kpis = asArray(slide.kpis || slide.metrics).map((item) => ({
    title: String(item && (item.label || item.title) || ""),
    value: String(item && item.value !== undefined ? item.value : ""),
    text: String(item && (item.delta || item.note || item.text) || "")
  }));
  const items = asArray(slide.items || slide.cards || slide.steps || slide.columns);
  if (kpis.length || items.length) mapped.items = [...kpis, ...items];
  if (slide.quote && !mapped.text) mapped.text = String(slide.quote);
  if (slide.notes) mapped.notes = String(slide.notes);

  if (slide.table !== undefined && slide.table !== null) {
    const rows = mapTable(slide.table);
    if (rows.length) {
      mapped.rows = rows;
      mapped.layout = "table";
    }
  }
  if (slide.chart !== undefined && slide.chart !== null) {
    const chart = mapChart(slide.chart);
    if (chart) {
      mapped.chart = chart;
      if (mapped.layout !== "chart-split") mapped.layout = "chart";
    }
  }
  if (slide.image !== undefined && slide.image !== null) {
    const image = mapImage(slide.image);
    if (image) mapped.image = image;
  }
  if (slide.images !== undefined && slide.images !== null) {
    const images = mapImages(slide.images);
    if (images.length) mapped.images = images;
  }
  applyBlocks(slide, mapped);
  return mapped;
}

function main() {
  let payload;
  try {
    payload = JSON.parse(fs.readFileSync(0, "utf8"));
  } catch (error) {
    return fail("INVALID_PAYLOAD", `invalid JSON request: ${error.message}`);
  }

  const spec = payload && typeof payload.spec === "object" && payload.spec ? payload.spec : {};
  const slides = asArray(spec.slides);
  const template = String(spec.template || "").trim();
  if (!slides.length && !template) {
    return fail("INVALID_ARGUMENT", "studio backend requires at least one slide or a template");
  }

  let mapped;
  try {
    mapped = slides.map(mapSlide);
  } catch (error) {
    const message = String(error && error.message || error);
    if (message.startsWith("STUDIO_UNSUPPORTED:")) {
      const feature = message.split(":")[1];
      return fail("STUDIO_UNSUPPORTED", `studio backend cannot render ${feature} slides; use pptxEngine=legacy`, { feature });
    }
    return fail("INVALID_ARGUMENT", message);
  }

  const outputPath = String(payload.output || "");
  if (!outputPath) return fail("INVALID_ARGUMENT", "output path is required");
  const outputDir = path.dirname(outputPath);
  const fileName = path.basename(outputPath, path.extname(outputPath));
  const rawMaster = spec.master ?? spec.masterPreset;
  if (rawMaster !== undefined && rawMaster !== null && typeof rawMaster !== "string") {
    // The legacy renderer also accepts a raw master object (footer/background).
    // Let it own that shape instead of pretending the studio bridge understands it.
    return fail("STUDIO_UNSUPPORTED", "studio backend cannot render a raw master object; use pptxEngine=legacy", { feature: "master" });
  }
  const masterId = typeof rawMaster === "string" ? rawMaster.trim() : "";
  const master = masterId ? MASTER_PRESETS.get(masterId) : null;
  if (masterId && !master) {
    return fail("MASTER_NOT_FOUND", `unknown master preset: ${masterId}`, {
      available: [...MASTER_PRESETS.keys()]
    });
  }
  const theme = master?.theme
    || THEME_MAP[String(spec.theme || "forge")]
    || String(spec.theme || "consulting");

  buildDeckAsync({
    title: String(spec.title || mapped[0]?.title || "Office Forge Presentation"),
    ...(mapped.length ? { slides: mapped } : {}),
    ...(template ? { template } : {}),
    ...(spec.brand || master?.brand ? { brand: { ...(master?.brand || {}), ...(spec.brand || {}) } } : {}),
    theme,
    outputDir,
    fileName,
    overwrite: true,
    cwd: String(payload.workspace || process.cwd()),
    writeHandout: false,
    motion: spec.motion !== false
  }).then((result) => {
    process.stdout.write(JSON.stringify({
      ok: true,
      pptxPath: result.pptxPath,
      htmlPath: result.htmlPath,
      jsonPath: result.jsonPath,
      slides: mapped.length,
      theme: result.theme || theme,
      engine: "studio"
    }));
  }).catch((error) => {
    fail("STUDIO_ERROR", String(error && error.message || error));
  });
}

main();
