#!/usr/bin/env node
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "officeforge-ppt-"));
let pass = 0;
const failures = [];

function check(name, ok, detail = "") {
  if (ok) {
    pass += 1;
    console.log(`PASS ${name}`);
  } else {
    failures.push(`${name}${detail ? `: ${detail}` : ""}`);
    console.log(`FAIL ${name}${detail ? `: ${detail}` : ""}`);
  }
}

function run(bridge, payload) {
  const result = spawnSync(process.execPath, [path.join(root, "lib", bridge)], {
    cwd: root,
    input: JSON.stringify(payload),
    encoding: "utf8",
    timeout: 120000,
  });
  let parsed = null;
  try { parsed = JSON.parse(result.stdout || "{}"); } catch { /* noop */ }
  return { result, parsed };
}

const themes = ["swiss","velvet","data","soft","bold","consulting","cobalt-cream","exec-jilan","exec-palace","academic","editorial-forest","cartesian","course-sage","linear-dark","keynote-light","financial-ft","medical-clean","luxury-noir","bcg-emerald","economist-brief","venture-daylight","stripe-craft","aurora-stage","obsidian-amber"];
const masters = ["mckinsey-board","bcg-growth","sequoia-memo","apple-cupertino","stripe-product","economist-macro","nature-defense","imperial-gov"];
const templates = ["weekly","defense","project","pitch","consulting","keynote","gov-report","workshop","architecture-review","financial-earnings","medical-case","postmortem","marketing-campaign","promotion-review","legal-compliance","creative-lookbook","quarterly-okr","product-prd","vc-daylight","industry-research"];

for (const theme of themes) {
  const out = path.join(outDir, `theme-${theme}.pptx`);
  const { result, parsed } = run("office_ppt_studio_bridge.mjs", { spec: { title: `Theme ${theme}`, theme, slides: [{ title: "封面", layout: "cover" }] }, output: out, workspace: root });
  check(`studio theme ${theme}`, result.status === 0 && parsed?.ok === true && fs.existsSync(out), result.stderr || JSON.stringify(parsed));
}

for (const master of masters) {
  const out = path.join(outDir, `master-${master}.pptx`);
  const { result, parsed } = run("office_ppt_studio_bridge.mjs", { spec: { title: `Master ${master}`, master, slides: [{ title: "封面", layout: "cover" }] }, output: out, workspace: root });
  check(`studio master ${master}`, result.status === 0 && parsed?.ok === true && fs.existsSync(out), result.stderr || JSON.stringify(parsed));
}

for (const template of templates) {
  const out = path.join(outDir, `template-${template}.pptx`);
  const { result, parsed } = run("office_ppt_studio_bridge.mjs", { spec: { title: `Template ${template}`, template }, output: out, workspace: root });
  check(`studio template ${template}`, result.status === 0 && parsed?.ok === true && fs.existsSync(out), result.stderr || JSON.stringify(parsed));
}

const layouts = ["3d-stack","ablation-compare","agenda","architecture","before-after","bento","browser","bubble-matrix","bullets","callout","cards","chart","chart-split","chevron","china-map","closing","code","comparison","cover","device","dumbbell","exec-summary","flowchart","flywheel","formula-callout","funnel","gallery","gallery-grid","gantt","geo-map","harvey","heat-matrix","heatmap","image","image-left","image-right","iso-layers","iso-pillars","iso-stack","isometric","issue-action","kpi","kpi-grid","layers","magic-quadrant","map","marimekko","matrix","matrix-2x2","mekko","mermaid","mobile","mockup","org","pillars","porter","process","quadrant","quote","radar","roadmap","sankey","scorecard","section","sequence","sequence-diagram","slope","snippet","spider","statement","steps","swimlane","swot","swot-matrix","table","takeaways","tam","team","timeline","value-chain","waterfall","world-map"];
const chart = { type: "column", categories: ["Q1", "Q2"], series: [{ name: "S", values: [1, 2] }] };
const table = { headers: ["A", "B"], rows: [["1", "2"]] };
const image = { src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nWQAAAAASUVORK5CYII=" };
let layoutOk = 0;
for (const layout of layouts) {
  const attempts = [{}, { chart }, { table }, { image }];
  let ok = false;
  let last = null;
  for (const extra of attempts) {
    const layoutOut = path.join(outDir, `layout-${layout}-${Object.keys(extra).join("-") || "base"}.pptx`);
    const slide = { layout, title: `Layout ${layout}`, bullets: ["示例内容"], ...extra };
    last = run("office_ppt_studio_bridge.mjs", { spec: { title: "Layouts", theme: "consulting", slides: [slide] }, output: layoutOut, workspace: root });
    if (last.result.status === 0 && last.parsed?.ok === true && fs.existsSync(layoutOut)) { ok = true; break; }
  }
  if (ok) layoutOk += 1;
  check(`studio layout ${layout}`, ok, last?.result.stderr || JSON.stringify(last?.parsed));
}
check("studio 82 layouts total", layoutOk === layouts.length, `${layoutOk}/${layouts.length}`);

const zbsOut = path.join(outDir, "zbsph-full.pptx");
const zbs = run("office_ppt_zbsp_bridge.mjs", {
  spec: {
    title: "ZBS Full Schema",
    slides: [{
      title: "全字段", layout: "title", notes: "演讲者备注", background: { type: "solid", color: "#F8FAFC" },
      table: { headers: ["列A", "列B"], rows: [["1", "2"]], header: true },
      chart: { type: "line", categories: ["Q1", "Q2"], series: [{ name: "系列", values: [1, 2] }], legend: true, axes: true },
      shapes: [{ kind: "roundRect", x: 40, y: 300, width: 180, height: 70, fill: { color: "#2563EB", alpha: 20 }, line: { color: "#1F2937", width: 1 } }],
      lines: [{ points: [[40, 420], [400, 420]], arrow: true, line: { color: "#1F2937", width: 2 } }],
      safeArea: { top: 10, bottom: 10 },
    }],
  },
  output: zbsOut, workspace: root,
});
check("zbsph full schema", zbs.result.status === 0 && zbs.parsed?.ok === true && Boolean(zbs.parsed?.pptxPath), zbs.result.stderr || JSON.stringify(zbs.parsed));

for (const template of ["business-blue", "academic-white", "tech-dark", "pitch-bold"]) {
  const vateOut = path.join(outDir, `vate-${template}.pptx`);
  const vate = run("office_ppt_zbsp_bridge.mjs", { spec: { title: `Vate ${template}`, template, slides: [{ title: "封面", layout: "title" }, { title: "第二页", bullets: ["A", "B"] }] }, output: vateOut, workspace: root });
  check(`zbsph Vate ${template}`, vate.result.status === 0 && vate.parsed?.ok === true && Boolean(vate.parsed?.pptxPath), vate.result.stderr || JSON.stringify(vate.parsed));
}

console.log(`PPT BRIDGE TESTS ${failures.length ? "FAILED" : "PASSED"}: ${pass} passed, ${failures.length} failed`);
if (failures.length) {
  console.log(failures.join("\n"));
  process.exitCode = 1;
}
fs.rmSync(outDir, { recursive: true, force: true });
