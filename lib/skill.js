// Office Forge bundled skill: the PPT intake and routing rules.
//
// This module deliberately mirrors the tiny, dependency-free skill loader used
// by dsh-ppt: read SKILL.md, parse only name/description from the frontmatter,
// and register the body as a runtime skill. A missing or malformed skill must
// never prevent the document tools from activating.

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { join } from "node:path";

export const SKILL_NAMES = [
	"officeforge",
	"officeforge-ppt",
	"guizang-ppt-skill",
	"guizang-page-skill",
	"guizang-social-card-skill",
	"humanizer-zh",
];

export function bundledSkillsDir() {
	return fileURLToPath(new URL("../skills/", import.meta.url));
}

function unquoteYamlScalar(value) {
	const trimmed = value.trim();
	if (
		(trimmed.startsWith('"') && trimmed.endsWith('"')) ||
		(trimmed.startsWith("'") && trimmed.endsWith("'"))
	) {
		return trimmed.slice(1, -1).replace(/\\"/g, '"').replace(/\\n/g, "\n").trim();
	}
	return trimmed;
}

export function parseSkillFile(text) {
	const normalized = text.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
	const match = /^---[ \t]*\n([\s\S]*?)\n---[ \t]*\n?/.exec(normalized);
	if (!match) return { name: "", description: "", content: normalized };
	let name = "";
	let description = "";
	let whenToUse = "";
	for (const rawLine of match[1].split("\n")) {
		const line = rawLine.trim();
		const keyMatch = /^([A-Za-z][A-Za-z0-9_-]*):[ \t]*(.*)$/.exec(line);
		if (keyMatch === null) continue;
		if (keyMatch[1] === "name") name = unquoteYamlScalar(keyMatch[2]);
		else if (keyMatch[1] === "description") description = unquoteYamlScalar(keyMatch[2]);
		else if (keyMatch[1] === "whenToUse") whenToUse = unquoteYamlScalar(keyMatch[2]);
	}
	return { name, description, whenToUse, content: normalized.slice(match[0].length).trimStart() };
}

export function registerOfficeForgeSkill(ctx) {
	if (!ctx?.skills || typeof ctx.skills.register !== "function") return undefined;
	const disposers = [];
	for (const skillName of SKILL_NAMES) {
		const dir = join(bundledSkillsDir(), skillName);
		const text = readFileSync(join(dir, "SKILL.md"), "utf8");
		const parsed = parseSkillFile(text);
		if (!parsed.name || !parsed.description || !parsed.content) {
			throw new Error(`skill ${skillName} has incomplete frontmatter`);
		}
		const dispose = ctx.skills.register({
			name: parsed.name,
			description: parsed.description,
			whenToUse: parsed.whenToUse || undefined,
			content: parsed.content,
			resourceBase: { kind: "directory", path: dir },
			source: "runtime",
		});
		if (typeof dispose === "function") disposers.push(dispose);
	}
	if (disposers.length === 0) return undefined;
	return () => {
		for (const dispose of disposers.reverse()) {
			try { dispose(); } catch { /* a failing skill disposer must not block the rest */ }
		}
	};
}
