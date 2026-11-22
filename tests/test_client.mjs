import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const source = readFileSync(join(root, "..", "lib", "client.js"), "utf8");

// The host half of dsh-client-modules reads these three facts from the manifest
// (see resolveMeta/clientExportOf in @deepseek-ai/dsh-client-modules). A package
// that declares dsh.client without a resolvable ./client export throws during
// composition, so the manifest shape is asserted before the bundle runs.
const manifest = JSON.parse(readFileSync(join(root, "..", "package.json"), "utf8"));
assert.equal(manifest.dsh.client.platform, "web", "dsh.client.platform must be web");
assert.equal(manifest.exports["./client"], "./lib/client.js", "exports['./client'] must point at the bundle");
assert.ok(
  readFileSync(join(root, "..", manifest.exports["./client"]), "utf8").length > 0,
  "the declared client bundle must exist"
);
assert.ok(Array.isArray(manifest.dsh.client.external ?? []), "dsh.client.external must be a string array when present");
for (const request of manifest.dsh.client.external ?? []) {
  assert.equal(typeof request, "string");
}

let registration;
globalThis.window = {
  __ModuleLoader__: {
    load(entry) {
      registration = entry;
    }
  }
};

// Execute the bundle exactly like the shell does: the file body registers one factory.
new Function("window", source)(globalThis.window);
assert.ok(registration, "bundle must call window.__ModuleLoader__.load");
assert.equal(registration.id, "dsh-office-forge");
assert.equal(typeof registration.factory, "function");

const fakeReact = {
  createElement(type, props, ...children) {
    return { type, props: props ?? {}, children };
  },
  useState(initial) {
    let value = typeof initial === "function" ? initial() : initial;
    return [value, (next) => { value = typeof next === "function" ? next(value) : next; }];
  }
};
const seed = (name) => {
  if (name === "react") return fakeReact;
  throw new Error(`bundle required unexpected module: ${name}`);
};

const exportsObject = registration.factory(seed);
assert.equal(typeof exportsObject.apply, "function");
assert.deepEqual(exportsObject.inject, ["slots"]);

let slotRegistration;
const ctx = {
  slots: {
    inject(name, callback) {
      assert.equal(name, "plugins.detail.section");
      callback();
    },
    register(options, component) {
      slotRegistration = { options, component };
    }
  }
};
exportsObject.apply(ctx);
assert.ok(slotRegistration, "slots.register must be called");
assert.equal(slotRegistration.options.name, "plugins.detail.section");
assert.equal(slotRegistration.options.id, "officeforge.status");

const tree = slotRegistration.component({
  subject: { kind: "bundle", pkg: { name: "dsh-office-forge", version: "0.6.1" } }
});
assert.equal(tree.type, "section");
assert.equal(tree.props["data-officeforge-status"], "dsh-office-forge");
assert.ok(tree.children.length >= 2, "status card must render at least two paragraphs");
const walk = (node, visit) => {
  if (!node || typeof node !== "object") return;
  visit(node);
  for (const child of node.children ?? []) walk(child, visit);
};
let checkButton;
walk(tree, (node) => {
  if (node.type === "button" && node.props["data-officeforge-check"] !== undefined) checkButton = node;
});
assert.ok(checkButton, "status card must render a check-environment button");
assert.equal(checkButton.props.type, "button");
assert.equal(typeof checkButton.props.onClick, "function");

const fallback = slotRegistration.component({});
assert.equal(fallback.props["data-officeforge-status"], "dsh-office-forge");

console.log("CLIENT UI TESTS PASSED");
