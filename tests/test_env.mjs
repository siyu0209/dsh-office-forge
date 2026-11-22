// Shared test environment discovery. No absolute user paths live in the repo.
//
// Resolution order:
//   1. Explicit env vars set by the developer/CI.
//   2. A DSH_HOME whose sibling resources/app/dsh holds the host checkout.
//   3. Common DSH runtime locations under the current user profile.
//
// Tests that need a live DSH host print a SKIP line and exit 0 when no host is
// available; unit-level tests still run everywhere.

import { existsSync, readdirSync } from "node:fs";
import { join } from "node:path";

export function dshAppRoot() {
  const explicit = process.env.DSH_APP_ROOT;
  if (explicit && existsSync(explicit)) return explicit;

  const candidates = [];
  if (process.env.DSH_HOME) {
    candidates.push(join(process.env.DSH_HOME, "..", "resources", "app", "dsh"));
  }
  if (process.env.LOCALAPPDATA) {
    candidates.push(join(process.env.LOCALAPPDATA, "Programs", "DSH", "resources", "app", "dsh"));
  }
  if (process.env.USERPROFILE) {
    for (const name of safeList(process.env.USERPROFILE)) {
      if (!name.startsWith("DSH")) continue;
      candidates.push(join(process.env.USERPROFILE, name, "app", "resources", "app", "dsh"));
    }
  }
  for (const candidate of candidates) {
    if (existsSync(candidate)) return candidate;
  }
  return null;
}

export function dshCli() {
  const root = dshAppRoot();
  if (!root) return null;
  const path = join(root, "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js");
  return existsSync(path) ? path : null;
}

export function dshClientModulesPath() {
  const root = dshAppRoot();
  if (!root) return null;
  const path = join(root, "node_modules", "@deepseek-ai", "dsh-client-modules", "lib", "client.js");
  return existsSync(path) ? path : null;
}

export function pythonPath() {
  const explicit = process.env.OFFICEFORGE_PYTHON;
  if (explicit && existsSync(explicit)) return explicit;

  const home = process.env.DSH_HOME;
  if (home) {
    const runtime = join(home, "dsh-runtimes");
    if (existsSync(runtime)) {
      for (const name of safeList(runtime)) {
        const candidate = join(runtime, name, "dependencies", "python", "python.exe");
        if (existsSync(candidate)) return candidate;
      }
    }
    const direct = join(home, "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "python.exe");
    if (existsSync(direct)) return direct;
  }
  return process.env.PYTHON || "python";
}

function safeList(path) {
  try {
    return readdirSync(path);
  } catch {
    return [];
  }
}
