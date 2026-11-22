#!/usr/bin/env node
/**
 * Host-machinery test for the browser half.
 *
 * Loads the REAL `@deepseek-ai/dsh-client-modules` client bundle from the
 * installed DSH (it is the exact code the shell runs), builds a
 * ClientModuleSystem over a one-row boot manifest, and lets it transport,
 * register and materialize our own `lib/client.js`. This proves the bundle
 * shape (`window.__ModuleLoader__.load` + factory returning `apply`/`inject`)
 * is what the host consumes, without starting DSH or touching any profile.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const packageRoot = join(root, '..');
const hostClientModules = 'C:/Users/zhoutao/DSH 官方 0.2.2/app/resources/app/dsh/node_modules/@deepseek-ai/dsh-client-modules/lib/client.js';

const queued = [];
globalThis.window = {
  __ModuleLoader__: {
    mode: 'queue',
    pendingQueue: queued,
    load(registration) {
      queued.push(registration);
    }
  }
};

const hostModule = await import(pathToFileURL(hostClientModules).href);
const bootstrapRegistration = queued.find((entry) => entry.id === '@deepseek-ai/dsh-client-modules');
assert.ok(bootstrapRegistration, 'host client bundle must register itself through __ModuleLoader__');
const bootstrapExports = bootstrapRegistration.factory(() => {
  throw new Error('the client-modules bundle must not require anything');
});
assert.equal(typeof bootstrapExports.createClientModuleSystem, 'function');

// The shell seeds the bootstrap module out of band, so the queue it drains must
// no longer hold that same registration (otherwise the system rejects it as a
// duplicate factory). Mirror that here before constructing the system.
globalThis.window.__ModuleLoader__.pendingQueue = queued.filter(
  (entry) => entry.id !== '@deepseek-ai/dsh-client-modules'
);

const bundleSource = readFileSync(join(packageRoot, 'lib', 'client.js'), 'utf8');
const bundleId = 'dsh-office-forge';
const bundleUrl = `http://localhost/plugins/${bundleId}/client.js?rev=deadbeef0001`;

const fakeReact = {
  createElement(type, props, ...children) {
    return { type, props: props ?? {}, children };
  },
  useState(initial) {
    let value = typeof initial === "function" ? initial() : initial;
    return [value, (next) => { value = typeof next === "function" ? next(value) : next; }];
  }
};

const system = bootstrapExports.createClientModuleSystem(
  globalThis.window.__ModuleLoader__,
  { id: '@deepseek-ai/dsh-client-modules', exports: bootstrapExports },
  {
    boot: {
      rev: 'rev-1',
      entries: [
        {
          id: bundleId,
          url: bundleUrl,
          rev: 'deadbeef0001',
          inject: ['slots'],
          external: []
        }
      ],
      batches: [{ phase: 'application', url: 'http://localhost/plugins/batch.js?rev=deadbeef0001', rev: 'deadbeef0001', entries: [bundleId] }]
    },
    staticModules: { react: fakeReact },
    loadBundle: async (url) => {
      assert.equal(url, bundleUrl, 'the system must transport the declared bundle URL');
      // Execute the bundle exactly like a <script> element would.
      new Function('window', bundleSource)(globalThis.window);
    }
  }
);

const exported = await system.import(bundleId);
assert.equal(typeof exported.apply, 'function', 'materialized bundle must expose apply');
assert.deepEqual(exported.inject, ['slots']);

let registration;
const ctx = {
  slots: {
    inject(name, callback) {
      assert.equal(name, 'plugins.detail.section');
      callback();
    },
    register(options, component) {
      registration = { options, component };
    }
  }
};
exported.apply(ctx);
assert.ok(registration, 'apply must register the detail section');
assert.equal(registration.options.name, 'plugins.detail.section');

const card = registration.component({
  subject: { kind: 'bundle', pkg: { name: bundleId, version: '0.7.0' } }
});
assert.equal(card.type, 'section');
assert.equal(card.props['data-officeforge-status'], bundleId);

// A second import is memoized by the host, so the bundle must not re-register.
const again = await system.import(bundleId);
assert.equal(again, exported, 'the host must memoize the materialized module');

console.log('CLIENT HOST MACHINERY TESTS PASSED');
