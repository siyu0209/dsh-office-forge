#!/usr/bin/env node
/**
 * One-shot DSH profile install test.
 *
 * Uses an isolated DSH_HOME under the OS temp directory, installs the local
 * tarball into a throwaway profile, checks the profile manifest, the composed
 * config, and the browser-client declaration, then removes the entire
 * temporary home. It never writes to the real DSH_HOME, never starts the DSH
 * desktop app, and never touches the user's active profile.
 */
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, existsSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const dshCli = 'C:/Users/zhoutao/DSH 官方 0.2.2/app/resources/app/dsh/node_modules/@deepseek-ai/dsh/lib/bin.js';
if (!existsSync(dshCli)) throw new Error(`DSH CLI not found: ${dshCli}`);

const tarball = resolve(
  import.meta.dirname,
  '..',
  readdirSync(resolve(import.meta.dirname, '..'))
    .filter((name) => /^dsh-office-forge-\d+\.\d+\.\d+\.tgz$/.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
    .pop() ?? 'dsh-office-forge-0.6.1.tgz'
);
if (!existsSync(tarball)) throw new Error(`tarball not found: ${tarball}`);
const expectedVersion = JSON.parse(readFileSync(resolve(import.meta.dirname, '..', 'package.json'), 'utf8')).version;

const home = mkdtempSync(join(tmpdir(), 'officeforge-dsh-'));
const env = { ...process.env, DSH_HOME: home };
const profile = 'officeforge-test';

function run(args) {
  const result = spawnSync(process.execPath, [dshCli, ...args], {
    env,
    encoding: 'utf8',
    timeout: 180000,
    windowsHide: true,
  });
  if (result.status !== 0) {
    throw new Error(`dsh ${args.join(' ')} failed (${result.status})\n${result.stdout}\n${result.stderr}`);
  }
  return `${result.stdout}\n${result.stderr}`;
}

try {
  const install = run(['plugin', '--profile', profile, 'add', tarball]);
  if (!/dsh-office-forge/.test(install)) throw new Error(`install output did not mention package\n${install}`);

  const profileDir = join(home, 'profiles', profile);
  const manifest = JSON.parse(readFileSync(join(profileDir, 'package.json'), 'utf8'));
  const bundles = manifest?.dsh?.profile?.bundles ?? [];
  if (!bundles.includes('dsh-office-forge')) throw new Error(`bundle missing from ${JSON.stringify(bundles)}`);
  if (manifest?.dependencies?.['dsh-office-forge'] === undefined) throw new Error('dependency missing');

  const config = run(['--profile', profile, '--dump-config']);
  if (!config.includes('- id: dsh-office-forge') || !config.includes('name: dsh-office-forge')) {
    throw new Error(`composed config missing officeforge entry\n${config.slice(0, 4000)}`);
  }

  const installedDir = join(profileDir, 'node_modules', 'dsh-office-forge');
  const installed = JSON.parse(readFileSync(join(installedDir, 'package.json'), 'utf8'));
  if (installed.version !== expectedVersion) throw new Error(`unexpected installed version ${installed.version}`);
  if (installed?.dsh?.bundle?.patch !== './cordis.patch.yml') throw new Error('bundle patch missing');

  // The browser half resolves dsh.client through exports["./client"]; a missing
  // file or a non-web platform fails composition after install, not at pack time.
  if (installed?.dsh?.client?.platform !== 'web') throw new Error('dsh.client.platform must be web');
  const clientExport = installed?.exports?.['./client'];
  if (typeof clientExport !== 'string') throw new Error('exports["./client"] must be a string');
  const clientPath = join(installedDir, clientExport.replace(/^\.\//, ''));
  if (!existsSync(clientPath)) throw new Error(`client bundle missing: ${clientPath}`);
  const clientSource = readFileSync(clientPath, 'utf8');
  if (!clientSource.includes('__ModuleLoader__.load')) throw new Error('client bundle must register through __ModuleLoader__.load');

  console.log(`DSH PROFILE TESTS PASSED (${installed.version})`);
} finally {
  rmSync(home, { recursive: true, force: true });
}
