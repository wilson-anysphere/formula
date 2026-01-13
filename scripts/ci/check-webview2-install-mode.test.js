import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-webview2-install-mode.mjs");

/**
 * @param {any} config
 */
function run(config) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-webview2-mode-"));
  const confPath = path.join(tmpdir, "tauri.conf.json");
  writeFileSync(confPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");

  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: confPath,
    },
  });

  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;
  return proc;
}

test("passes for downloadBootstrapper object mode", () => {
  const proc = run({ bundle: { windows: { webviewInstallMode: { type: "downloadBootstrapper" } } } });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /webviewInstallMode\.type=downloadBootstrapper/);
});

test("passes for downloadBootstrapper string mode", () => {
  const proc = run({ bundle: { windows: { webviewInstallMode: "downloadBootstrapper" } } });
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when webviewInstallMode is missing", () => {
  const proc = run({ bundle: { windows: {} } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing .*webviewInstallMode/);
});

test("fails when webviewInstallMode is skip", () => {
  const proc = run({ bundle: { windows: { webviewInstallMode: "skip" } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must not be 'skip'|set to \"skip\"/i);
});

