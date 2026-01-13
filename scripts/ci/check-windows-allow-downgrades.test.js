import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-windows-allow-downgrades.mjs");

/**
 * @param {any} config
 */
function run(config) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-allow-downgrades-"));
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

test("passes when bundle.windows.allowDowngrades is true", () => {
  const proc = run({ bundle: { windows: { allowDowngrades: true } } });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /allowDowngrades=true/);
});

test("fails when bundle.windows.allowDowngrades is missing", () => {
  const proc = run({ bundle: { windows: {} } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing.*allowDowngrades/);
});

test("fails when bundle.windows.allowDowngrades is false", () => {
  const proc = run({ bundle: { windows: { allowDowngrades: false } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must be true/);
});

test("fails when bundle.windows.allowDowngrades is not a boolean", () => {
  const proc = run({ bundle: { windows: { allowDowngrades: "true" } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Expected boolean/);
});

