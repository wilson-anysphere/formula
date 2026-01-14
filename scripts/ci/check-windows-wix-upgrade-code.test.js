import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-windows-wix-upgrade-code.mjs");

const expectedUpgradeCode = "a91423b1-a874-5245-a74f-62778e7f1e84";

/**
 * @param {any} config
 */
function run(config) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-wix-upgrade-code-"));
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

test("passes when bundle.windows.wix.upgradeCode is pinned to the expected GUID", () => {
  const proc = run({ bundle: { windows: { wix: { upgradeCode: expectedUpgradeCode } } } });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /windows-wix-upgrade-code: OK/);
});

test("fails when bundle.windows.wix.upgradeCode is missing", () => {
  const proc = run({ bundle: { windows: { wix: {} } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing .*upgradeCode/);
});

test("fails when bundle.windows.wix.upgradeCode is not a UUID", () => {
  const proc = run({ bundle: { windows: { wix: { upgradeCode: "not-a-guid" } } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must be a valid UUID/i);
});

test("fails when bundle.windows.wix.upgradeCode does not match the pinned value", () => {
  const proc = run({ bundle: { windows: { wix: { upgradeCode: "00000000-0000-0000-0000-000000000000" } } } });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must not change/i);
});

