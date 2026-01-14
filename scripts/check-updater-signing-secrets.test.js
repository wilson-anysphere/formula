import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-updater-signing-secrets.mjs");

/**
 * @param {Record<string, string | undefined>} env
 */
function run(env) {
  const proc = spawnSync(process.execPath, [scriptPath], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      ...env,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("delegates to the CI updater secrets validator (fails when TAURI_PRIVATE_KEY missing)", () => {
  const proc = run({ TAURI_PRIVATE_KEY: undefined, TAURI_KEY_PASSWORD: undefined });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Missing Tauri updater signing secrets/i);
});

test("fails when TAURI_PRIVATE_KEY is an empty string (GitHub Actions missing secret behavior)", () => {
  const proc = run({ TAURI_PRIVATE_KEY: "", TAURI_KEY_PASSWORD: "" });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /\bTAURI_PRIVATE_KEY\b/);
});

test("passes through when TAURI_PRIVATE_KEY is present (unencrypted raw key)", () => {
  const raw = Buffer.alloc(64, 1).toString("base64").replace(/=+$/, ""); // unpadded base64
  const proc = run({ TAURI_PRIVATE_KEY: raw, TAURI_KEY_PASSWORD: "" });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("skips when updater is inactive (plugins.updater.active !== true)", () => {
  const fixturesDir = path.join(repoRoot, "target", "test-fixtures");
  fs.mkdirSync(fixturesDir, { recursive: true });
  const inactiveConfigPath = path.join(fixturesDir, "tauri-updater-inactive.json");

  const originalConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const config = JSON.parse(fs.readFileSync(originalConfigPath, "utf8"));
  config.plugins ??= {};
  config.plugins.updater ??= {};
  config.plugins.updater.active = false;
  fs.writeFileSync(inactiveConfigPath, JSON.stringify(config, null, 2));

  const proc = run({
    FORMULA_TAURI_CONF_PATH: path.relative(repoRoot, inactiveConfigPath),
    TAURI_PRIVATE_KEY: "",
    TAURI_KEY_PASSWORD: "",
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /skipping/i);
});
