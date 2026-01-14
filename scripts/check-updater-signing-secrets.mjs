#!/usr/bin/env node
/**
 * Release preflight: validate that Tauri updater signing secrets are present when the updater is active.
 *
 * This script intentionally lives in `scripts/` because the release workflow references it. The detailed
 * key validation logic lives in `scripts/ci/check-tauri-updater-secrets.mjs` (it understands minisign
 * key formats, encryption detection, and pubkey/private-key mismatch checks).
 */
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
// Test/debug hook: allow overriding the Tauri config path so node:test suites can operate on a temp
// copy instead of mutating the real repo config.
const configPathOverride =
  process.env.FORMULA_TAURI_CONF_PATH || process.env.FORMULA_TAURI_CONFIG_PATH || null;
const configPath = configPathOverride ? path.resolve(repoRoot, configPathOverride) : defaultConfigPath;

/** @type {any} */
let config;
try {
  config = JSON.parse(readFileSync(configPath, "utf8"));
} catch (e) {
  const msg = e instanceof Error ? e.message : String(e);
  console.error("");
  console.error("Updater signing secrets preflight failed");
  console.error(`  - Failed to read/parse ${path.relative(repoRoot, configPath)}.`);
  console.error(`  - Error: ${msg}`);
  console.error("");
  process.exit(1);
}

const updater = config?.plugins?.updater;
const active = updater?.active === true;
if (!active) {
  console.log(
    `Updater signing secrets preflight: updater is not active (plugins.updater.active !== true); skipping validation.`,
  );
  process.exit(0);
}

const scriptPath = path.join(repoRoot, "scripts", "ci", "check-tauri-updater-secrets.mjs");

const proc = spawnSync(process.execPath, [scriptPath], {
  cwd: repoRoot,
  env: process.env,
  stdio: "inherit",
  encoding: "utf8",
});
if (proc.error) {
  console.error(proc.error);
  process.exit(1);
}
process.exit(proc.status ?? 1);
