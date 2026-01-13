#!/usr/bin/env node
/**
 * Release preflight: ensure Windows installers allow manual downgrades (rollback via Releases page).
 *
 * Windows MSI installers (WiX) commonly block downgrades unless explicitly configured.
 * Tauri supports this via:
 *   apps/desktop/src-tauri/tauri.conf.json -> bundle.windows.allowDowngrades
 *
 * We enforce this in CI so the "manual downgrade via Releases page" rollback requirement
 * doesn't regress silently.
 */

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
// Test hook: allow overriding the Tauri config path so unit tests can operate on a temp copy
// instead of mutating the real repo config.
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const configPath = process.env.FORMULA_TAURI_CONF_PATH
  ? path.resolve(repoRoot, process.env.FORMULA_TAURI_CONF_PATH)
  : defaultConfigPath;
const relConfigPath = path.relative(repoRoot, configPath);

/**
 * @param {string} message
 */
function die(message) {
  console.error(message);
  process.exitCode = 1;
}

function main() {
  /** @type {any} */
  let config;
  try {
    config = JSON.parse(fs.readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    die(`windows-downgrades: ERROR Failed to read/parse ${relConfigPath}: ${msg}`);
    return;
  }

  const allowDowngrades = config?.bundle?.windows?.allowDowngrades;
  if (allowDowngrades === undefined || allowDowngrades === null) {
    die(
      `windows-downgrades: ERROR Missing ${relConfigPath} -> bundle.windows.allowDowngrades.\n` +
        `This repo requires a manual rollback path on Windows (install an older release from GitHub Releases).\n` +
        `Set it to true:\n` +
        `  "bundle": { "windows": { "allowDowngrades": true } }`,
    );
    return;
  }

  if (typeof allowDowngrades !== "boolean") {
    die(
      `windows-downgrades: ERROR Invalid bundle.windows.allowDowngrades in ${relConfigPath}.\n` +
        `Expected boolean, got: ${JSON.stringify(allowDowngrades)}`,
    );
    return;
  }

  if (allowDowngrades !== true) {
    die(
      `windows-downgrades: ERROR bundle.windows.allowDowngrades must be true (${relConfigPath}).\n` +
        `Got: ${JSON.stringify(allowDowngrades)}\n` +
        `If you intentionally need to block in-place downgrades, update docs/release.md to clearly\n` +
        `document the required uninstall-before-install rollback steps.`,
    );
    return;
  }

  console.log("windows-downgrades: OK bundle.windows.allowDowngrades=true");
}

main();
