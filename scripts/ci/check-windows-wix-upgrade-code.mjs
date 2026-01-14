#!/usr/bin/env node
/**
 * Release preflight: ensure Windows MSI upgrades/downgrades remain stable by pinning WiX upgradeCode.
 *
 * On Windows, Formula's auto-updater installs the `.msi` bundle (see docs/desktop-updater-target-mapping.md).
 * MSI major upgrades/downgrades rely on a stable `UpgradeCode` GUID. If it changes between versions,
 * users can end up with side-by-side installs and downgrades can be blocked.
 *
 * Source of truth:
 *   apps/desktop/src-tauri/tauri.conf.json -> bundle.windows.wix.upgradeCode
 *
 * This repo intentionally pins a specific GUID that matches the historical Tauri default for Formula,
 * and CI enforces that it never changes after a release is shipped.
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

// Keep in sync with apps/desktop/src-tauri/tauri.conf.json and guardrail tests.
const expectedUpgradeCode = "a91423b1-a874-5245-a74f-62778e7f1e84";

const uuidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

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
    die(`windows-wix-upgrade-code: ERROR Failed to read/parse ${relConfigPath}: ${msg}`);
    return;
  }

  const upgradeCode = config?.bundle?.windows?.wix?.upgradeCode;
  if (upgradeCode === undefined || upgradeCode === null) {
    die(
      `windows-wix-upgrade-code: ERROR Missing ${relConfigPath} -> bundle.windows.wix.upgradeCode.\n` +
        `Formula requires a stable WiX upgrade code so MSI upgrades/downgrades keep working.\n` +
        `Set it to the pinned value used by shipped releases:\n` +
        `  "bundle": { "windows": { "wix": { "upgradeCode": "${expectedUpgradeCode}" } } }`,
    );
    return;
  }

  if (typeof upgradeCode !== "string" || upgradeCode.trim().length === 0) {
    die(
      `windows-wix-upgrade-code: ERROR Invalid bundle.windows.wix.upgradeCode in ${relConfigPath}.\n` +
        `Expected a non-empty string GUID, got: ${JSON.stringify(upgradeCode)}`,
    );
    return;
  }

  const normalized = upgradeCode.trim();
  if (!uuidRe.test(normalized)) {
    die(
      `windows-wix-upgrade-code: ERROR bundle.windows.wix.upgradeCode must be a valid UUID (${relConfigPath}).\n` +
        `Got: ${JSON.stringify(normalized)}`,
    );
    return;
  }

  if (normalized.toLowerCase() !== expectedUpgradeCode) {
    die(
      `windows-wix-upgrade-code: ERROR bundle.windows.wix.upgradeCode must not change after shipping (${relConfigPath}).\n` +
        `Expected: ${expectedUpgradeCode}\n` +
        `Got:      ${normalized}\n` +
        `Changing the upgrade code breaks MSI upgrade/downgrade behavior and can create side-by-side installs.`,
    );
    return;
  }

  console.log(`windows-wix-upgrade-code: OK bundle.windows.wix.upgradeCode=${normalized}`);
}

main();

