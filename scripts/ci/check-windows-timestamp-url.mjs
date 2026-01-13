#!/usr/bin/env node
/**
 * Release preflight: ensure Windows Authenticode timestamping uses HTTPS.
 *
 * We want to avoid plaintext HTTP timestamping to reduce MITM/proxy risks and
 * improve reliability on locked-down networks.
 *
 * Source of truth:
 *   apps/desktop/src-tauri/tauri.conf.json -> bundle.windows.timestampUrl
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
    die(`windows-timestamp: ERROR Failed to read/parse ${relConfigPath}: ${msg}`);
    return;
  }

  const timestampUrl = config?.bundle?.windows?.timestampUrl;
  if (timestampUrl === undefined || timestampUrl === null) {
    die(
      `windows-timestamp: ERROR Missing ${relConfigPath} -> bundle.windows.timestampUrl.\n` +
        `Set it to an HTTPS timestamp server URL provided/recommended by your signing certificate vendor.\n` +
        `Example (DigiCert): https://timestamp.digicert.com`,
    );
    return;
  }

  if (typeof timestampUrl !== "string" || timestampUrl.trim().length === 0) {
    die(
      `windows-timestamp: ERROR Invalid bundle.windows.timestampUrl in ${relConfigPath}.\n` +
        `Expected a non-empty string, got: ${JSON.stringify(timestampUrl)}`,
    );
    return;
  }

  const normalized = timestampUrl.trim();
  let parsed;
  try {
    parsed = new URL(normalized);
  } catch {
    die(
      `windows-timestamp: ERROR bundle.windows.timestampUrl must be a valid absolute URL (${relConfigPath}).\n` +
        `Got: ${JSON.stringify(normalized)}`,
    );
    return;
  }

  if (parsed.protocol !== "https:") {
    die(
      `windows-timestamp: ERROR bundle.windows.timestampUrl must use HTTPS (${relConfigPath}).\n` +
        `Got: ${JSON.stringify(normalized)}`,
    );
    return;
  }

  console.log(`windows-timestamp: OK bundle.windows.timestampUrl=${normalized}`);
}

main();
