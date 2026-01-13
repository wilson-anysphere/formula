#!/usr/bin/env node
/**
 * Release preflight: ensure Windows bundling is configured to install WebView2 when missing.
 *
 * This is a fast guardrail (JSON-only) that fails early before we spend time building installers.
 *
 * The stronger guardrail lives in `scripts/ci/check-windows-webview2-installer.py`, which inspects
 * the produced installers and asserts they contain a WebView2 bootstrapper/runtime reference.
 */

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const relConfigPath = path.relative(repoRoot, configPath);

const allowedTypes = new Set(["downloadBootstrapper", "embedBootstrapper", "offlineInstaller", "fixedRuntime"]);

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
    die(`webview2-config: ERROR Failed to read/parse ${relConfigPath}: ${msg}`);
    return;
  }

  const mode = config?.bundle?.windows?.webviewInstallMode;
  if (mode === undefined || mode === null) {
    die(
      `webview2-config: ERROR Missing ${relConfigPath} -> bundle.windows.webviewInstallMode.\n` +
        `Set it to a non-'skip' option so Windows installers can install WebView2 on clean machines.\n` +
        `Recommended: { "type": "downloadBootstrapper", "silent": true }`,
    );
    return;
  }

  /** @type {string | null} */
  let type = null;
  if (typeof mode === "string") {
    type = mode.trim();
  } else if (typeof mode === "object" && typeof mode.type === "string") {
    type = mode.type.trim();
  } else {
    die(
      `webview2-config: ERROR Invalid bundle.windows.webviewInstallMode in ${relConfigPath}.\n` +
        `Expected a string or an object like { "type": "downloadBootstrapper" }, got: ${JSON.stringify(mode)}`,
    );
    return;
  }

  if (!type) {
    die(`webview2-config: ERROR bundle.windows.webviewInstallMode must not be empty (${relConfigPath}).`);
    return;
  }

  if (type.toLowerCase() === "skip") {
    die(
      `webview2-config: ERROR bundle.windows.webviewInstallMode is set to "skip" (${relConfigPath}).\n` +
        `This produces installers that require users to manually install the WebView2 Runtime.`,
    );
    return;
  }

  if (!allowedTypes.has(type)) {
    die(
      `webview2-config: ERROR Unknown WebView2 install mode type: ${JSON.stringify(type)} (${relConfigPath}).\n` +
        `Expected one of: ${Array.from(allowedTypes)
          .map((t) => JSON.stringify(t))
          .join(", ")}\n` +
        `If Tauri added a new install mode, update scripts/ci/check-webview2-install-mode.mjs.`,
    );
    return;
  }

  console.log(`webview2-config: OK bundle.windows.webviewInstallMode.type=${type}`);
}

main();
