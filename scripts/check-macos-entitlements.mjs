#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const entitlementsPath = path.join(
  repoRoot,
  "apps",
  "desktop",
  "src-tauri",
  "entitlements.plist"
);
const relativeEntitlementsPath = path.relative(repoRoot, entitlementsPath);

/**
 * @param {string} message
 */
function err(message) {
  process.exitCode = 1;
  console.error(message);
}

/**
 * @param {string} heading
 * @param {string[]} details
 */
function errBlock(heading, details) {
  err(`\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}`);
}

/**
 * Best-effort "plist-ish" check: ensure the entitlement exists and is set to `<true/>`.
 *
 * We avoid parsing the XML so the check can run on non-macOS CI runners without
 * depending on `plutil`.
 *
 * @param {string} xml
 * @param {string} key
 * @returns {boolean}
 */
function hasTrueEntitlement(xml, key) {
  const marker = `<key>${key}</key>`;
  const start = xml.indexOf(marker);
  if (start === -1) return false;
  const nextKey = xml.indexOf("<key>", start + marker.length);
  const block = xml.slice(start, nextKey === -1 ? xml.length : nextKey);
  return /<true\s*\/>/.test(block);
}

function main() {
  let xml;
  try {
    xml = readFileSync(entitlementsPath, "utf8");
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock(`macOS entitlements preflight failed`, [
      `Failed to read ${relativeEntitlementsPath}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  // Hardened Runtime + WKWebView (wry) commonly require these two entitlements
  // for JavaScript/WASM execution in signed/notarized builds.
  const required = [
    "com.apple.security.cs.allow-jit",
    "com.apple.security.cs.allow-unsigned-executable-memory",
  ];

  const missing = required.filter((key) => !hasTrueEntitlement(xml, key));
  if (missing.length > 0) {
    errBlock(`Invalid macOS entitlements (${relativeEntitlementsPath})`, [
      `Missing required hardened-runtime JIT entitlements:`,
      ...missing.map((key) => key),
      ``,
      `These are required for WKWebView/JavaScriptCore (including WebAssembly) to run reliably under the hardened runtime.`,
      `A common symptom is a signed/notarized build launching with a blank window.`,
    ]);
  }

  if (process.exitCode) {
    err(`\nmacOS entitlements preflight failed.\n`);
    return;
  }

  console.log(`macOS entitlements preflight passed (${relativeEntitlementsPath}).`);
}

main();

