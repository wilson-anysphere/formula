#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const defaultRepoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));

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

/**
 * @param {string[]} argv
 * @returns {{ repoRoot: string; entitlementsPath: string }}
 */
function parseArgs(argv) {
  let repoRoot = defaultRepoRoot;
  /** @type {string | undefined} */
  let entitlementsPath;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      console.log(
        [
          "Validate macOS hardened-runtime entitlements required by WKWebView/JavaScriptCore.",
          "",
          "Usage:",
          "  node scripts/check-macos-entitlements.mjs",
          "  node scripts/check-macos-entitlements.mjs --root <repoRoot>",
          "  node scripts/check-macos-entitlements.mjs --path <entitlements.plist>",
          "",
          "Defaults:",
          "  --root defaults to the repository root (derived from this script's location).",
          "  --path defaults to apps/desktop/src-tauri/entitlements.plist under --root.",
          "",
        ].join("\n"),
      );
      process.exit(0);
    }

    if (arg === "--root") {
      const value = argv[i + 1];
      if (!value) {
        errBlock("macOS entitlements preflight failed", [`Missing value for --root.`]);
        return { repoRoot, entitlementsPath: entitlementsPath ?? "" };
      }
      repoRoot = path.resolve(value);
      i += 1;
      continue;
    }

    if (arg === "--path") {
      const value = argv[i + 1];
      if (!value) {
        errBlock("macOS entitlements preflight failed", [`Missing value for --path.`]);
        return { repoRoot, entitlementsPath: entitlementsPath ?? "" };
      }
      entitlementsPath = value;
      i += 1;
      continue;
    }
  }

  const resolvedEntitlementsPath = entitlementsPath
    ? path.isAbsolute(entitlementsPath)
      ? entitlementsPath
      : path.resolve(repoRoot, entitlementsPath)
    : path.join(repoRoot, "apps", "desktop", "src-tauri", "entitlements.plist");

  return { repoRoot, entitlementsPath: resolvedEntitlementsPath };
}

function main() {
  const { repoRoot, entitlementsPath } = parseArgs(process.argv.slice(2));
  if (!entitlementsPath) {
    // Argument parsing already reported an error.
    return;
  }

  const relativeEntitlementsPath = path.relative(repoRoot, entitlementsPath);

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
