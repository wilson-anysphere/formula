#!/usr/bin/env node
import { spawnSync } from "node:child_process";
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
  let entitlementsPathOverride;

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
          "  --path defaults to the file referenced by bundle.macOS.entitlements in apps/desktop/src-tauri/tauri.conf.json.",
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
        return { repoRoot, entitlementsPathOverride: entitlementsPathOverride ?? "" };
      }
      entitlementsPathOverride = value;
      i += 1;
      continue;
    }

    if (arg.startsWith("-")) {
      errBlock("macOS entitlements preflight failed", [
        `Unknown argument: ${arg}`,
        `Run with --help for usage.`,
      ]);
      return { repoRoot, entitlementsPathOverride: entitlementsPathOverride ?? "" };
    }
  }

  const resolvedEntitlementsPathOverride = entitlementsPathOverride
    ? path.isAbsolute(entitlementsPathOverride)
      ? entitlementsPathOverride
      : path.resolve(repoRoot, entitlementsPathOverride)
    : "";

  return { repoRoot, entitlementsPathOverride: resolvedEntitlementsPathOverride };
}

function main() {
  const { repoRoot, entitlementsPathOverride } = parseArgs(process.argv.slice(2));
  if (process.exitCode) return;
  const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const relativeConfigPath = path.relative(repoRoot, configPath);

  let entitlementsPath = entitlementsPathOverride;

  if (!entitlementsPath) {
    /** @type {any} */
    let config;
    try {
      config = JSON.parse(readFileSync(configPath, "utf8"));
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      errBlock(`macOS entitlements preflight failed`, [
        `Failed to read/parse ${relativeConfigPath}.`,
        `Error: ${msg}`,
      ]);
      return;
    }

    const entitlementsSetting = config?.bundle?.macOS?.entitlements;
    if (typeof entitlementsSetting !== "string" || entitlementsSetting.trim().length === 0) {
      errBlock(`Invalid macOS signing config (${relativeConfigPath})`, [
        `Expected bundle.macOS.entitlements to be a non-empty string path.`,
        `This repo requires an explicit entitlements plist so WKWebView/JavaScriptCore can run under the hardened runtime.`,
      ]);
      return;
    }

    entitlementsPath = path.resolve(path.dirname(configPath), entitlementsSetting);
  }

  if (!entitlementsPath) {
    // Argument parsing already reported an error.
    return;
  }

  const relativeEntitlementsPath = path.relative(repoRoot, entitlementsPath);

  // On macOS runners, validate the plist is syntactically valid using the system tool.
  // (On other platforms, we avoid depending on `plutil`.)
  if (process.platform === "darwin") {
    const lint = spawnSync("plutil", ["-lint", entitlementsPath], { encoding: "utf8" });
    if (lint.error) {
      errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
        `Failed to run plutil -lint.`,
        `Error: ${lint.error instanceof Error ? lint.error.message : String(lint.error)}`,
      ]);
      return;
    }
    if (lint.status !== 0) {
      const output = `${lint.stdout ?? ""}\n${lint.stderr ?? ""}`.trim();
      errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
        `plutil -lint failed.`,
        ...(output ? [`Output:\n${output}`] : []),
      ]);
      return;
    }
  }

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
