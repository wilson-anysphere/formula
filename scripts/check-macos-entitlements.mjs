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
 * We avoid fully parsing plist XML so the check can run on non-macOS CI runners without
 * depending on `plutil`. This helper is resilient to duplicate keys: it returns true if
 * *any* instance of the key is followed by a `<true/>` value.
 *
 * @param {string} xml
 * @param {string} key
 * @returns {boolean}
 */
function hasTrueEntitlement(xml, key) {
  const marker = `<key>${key}</key>`;
  let start = xml.indexOf(marker);
  while (start !== -1) {
    let i = start + marker.length;

    while (i < xml.length) {
      // Skip whitespace.
      if (/\s/.test(xml[i])) {
        i += 1;
        continue;
      }

      // Skip XML comments.
      if (xml.startsWith("<!--", i)) {
        const end = xml.indexOf("-->", i + 4);
        if (end === -1) break;
        i = end + 3;
        continue;
      }

      break;
    }

    // `<true/>` is typically short, but tolerate weird formatting like `<true     />`.
    // Anchor at the start so we don't accidentally match the `<true/>` value for a different key.
    if (/^<true\s*\/>/.test(xml.slice(i, i + 64))) return true;
    start = xml.indexOf(marker, i);
  }

  return false;
}

/**
 * Best-effort scan of entitlement keys in a plist XML string.
 *
 * This is intentionally lightweight so it can run on non-macOS CI runners without
 * depending on `plutil` or third-party plist parsers.
 *
 * @param {string} xml
 * @returns {string[]}
 */
function extractEntitlementKeys(xml) {
  /** @type {string[]} */
  const keys = [];
  const re = /<key>([^<]+)<\/key>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const key = m[1]?.trim();
    if (key) keys.push(key);
  }
  return keys;
}

/**
 * @param {string[]} argv
 * @returns {{ repoRoot: string; entitlementsPathOverride: string }}
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
          "This script also guards against accidentally enabling broad/debug entitlements in Developer ID builds.",
          "",
          "Usage:",
          "  node scripts/check-macos-entitlements.mjs",
          "  node scripts/check-macos-entitlements.mjs --root <repoRoot>",
          "  node scripts/check-macos-entitlements.mjs --path <entitlements.plist>",
          "",
           "Defaults:",
           "  --root defaults to the repository root (derived from this script's location).",
           "  --path defaults to the file referenced by bundle.macOS.entitlements in tauri.conf.json.",
           "  - tauri.conf.json defaults to apps/desktop/src-tauri/tauri.conf.json (override via FORMULA_TAURI_CONF_PATH).",
           "",
           "Notes:",
          "  - When com.apple.security.app-sandbox is enabled, the guardrail also requires",
          "    com.apple.security.network.server (Formula runs an OAuth loopback redirect listener).",
          "",
        ].join("\n"),
      );
      process.exit(0);
    }

    if (arg === "--root") {
      const value = argv[i + 1];
      if (!value) {
        errBlock("macOS entitlements preflight failed", [`Missing value for --root.`]);
        return { repoRoot, entitlementsPathOverride: entitlementsPathOverride ?? "" };
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

    // Disallow positional args to avoid silently ignoring typos (e.g. passing a path without --path).
    if (typeof arg === "string" && arg.trim().length > 0) {
      errBlock("macOS entitlements preflight failed", [
        `Unknown argument: ${arg}`,
        `Expected flags only (use --path to override the entitlements plist).`,
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
  const defaultConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
  const configPathOverride = process.env.FORMULA_TAURI_CONF_PATH;
  const configPath =
    configPathOverride && String(configPathOverride).trim()
      ? path.isAbsolute(String(configPathOverride).trim())
        ? String(configPathOverride).trim()
        : path.resolve(repoRoot, String(configPathOverride).trim())
      : path.join(repoRoot, defaultConfigRelativePath);
  const relativeConfigPath = path.relative(repoRoot, configPath) || defaultConfigRelativePath;

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

  // Strip XML comments so we don't accidentally treat commented-out entitlements as enabled.
  // (Some developers temporarily comment out `<key>...</key>` blocks during debugging.)
  const xmlForScan = xml.replace(/<!--[\s\S]*?-->/g, "");

  if (!xmlForScan.includes("<plist") || !xmlForScan.includes("<dict")) {
    errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
      `File does not look like a plist (<plist>/<dict> tags not found).`,
      `Expected an XML plist file containing the macOS code signing entitlements.`,
    ]);
    return;
  }

  const allKeys = extractEntitlementKeys(xmlForScan);
  if (allKeys.length === 0) {
    errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
      `No <key>...</key> entries found.`,
      `Expected at least one entitlement key/value in the top-level plist <dict>.`,
    ]);
    return;
  }

  /** @type {Map<string, number>} */
  const keyCounts = new Map();
  for (const key of allKeys) {
    keyCounts.set(key, (keyCounts.get(key) ?? 0) + 1);
  }

  const duplicateKeys = [...keyCounts.entries()]
    .filter(([, count]) => count > 1)
    .map(([key, count]) => `${key} (x${count})`);
  if (duplicateKeys.length > 0) {
    errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
      `Duplicate keys detected (plist dict keys should be unique):`,
      ...duplicateKeys,
    ]);
  }

  const nonTrueKeys = [...keyCounts.keys()].filter((key) => !hasTrueEntitlement(xmlForScan, key));
  if (nonTrueKeys.length > 0) {
    errBlock(`Invalid macOS entitlements (${relativeEntitlementsPath})`, [
      `All entitlement keys must be set to boolean <true/> (no <false/>, strings, arrays, or dict values).`,
      ...nonTrueKeys.map((key) => `${key} — not set to <true/>`),
    ]);
  }

  // Hardened Runtime + WKWebView (wry) commonly require these two entitlements
  // for JavaScript/WASM execution in signed/notarized builds.
  //
  // We also require outbound network entitlement so that if/when we enable the App Sandbox,
  // core app functionality (updater, HTTPS fetches) doesn't break silently.
  /** @type {{ key: string; reason: string }[]} */
  const required = [
    {
      key: "com.apple.security.cs.allow-jit",
      reason: "WKWebView/JavaScriptCore JIT (blank WebView/crash if missing under hardened runtime).",
    },
    {
      key: "com.apple.security.cs.allow-unsigned-executable-memory",
      reason: "WKWebView/JavaScriptCore JIT executable memory (required for JS/WASM).",
    },
    {
      key: "com.apple.security.network.client",
      reason:
        "Outbound network access (required for updater/HTTPS when sandboxing is enabled).",
    },
  ];

  // If someone opts into the App Sandbox, ensure we include the additional sandbox entitlements
  // required by Formula's runtime features (notably the OAuth loopback redirect listener).
  if (hasTrueEntitlement(xmlForScan, "com.apple.security.app-sandbox")) {
    required.push({
      key: "com.apple.security.network.server",
      reason:
        "Incoming network access (required when sandboxing is enabled because Formula runs a loopback HTTP listener for OAuth redirects).",
    });
  }

  /** @type {Set<string>} */
  const allowlisted = new Set(required.map(({ key }) => key));
  if (hasTrueEntitlement(xmlForScan, "com.apple.security.app-sandbox")) {
    // App Sandbox is currently unused in this repo, but we allow it as an opt-in
    // (and already require network.server above when enabled).
    allowlisted.add("com.apple.security.app-sandbox");
  }
  const unexpectedKeys = [...keyCounts.keys()].filter((key) => !allowlisted.has(key));
  if (unexpectedKeys.length > 0) {
    errBlock(`Unexpected macOS entitlements enabled (${relativeEntitlementsPath})`, [
      `Keep the signed entitlement surface minimal. The following keys are present but not allowlisted:`,
      ...unexpectedKeys,
      ``,
      `If you believe an additional entitlement is required, add a justification to entitlements.plist and update scripts/check-macos-entitlements.mjs accordingly.`,
    ]);
  }

  const missing = required.filter(({ key }) => !hasTrueEntitlement(xmlForScan, key));
  if (missing.length > 0) {
    errBlock(`Invalid macOS entitlements (${relativeEntitlementsPath})`, [
      `Missing required entitlement(s) (must be present and set to <true/>):`,
      ...missing.map(({ key, reason }) => `${key} — ${reason}`),
      ``,
      `Common symptom for missing WKWebView JIT entitlements: a signed/notarized build launches with a blank window.`,
    ]);
  }

  // Guardrail against broad/production-inappropriate entitlements being enabled by accident.
  const forbiddenTrue = [
    {
      key: "com.apple.security.get-task-allow",
      reason: "Debug entitlement (should never be true for Developer ID distribution).",
    },
    {
      key: "com.apple.security.cs.disable-library-validation",
      reason: "Allows loading unsigned/untrusted dylibs; only enable if absolutely required.",
    },
    {
      key: "com.apple.security.cs.disable-executable-page-protection",
      reason: "Broader W+X memory permission; prefer targeted JIT entitlements instead.",
    },
    {
      key: "com.apple.security.cs.allow-dyld-environment-variables",
      reason: "Allows DYLD_* injection; avoid in production builds.",
    },
  ];

  const forbiddenEnabled = forbiddenTrue.filter(({ key }) => hasTrueEntitlement(xmlForScan, key));
  if (forbiddenEnabled.length > 0) {
    errBlock(`Disallowed macOS entitlements enabled (${relativeEntitlementsPath})`, [
      `The following entitlements are present and set to <true/>:`,
      ...forbiddenEnabled.map(({ key, reason }) => `${key} — ${reason}`),
      ``,
      `If you believe one of these is required, update entitlements.plist with a clear justification and adjust this guardrail accordingly.`,
    ]);
  }

  if (process.exitCode) {
    err(`\nmacOS entitlements preflight failed.\n`);
    return;
  }

  console.log(`macOS entitlements preflight passed (${relativeEntitlementsPath}).`);
}

main();
