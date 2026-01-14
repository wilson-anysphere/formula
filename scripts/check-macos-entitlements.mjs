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

    if (/<true\s*\/>/.test(xml.slice(i, i + 10))) return true;
    start = xml.indexOf(marker, i);
  }

  return false;
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

  if (!xml.includes("<plist") || !xml.includes("<dict")) {
    errBlock(`Invalid entitlements plist (${relativeEntitlementsPath})`, [
      `File does not look like a plist (<plist>/<dict> tags not found).`,
      `Expected an XML plist file containing the macOS code signing entitlements.`,
    ]);
    return;
  }

  // Hardened Runtime + WKWebView (wry) commonly require these two entitlements
  // for JavaScript/WASM execution in signed/notarized builds.
  //
  // We also require outbound network entitlement so that if/when we enable the App Sandbox,
  // core app functionality (updater, HTTPS fetches) doesn't break silently.
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

  const missing = required.filter(({ key }) => !hasTrueEntitlement(xml, key));
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

  const forbiddenEnabled = forbiddenTrue.filter(({ key }) => hasTrueEntitlement(xml, key));
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
