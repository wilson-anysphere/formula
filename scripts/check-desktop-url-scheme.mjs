#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const tauriConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const infoPlistPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "Info.plist");

const REQUIRED_SCHEME = "formula";

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
 * @param {unknown} pluginConfig
 * @returns {string[]}
 */
function extractDeepLinkSchemes(pluginConfig) {
  if (!pluginConfig || typeof pluginConfig !== "object") return [];

  const desktop = /** @type {any} */ (pluginConfig).desktop;
  if (!desktop) return [];

  // plugins.deep-link.desktop is a `DeepLinkProtocol` or an array of `DeepLinkProtocol`.
  if (Array.isArray(desktop)) {
    return desktop.flatMap((p) => (Array.isArray(p?.schemes) ? p.schemes : [])).filter(Boolean);
  }
  if (typeof desktop === "object") {
    return Array.isArray(desktop.schemes) ? desktop.schemes.filter(Boolean) : [];
  }
  return [];
}

function main() {
  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(tauriConfigPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Desktop URL scheme preflight failed", [
      "Failed to read/parse apps/desktop/src-tauri/tauri.conf.json.",
      `Error: ${msg}`,
    ]);
    return;
  }

  // ---- macOS: Info.plist contains CFBundleURLSchemes -> formula
  try {
    const plist = readFileSync(infoPlistPath, "utf8");
    // Keep this lightweight: we just need to know the scheme is present.
    if (!plist.includes("<key>CFBundleURLSchemes</key>") || !plist.includes(`<string>${REQUIRED_SCHEME}</string>`)) {
      errBlock("Missing macOS URL scheme registration (Info.plist)", [
        "Expected apps/desktop/src-tauri/Info.plist to declare CFBundleURLSchemes including:",
        `  - ${REQUIRED_SCHEME}`,
        "Fix: add/update CFBundleURLTypes/CFBundleURLSchemes in Info.plist.",
      ]);
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Desktop URL scheme preflight failed", [
      "Failed to read apps/desktop/src-tauri/Info.plist.",
      `Error: ${msg}`,
    ]);
  }

  // ---- All platforms: bundler must know about formula:// so installers register it.
  const deepLinkConfig = config?.plugins?.["deep-link"];
  const schemes = extractDeepLinkSchemes(deepLinkConfig);
  const normalizedSchemes = schemes.map((s) => String(s).trim()).filter(Boolean);

  if (!normalizedSchemes.includes(REQUIRED_SCHEME)) {
    const found = normalizedSchemes.length > 0 ? normalizedSchemes.join(", ") : "(none)";
    errBlock("Missing desktop deep-link scheme configuration (tauri.conf.json)", [
      "Expected apps/desktop/src-tauri/tauri.conf.json to include:",
      `  plugins[\"deep-link\"].desktop.schemes = [\"${REQUIRED_SCHEME}\"]`,
      `Found schemes: ${found}`,
      "Fix: add the deep-link plugin config so installers register the URL scheme.",
    ]);
  }

  // ---- Linux: freedesktop .desktop generation only includes explicit file association mimeType values.
  const fileAssociations = config?.bundle?.fileAssociations;
  if (!Array.isArray(fileAssociations) || fileAssociations.length === 0) {
    errBlock("Missing bundle.fileAssociations (tauri.conf.json)", [
      "Expected apps/desktop/src-tauri/tauri.conf.json to include bundle.fileAssociations so Linux .desktop metadata can include file MIME types.",
    ]);
  } else {
    const missingMime = fileAssociations
      .map((assoc, i) => ({ assoc, i }))
      .filter(({ assoc }) => typeof assoc?.mimeType !== "string" || assoc.mimeType.trim().length === 0);

    if (missingMime.length > 0) {
      errBlock("Missing Linux mimeType fields in bundle.fileAssociations", [
        "Tauri's Linux .desktop file generation only includes file MIME types when bundle.fileAssociations[].mimeType is set.",
        ...missingMime.map(({ assoc, i }) => {
          const exts = Array.isArray(assoc?.ext) ? assoc.ext.join(", ") : "(missing ext)";
          return `fileAssociations[${i}] ext=[${exts}] is missing mimeType`;
        }),
        "Fix: add mimeType to each file association entry (Linux-only field).",
      ]);
    }
  }

  if (process.exitCode) {
    err("\nDesktop URL scheme preflight failed. Fix the errors above before tagging a release.\n");
    return;
  }

  console.log(
    `Desktop URL scheme preflight passed: ${REQUIRED_SCHEME}:// is configured for installers (and Info.plist declares it).`
  );
}

main();

