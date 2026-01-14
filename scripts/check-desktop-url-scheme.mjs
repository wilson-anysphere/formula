#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));

/**
 * Resolve a path that may be absolute or repo-relative.
 * @param {string | undefined} value
 * @param {string} fallback
 */
function resolvePath(value, fallback) {
  if (value && String(value).trim()) {
    const p = String(value).trim();
    return path.isAbsolute(p) ? p : path.join(repoRoot, p);
  }
  return fallback;
}

const tauriConfigPath = resolvePath(
  process.env.FORMULA_TAURI_CONF_PATH,
  path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"),
);
const infoPlistPath = resolvePath(
  process.env.FORMULA_INFO_PLIST_PATH,
  path.join(repoRoot, "apps", "desktop", "src-tauri", "Info.plist"),
);
const parquetMimeDefinitionPath = path.join(
  repoRoot,
  "apps",
  "desktop",
  "src-tauri",
  "mime",
  "app.formula.desktop.xml",
);

const REQUIRED_SCHEME = "formula";
const REQUIRED_FILE_EXT = "xlsx";

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
    return desktop
      .flatMap((p) => {
        const raw = p?.schemes;
        if (typeof raw === "string") return [raw];
        if (Array.isArray(raw)) return raw;
        return [];
      })
      .filter(Boolean);
  }
  if (typeof desktop === "object") {
    const raw = desktop.schemes;
    if (typeof raw === "string") return [raw].filter(Boolean);
    if (Array.isArray(raw)) return raw.filter(Boolean);
    return [];
  }
  return [];
}

/**
 * @param {any} config
 */
function isParquetAssociationConfigured(config) {
  const fileAssociations = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  return fileAssociations.some((assoc) => {
    const rawExt = /** @type {any} */ (assoc)?.ext;
    const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
    return exts.some((e) => String(e).trim().toLowerCase().replace(/^\./, "") === "parquet");
  });
}

/**
 * Parquet is not consistently defined in distros' shared-mime-info DB.
 *
 * If we advertise Parquet (`application/vnd.apache.parquet`) in `.desktop` MimeType=
 * we should also ship a shared-mime-info definition file inside Linux bundles so
 * `*.parquet` resolves to that MIME type after install (via update-mime-database).
 *
 * @param {any} config
 */
function validateParquetMimeDefinition(config) {
  if (!isParquetAssociationConfigured(config)) return;

  const linux = config?.bundle?.linux;
  if (!linux || typeof linux !== "object") {
    errBlock("Parquet file association configured, but bundle.linux is missing (tauri.conf.json)", [
      "Expected bundle.linux.{deb,rpm,appimage} to be configured so we can ship a shared-mime-info definition for Parquet.",
    ]);
    return;
  }

  const expectedDest = "usr/share/mime/packages/app.formula.desktop.xml";
  const expectedSrc = "mime/app.formula.desktop.xml";

  for (const target of ["deb", "rpm", "appimage"]) {
    const files = linux?.[target]?.files;
    if (!files || typeof files !== "object") {
      errBlock("Parquet file association configured, but Linux bundle file mappings are missing", [
        `Expected bundle.linux.${target}.files to map:`,
        `  - ${expectedDest} -> ${expectedSrc}`,
      ]);
      continue;
    }
    if (files[expectedDest] !== expectedSrc) {
      errBlock("Parquet shared-mime-info mapping mismatch (tauri.conf.json)", [
        `Expected bundle.linux.${target}.files["${expectedDest}"] = "${expectedSrc}"`,
        `Found: ${JSON.stringify(files[expectedDest] ?? null)}`,
      ]);
    }
  }

  // Ensure update-mime-database triggers exist at install time.
  const debDepends = linux?.deb?.depends;
  if (!Array.isArray(debDepends) || !debDepends.includes("shared-mime-info")) {
    errBlock("Parquet file association configured, but shared-mime-info is not declared as a DEB dependency", [
      'Expected bundle.linux.deb.depends to include "shared-mime-info" so update-mime-database triggers run.',
    ]);
  }
  const rpmDepends = linux?.rpm?.depends;
  if (!Array.isArray(rpmDepends) || !rpmDepends.includes("shared-mime-info")) {
    errBlock("Parquet file association configured, but shared-mime-info is not declared as an RPM dependency", [
      'Expected bundle.linux.rpm.depends to include "shared-mime-info" so update-mime-database triggers run.',
    ]);
  }

  try {
    const xml = readFileSync(parquetMimeDefinitionPath, "utf8");
    if (!xml.includes('mime-type type="application/vnd.apache.parquet"') || !xml.includes('glob pattern="*.parquet"')) {
      errBlock("Parquet shared-mime-info definition file is missing expected content", [
        `File: ${parquetMimeDefinitionPath}`,
        'Expected to find: mime-type type="application/vnd.apache.parquet"',
        'Expected to find: glob pattern="*.parquet"',
      ]);
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Parquet file association configured, but shared-mime-info definition file is missing", [
      `Expected file to exist: ${parquetMimeDefinitionPath}`,
      `Error: ${msg}`,
    ]);
  }
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
    const hasRequiredExt = fileAssociations.some((assoc) => {
      const rawExt = /** @type {any} */ (assoc)?.ext;
      const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
      return exts.some((e) => String(e).trim().toLowerCase().replace(/^\./, "") === REQUIRED_FILE_EXT);
    });

    if (!hasRequiredExt) {
      errBlock("Missing desktop file association configuration (tauri.conf.json)", [
        `Expected apps/desktop/src-tauri/tauri.conf.json bundle.fileAssociations to include an entry for .${REQUIRED_FILE_EXT}.`,
        `Fix: add { ext: [\"${REQUIRED_FILE_EXT}\"], mimeType: \"...\" } to bundle.fileAssociations.`,
      ]);
    }

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

  validateParquetMimeDefinition(config);

  if (process.exitCode) {
    err("\nDesktop URL scheme preflight failed. Fix the errors above before tagging a release.\n");
    return;
  }

  console.log(
    `Desktop URL scheme preflight passed: ${REQUIRED_SCHEME}:// is configured for installers (and Info.plist declares it).`
  );
}

main();
