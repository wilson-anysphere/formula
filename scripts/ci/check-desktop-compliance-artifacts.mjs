import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const configPath = process.env.FORMULA_TAURI_CONF_PATH
  ? path.resolve(process.env.FORMULA_TAURI_CONF_PATH)
  : defaultConfigPath;

function die(message) {
  console.error(`desktop-compliance: ERROR ${message}`);
  process.exit(1);
}

/**
 * @param {unknown} value
 * @returns {Record<string, unknown> | null}
 */
function asObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  // @ts-ignore - runtime type guard
  return value;
}

/**
 * @param {unknown} resources
 * @param {string} filename
 * @returns {boolean}
 */
function resourcesInclude(resources, filename) {
  if (!Array.isArray(resources)) return false;
  return resources.some((entry) => typeof entry === "string" && path.basename(entry) === filename);
}

/**
 * @param {string} src
 * @param {string} filename
 * @returns {string}
 */
function resolveAndAssertFileExists(src, filename) {
  const base = path.basename(src);
  if (base !== filename) {
    die(`expected source path basename to be ${filename} (got ${JSON.stringify(src)})`);
  }

  const resolved = path.isAbsolute(src) ? src : path.resolve(path.dirname(configPath), src);
  let stat;
  try {
    stat = fs.statSync(resolved);
  } catch (err) {
    die(`missing source file for ${filename}: ${resolved}`);
  }
  if (!stat.isFile()) {
    die(`expected source file for ${filename} to be a file: ${resolved}`);
  }
  return resolved;
}

/**
 * @param {unknown} files
 * @param {string} mainBinaryName
 * @param {string} kind
 */
function validateLinuxFiles(files, mainBinaryName, kind) {
  const obj = asObject(files);
  if (!obj) {
    die(`bundle.linux.${kind}.files must be an object mapping destination -> source`);
  }

  const requiredDest = {
    LICENSE: `usr/share/doc/${mainBinaryName}/LICENSE`,
    NOTICE: `usr/share/doc/${mainBinaryName}/NOTICE`,
  };

  for (const [filename, destPath] of Object.entries(requiredDest)) {
    if (!(destPath in obj)) {
      const keys = Object.keys(obj);
      die(
        `bundle.linux.${kind}.files is missing ${JSON.stringify(destPath)} (for ${filename}). Present keys: ${JSON.stringify(keys)}`,
      );
    }
    const src = obj[destPath];
    if (typeof src !== "string") {
      die(
        `bundle.linux.${kind}.files[${JSON.stringify(destPath)}] must be a string source path (got ${typeof src})`,
      );
    }
    resolveAndAssertFileExists(src, filename);
  }
}

/**
 * @param {unknown} fileAssociations
 * @returns {boolean}
 */
function hasParquetAssociation(fileAssociations) {
  if (!Array.isArray(fileAssociations)) return false;
  for (const assoc of fileAssociations) {
    if (!assoc || typeof assoc !== "object" || Array.isArray(assoc)) continue;
    // @ts-ignore - runtime inspection
    const mimeType = assoc.mimeType;
    if (
      typeof mimeType === "string" &&
      mimeType.trim().toLowerCase() === "application/vnd.apache.parquet"
    ) {
      return true;
    }
    // @ts-ignore - runtime inspection
    const ext = assoc.ext;
    if (Array.isArray(ext) && ext.some((v) => typeof v === "string" && v.trim().toLowerCase() === "parquet")) {
      return true;
    }
  }
  return false;
}

/**
 * @param {unknown} depends
 * @param {string} kind
 */
function validateLinuxDepends(depends, kind) {
  if (!Array.isArray(depends)) {
    die(`bundle.linux.${kind}.depends must be an array of package dependency strings`);
  }
  const deps = depends.map((v) => String(v));
  if (!deps.some((v) => v.toLowerCase().includes("shared-mime-info"))) {
    die(`bundle.linux.${kind}.depends must include shared-mime-info (required for MIME database integration)`);
  }
}

/**
 * @param {unknown} files
 * @param {string} kind
 */
function validateLinuxMimeFiles(files, kind) {
  const obj = asObject(files);
  if (!obj) {
    die(`bundle.linux.${kind}.files must be an object mapping destination -> source`);
  }
  const dest = "usr/share/mime/packages/app.formula.desktop.xml";
  if (!(dest in obj)) {
    const keys = Object.keys(obj);
    die(
      `bundle.linux.${kind}.files is missing ${JSON.stringify(dest)} (required for Parquet/shared-mime-info integration). Present keys: ${JSON.stringify(keys)}`,
    );
  }
  const src = obj[dest];
  if (typeof src !== "string") {
    die(`bundle.linux.${kind}.files[${JSON.stringify(dest)}] must be a string source path (got ${typeof src})`);
  }
  resolveAndAssertFileExists(src, "app.formula.desktop.xml");
}

let raw = "";
try {
  raw = fs.readFileSync(configPath, "utf8");
} catch (err) {
  die(`failed to read config: ${configPath}\n${err}`);
}

let config;
try {
  config = JSON.parse(raw);
} catch (err) {
  die(`failed to parse JSON: ${configPath}\n${err}`);
}

const mainBinaryName = String(config?.mainBinaryName ?? "").trim() || "formula-desktop";
const bundle = asObject(config?.bundle) ?? {};

const resources = bundle.resources;
if (!resourcesInclude(resources, "LICENSE")) {
  die(
    `bundle.resources must include LICENSE (expected an entry whose basename is "LICENSE"). Found: ${JSON.stringify(resources ?? null)}`,
  );
}
if (!resourcesInclude(resources, "NOTICE")) {
  die(
    `bundle.resources must include NOTICE (expected an entry whose basename is "NOTICE"). Found: ${JSON.stringify(resources ?? null)}`,
  );
}

const linux = asObject(bundle.linux) ?? {};
validateLinuxFiles(asObject(linux.deb)?.files, mainBinaryName, "deb");
validateLinuxFiles(asObject(linux.rpm)?.files, mainBinaryName, "rpm");

// Tauri's config key is `appimage` (lowercase).
validateLinuxFiles(asObject(linux.appimage)?.files, mainBinaryName, "appimage");

// When the app advertises Parquet support on Linux, we also ship a shared-mime-info XML definition
// so the MIME database can map `*.parquet` correctly even on distros whose shared-mime-info
// package does not include a Parquet glob by default.
if (hasParquetAssociation(bundle.fileAssociations)) {
  validateLinuxMimeFiles(asObject(linux.deb)?.files, "deb");
  validateLinuxMimeFiles(asObject(linux.rpm)?.files, "rpm");
  validateLinuxMimeFiles(asObject(linux.appimage)?.files, "appimage");

  // Ensure the distro-native packages pull in `update-mime-database` on install.
  validateLinuxDepends(asObject(linux.deb)?.depends, "deb");
  validateLinuxDepends(asObject(linux.rpm)?.depends, "rpm");
}

// Ensure the declared resources actually exist on disk as files.
if (!Array.isArray(resources)) {
  die("bundle.resources must be an array");
}
const licenseEntry = resources.find((entry) => typeof entry === "string" && path.basename(entry) === "LICENSE");
const noticeEntry = resources.find((entry) => typeof entry === "string" && path.basename(entry) === "NOTICE");
if (typeof licenseEntry !== "string") {
  die("bundle.resources must contain an entry for LICENSE");
}
if (typeof noticeEntry !== "string") {
  die("bundle.resources must contain an entry for NOTICE");
}
resolveAndAssertFileExists(licenseEntry, "LICENSE");
resolveAndAssertFileExists(noticeEntry, "NOTICE");

console.log(`desktop-compliance: OK (LICENSE/NOTICE bundling configured for ${mainBinaryName})`);
