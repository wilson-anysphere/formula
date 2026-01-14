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
    const base = path.basename(src);
    if (base !== filename) {
      die(
        `bundle.linux.${kind}.files[${JSON.stringify(destPath)}] must point at ${filename} (got ${JSON.stringify(src)})`,
      );
    }
  }
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

console.log(`desktop-compliance: OK (LICENSE/NOTICE bundling configured for ${mainBinaryName})`);
