import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
const tauriConfigPath = path.join(repoRoot, tauriConfigRelativePath);
const cargoManifestRelativePath = "apps/desktop/src-tauri/Cargo.toml";
const cargoManifestPath = path.join(repoRoot, cargoManifestRelativePath);

const refName = process.argv[2] ?? process.env.GITHUB_REF_NAME;
if (!refName) {
  console.error(
    "Missing tag name. Usage: node scripts/check-desktop-version.mjs <tag> (example: v0.2.3)",
  );
  process.exit(1);
}

const normalizedRefName = refName.startsWith("refs/tags/")
  ? refName.slice("refs/tags/".length)
  : refName;
const tagVersion = normalizedRefName.startsWith("v") ? normalizedRefName.slice(1) : normalizedRefName;

/**
 * Extract `[package].version` from a Cargo.toml text.
 *
 * We intentionally avoid pulling in a TOML parser to keep this script dependency-free.
 * This expects a standard manifest entry like:
 *   [package]
 *   version = "1.2.3"
 *
 * @param {string} tomlText
 * @returns {string}
 */
function extractCargoPackageVersion(tomlText) {
  const lines = tomlText.split(/\r?\n/);
  let inPackage = false;
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (line.startsWith("[") && line.endsWith("]")) {
      inPackage = line === "[package]";
      continue;
    }
    if (!inPackage || line.length === 0 || line.startsWith("#")) {
      continue;
    }
    const match = line.match(/^version\s*=\s*["']([^"']+)["']\s*(?:#.*)?$/);
    if (match) {
      return match[1].trim();
    }
  }
  return "";
}

/** @type {any} */
let config;
try {
  const configText = await readFile(tauriConfigPath, "utf8");
  config = JSON.parse(configText);
} catch (err) {
  console.error(`Failed to read/parse ${tauriConfigRelativePath}.`);
  console.error(err);
  process.exit(1);
}

const desktopVersion = typeof config?.version === "string" ? config.version.trim() : "";
if (!desktopVersion) {
  console.error(`Expected ${tauriConfigRelativePath} to contain a non-empty "version" field.`);
  process.exit(1);
}

let cargoVersionText = "";
try {
  cargoVersionText = await readFile(cargoManifestPath, "utf8");
} catch (err) {
  console.error(`Failed to read ${cargoManifestRelativePath}.`);
  console.error(err);
  process.exit(1);
}

const cargoVersion = extractCargoPackageVersion(cargoVersionText);
if (!cargoVersion) {
  console.error(
    `Expected ${cargoManifestRelativePath} to contain a [package] version = "..." entry.`,
  );
  process.exit(1);
}

if (desktopVersion !== tagVersion || cargoVersion !== tagVersion) {
  console.error("Desktop version mismatch detected.");
  console.error(`- Git tag: ${normalizedRefName} (expects version "${tagVersion}")`);
  console.error(`- ${tauriConfigRelativePath}: version "${desktopVersion}"`);
  console.error(`- ${cargoManifestRelativePath}: [package].version "${cargoVersion}"`);
  console.error("");
  console.error("Fix:");
  let step = 1;
  if (desktopVersion !== tagVersion) {
    console.error(`${step}) Bump "version" in ${tauriConfigRelativePath} to "${tagVersion}".`);
    step += 1;
  }
  if (cargoVersion !== tagVersion) {
    console.error(`${step}) Bump [package].version in ${cargoManifestRelativePath} to "${tagVersion}".`);
    step += 1;
  }
  console.error(`${step}) Commit the change${step > 2 ? "s" : ""}.`);
  step += 1;
  console.error(`${step}) Re-tag the release at v${tagVersion} (delete/re-create the existing tag if needed).`);
  console.error("");
  console.error(
    "This check prevents publishing broken desktop artifacts/updater metadata when the git tag, tauri.conf.json, and Cargo.toml disagree.",
  );
  process.exit(1);
}

console.log(
  `Desktop version check passed: tag ${normalizedRefName} matches tauri.conf.json (${desktopVersion}) and Cargo.toml (${cargoVersion}).`,
);
