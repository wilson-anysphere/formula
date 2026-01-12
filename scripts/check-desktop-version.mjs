import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
const tauriConfigPath = path.join(repoRoot, tauriConfigRelativePath);

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

if (desktopVersion !== tagVersion) {
  console.error("Desktop version/tag mismatch detected.");
  console.error(`- Git tag: ${refName} (expects version "${tagVersion}")`);
  console.error(`- ${tauriConfigRelativePath}: version "${desktopVersion}"`);
  console.error("");
  console.error("Fix:");
  console.error(`1) Bump "version" in ${tauriConfigRelativePath} to "${tagVersion}".`);
  console.error("2) Commit the change.");
  console.error(
    `3) Re-tag the release at v${tagVersion} (delete/re-create the existing tag if needed).`,
  );
  console.error("");
  console.error("This check prevents publishing broken desktop updater metadata for the wrong version.");
  process.exit(1);
}

console.log(`Desktop version check passed: tag ${refName} matches tauri.conf.json version ${desktopVersion}.`);
