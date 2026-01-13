import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

const cargoTomlRelativePath = "apps/desktop/src-tauri/Cargo.toml";
const cargoTomlPath = path.join(repoRoot, cargoTomlRelativePath);

const releaseWorkflowRelativePath = ".github/workflows/release.yml";
const releaseWorkflowPath = path.join(repoRoot, releaseWorkflowRelativePath);

function parseMajorMinor(version) {
  const normalized = String(version ?? "").trim().replace(/^[^0-9]*/, "");
  const match = normalized.match(/^(\d+)\.(\d+)/);
  if (!match) {
    throw new Error(`Expected a semver-ish version like "2.9.5" (got "${version}")`);
  }
  return `${match[1]}.${match[2]}`;
}

function parsePinnedCliVersion(version) {
  const normalized = String(version ?? "").trim();
  // For determinism, TAURI_CLI_VERSION must be a fully pinned patch version.
  //
  // `cargo install tauri-cli --version 2.9` would float to the latest 2.9.x and
  // would therefore not be reproducible when new patch releases land.
  const cleaned = normalized.replace(/^v/, "");
  const match = cleaned.match(/^(\d+)\.(\d+)\.(\d+)(?:[-+].*)?$/);
  if (!match) {
    throw new Error(
      `Expected TAURI_CLI_VERSION to be pinned to an exact patch version like "2.9.5" (got "${version}")`,
    );
  }
  return { majorMinor: `${match[1]}.${match[2]}` };
}

function extractTauriVersionFromCargoToml(tomlText) {
  // Common forms:
  //   tauri = "2.9"
  //   tauri = { version = "2.9", optional = true, features = [...] }
  const objectMatch = tomlText.match(
    /^[\t ]*tauri[\t ]*=[\t ]*\{[^}]*\bversion[\t ]*=[\t ]*"([^"]+)"[^}]*\}/m,
  );
  if (objectMatch) return objectMatch[1];

  const stringMatch = tomlText.match(/^[\t ]*tauri[\t ]*=[\t ]*"([^"]+)"/m);
  if (stringMatch) return stringMatch[1];

  return null;
}

function extractPinnedCliVersionFromWorkflow(workflowText) {
  const match = workflowText.match(/^[\t ]*TAURI_CLI_VERSION:[\t ]*["']?([^"'\n]+)["']?/m);
  return match ? match[1].trim() : null;
}

let cargoTomlText = "";
try {
  cargoTomlText = await readFile(cargoTomlPath, "utf8");
} catch (err) {
  console.error(`Failed to read ${cargoTomlRelativePath}.`);
  console.error(err);
  process.exit(1);
}

const tauriVersion = extractTauriVersionFromCargoToml(cargoTomlText);
if (!tauriVersion) {
  console.error(
    `Failed to locate a tauri dependency version in ${cargoTomlRelativePath} (expected e.g. tauri = "2.9").`,
  );
  process.exit(1);
}

const tauriMajorMinor = parseMajorMinor(tauriVersion);

const cliVersionFromEnv = process.env.TAURI_CLI_VERSION;
const cliVersionFromArg = process.argv[2];

const pinnedSource = cliVersionFromEnv ? "env" : cliVersionFromArg ? "arg" : "workflow";

let pinnedCliVersion = cliVersionFromEnv || cliVersionFromArg || null;
let releaseWorkflowPinnedCliVersion = null;
try {
  const workflowText = await readFile(releaseWorkflowPath, "utf8");
  releaseWorkflowPinnedCliVersion = extractPinnedCliVersionFromWorkflow(workflowText);
} catch {
  // Best-effort: this script is sometimes run in ad-hoc contexts where the
  // release workflow file isn't present. In CI we expect it to exist.
  releaseWorkflowPinnedCliVersion = null;
}
if (!pinnedCliVersion) {
  pinnedCliVersion = releaseWorkflowPinnedCliVersion;
}

if (!pinnedCliVersion) {
  console.error(
    "Missing TAURI_CLI_VERSION. Provide it via env/CLI arg, or define it in .github/workflows/release.yml.",
  );
  process.exit(1);
}

if (
  pinnedSource !== "workflow" &&
  releaseWorkflowPinnedCliVersion &&
  releaseWorkflowPinnedCliVersion !== pinnedCliVersion
) {
  console.error("TAURI_CLI_VERSION mismatch detected.");
  console.error(`- ${releaseWorkflowRelativePath}: TAURI_CLI_VERSION="${releaseWorkflowPinnedCliVersion}"`);
  console.error(`- provided via ${pinnedSource}: TAURI_CLI_VERSION="${pinnedCliVersion}"`);
  console.error("");
  console.error("Fix:");
  console.error(`- Update the caller's TAURI_CLI_VERSION to match ${releaseWorkflowRelativePath}.`);
  process.exit(1);
}

const cliMajorMinor = parseMajorMinor(pinnedCliVersion);
try {
  // Fail fast if someone accidentally loosens the pin (e.g. "2.9") and
  // reintroduces toolchain drift.
  parsePinnedCliVersion(pinnedCliVersion);
} catch (err) {
  console.error(String(err instanceof Error ? err.message : err));
  process.exit(1);
}

if (cliMajorMinor !== tauriMajorMinor) {
  console.error("Pinned Tauri CLI version does not match the repo's Tauri major/minor version.");
  console.error(`- ${cargoTomlRelativePath}: tauri = "${tauriVersion}" (major/minor ${tauriMajorMinor})`);
  console.error(`- pinned TAURI_CLI_VERSION: "${pinnedCliVersion}" (major/minor ${cliMajorMinor})`);
  console.error("");
  console.error("Fix:");
  console.error(`- Bump TAURI_CLI_VERSION in ${releaseWorkflowRelativePath} to ${tauriMajorMinor}.x`);
  console.error("- Update docs/release.md to match (local release instructions).");
  process.exit(1);
}

console.log(
  `Pinned Tauri CLI version check passed: tauri=${tauriVersion} (major/minor ${tauriMajorMinor}) matches TAURI_CLI_VERSION=${pinnedCliVersion}.`,
);
