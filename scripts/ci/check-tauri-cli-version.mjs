import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import fs from "node:fs";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

const cargoTomlRelativePath = "apps/desktop/src-tauri/Cargo.toml";
const cargoTomlPath = path.join(repoRoot, cargoTomlRelativePath);

const releaseWorkflowRelativePath = ".github/workflows/release.yml";
const releaseWorkflowPath = path.join(repoRoot, releaseWorkflowRelativePath);

const docsReleaseRelativePath = "docs/release.md";
const docsReleasePath = path.join(repoRoot, docsReleaseRelativePath);

const cargoLockRelativePath = "Cargo.lock";
const cargoLockPath = path.join(repoRoot, cargoLockRelativePath);

const workflowsDirRelativePath = ".github/workflows";
const workflowsDirPath = path.join(repoRoot, workflowsDirRelativePath);

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
  for (const line of yamlLinesOutsideBlockScalars(workflowText)) {
    const match = line.match(/^[\t ]*TAURI_CLI_VERSION:[\t ]*["']?([^"'\n]+)["']?/);
    if (match) return match[1].trim();
  }
  return null;
}

export function extractPinnedCliVersionsFromWorkflow(workflowText) {
  const versions = [];
  for (const line of yamlLinesOutsideBlockScalars(workflowText)) {
    const match = line.match(/^[\t ]*TAURI_CLI_VERSION:[\t ]*["']?([^"'\n]+)["']?/);
    if (match) versions.push(match[1].trim());
  }
  return versions;
}

function yamlLinesOutsideBlockScalars(yamlText) {
  // GitHub Actions YAML uses block scalars (`run: |`, `script: >-`, etc) which can contain arbitrary
  // text. When scanning workflow configuration we must ignore the *contents* of block scalars so
  // YAML-ish strings inside scripts cannot satisfy or fail guardrails.
  //
  // We keep this YAML handling intentionally lightweight (no YAML parser dependency).
  const rawLines = stripHashComments(String(yamlText ?? "")).split(/\r?\n/);
  let inBlock = false;
  let blockIndent = 0;
  const blockRe = /:[\t ]*[>|][0-9+-]*[\t ]*$/;

  for (let i = 0; i < rawLines.length; i += 1) {
    const detect = rawLines[i] ?? "";

    const indentMatch = detect.match(/^[ \t]*/);
    const indentLen = indentMatch ? indentMatch[0].length : 0;

    if (inBlock) {
      // Blank lines can appear inside block scalars at any indentation; treat them as part of the scalar.
      if (detect.trim() === "") {
        rawLines[i] = "";
        continue;
      }
      if (indentLen > blockIndent) {
        rawLines[i] = "";
        continue;
      }
      inBlock = false;
    }

    const detectTrimmedEnd = detect.trimEnd();
    if (blockRe.test(detectTrimmedEnd)) {
      inBlock = true;
      blockIndent = indentLen;
    }
  }

  return rawLines;
}

export function findTauriActionScriptIssues(workflowText) {
  // Enforce that any workflow using tauri-apps/tauri-action runs it through the
  // Cargo-installed Tauri CLI (`cargo tauri`) so we don't drift to a floating
  // `@tauri-apps/cli@v2` toolchain.
  //
  // We intentionally keep this YAML parsing very lightweight/dependency-free.
  const issues = [];
  const lines = yamlLinesOutsideBlockScalars(workflowText);
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const trimmed = line.trimStart();
    if (trimmed.startsWith("#")) continue;

    const dashUses = line.match(/^(?<indent>\s*)-\s+uses:\s*tauri-apps\/tauri-action@/);
    const plainUses = line.match(/^(?<indent>\s*)uses:\s*tauri-apps\/tauri-action@/);
    if (!dashUses && !plainUses) continue;

    const indentLen = (dashUses?.groups?.indent ?? plainUses?.groups?.indent ?? "").length;
    const stepIndentLen = dashUses ? indentLen : Math.max(0, indentLen - 2);

    // Find the start of the step (the `- ...` line) so we can scan the entire
    // mapping for `tauriScript`, regardless of key ordering.
    let stepStartIndex = i;
    if (!dashUses) {
      for (let k = i; k >= 0; k -= 1) {
        const m = lines[k].match(/^(\s*)-\s+/);
        if (!m) continue;
        if ((m[1] ?? "").length === stepIndentLen) {
          stepStartIndex = k;
          break;
        }
      }
    }

    // Find the end of the step: the next `- ...` at the same or lower indent.
    let stepEndIndex = lines.length;
    for (let j = stepStartIndex + 1; j < lines.length; j += 1) {
      const nextLine = lines[j];
      const nextTrimmed = nextLine.trimStart();
      if (nextTrimmed.startsWith("#")) continue;
      const stepStart = nextLine.match(/^(\s*)-\s+/);
      if (stepStart && (stepStart[1] ?? "").length <= stepIndentLen) {
        stepEndIndex = j;
        break;
      }
    }

    let tauriScriptValue = null;
    let tauriScriptLine = null;
    for (let j = stepStartIndex; j < stepEndIndex; j += 1) {
      const l = lines[j];
      const lTrim = l.trimStart();
      if (lTrim.startsWith("#")) continue;
      const tsMatch = l.match(/^\s*tauriScript:\s*["']?([^"'\n#]+)["']?\s*(?:#.*)?$/);
      if (!tsMatch) continue;
      tauriScriptValue = tsMatch[1].trim();
      tauriScriptLine = j + 1;
      break;
    }

    if (!tauriScriptValue) {
      issues.push({
        line: stepStartIndex + 1,
        message: "tauri-action step must set tauriScript: cargo tauri",
      });
      continue;
    }

    if (tauriScriptValue !== "cargo tauri") {
      issues.push({
        line: tauriScriptLine ?? stepStartIndex + 1,
        message: `tauriScript must be \"cargo tauri\" (found ${JSON.stringify(tauriScriptValue)})`,
      });
    }
  }

  return issues;
}

function extractPinnedCliVersionsFromDocs(markdownText) {
  // Look for shell-style assignments like:
  //   TAURI_CLI_VERSION=2.9.5
  // (possibly indented, possibly with quotes).
  const versions = [];
  const re = /TAURI_CLI_VERSION\s*=\s*["']?([0-9]+\.[0-9]+\.[0-9]+(?:[-+][^\s"']+)?)["']?/g;
  for (const match of markdownText.matchAll(re)) {
    versions.push(match[1]);
  }
  return versions;
}

function extractTauriVersionFromCargoLock(lockText) {
  // Cargo.lock format:
  // [[package]]
  // name = "tauri"
  // version = "2.9.5"
  //
  // We only need the resolved version string; parsing with a regex keeps this script dependency-free.
  const match = lockText.match(
    /\[\[package\]\]\s*\r?\nname\s*=\s*"tauri"\s*\r?\nversion\s*=\s*"([^"]+)"/m,
  );
  return match ? match[1].trim() : null;
}

async function main() {
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

  if (releaseWorkflowPinnedCliVersion) {
    try {
      const entries = fs.readdirSync(workflowsDirPath, { withFileTypes: true });
      const mismatchLines = [];
      const tauriActionIssues = [];
      for (const entry of entries) {
        if (!entry.isFile()) continue;
        if (!entry.name.endsWith(".yml") && !entry.name.endsWith(".yaml")) continue;
        const rel = `${workflowsDirRelativePath}/${entry.name}`;
        const fullPath = path.join(workflowsDirPath, entry.name);
        const text = await readFile(fullPath, "utf8");
        const versions = extractPinnedCliVersionsFromWorkflow(text);
        for (const v of versions) {
          if (v !== releaseWorkflowPinnedCliVersion) {
            mismatchLines.push(`- ${rel}: TAURI_CLI_VERSION="${v}"`);
          }
        }

        const issues = findTauriActionScriptIssues(text);
        for (const issue of issues) {
          tauriActionIssues.push(`- ${rel}:${issue.line}: ${issue.message}`);
        }
      }
      if (mismatchLines.length > 0) {
        console.error("TAURI_CLI_VERSION mismatch across workflow files detected.");
        console.error(`- ${releaseWorkflowRelativePath}: TAURI_CLI_VERSION="${releaseWorkflowPinnedCliVersion}"`);
        mismatchLines.sort();
        for (const line of mismatchLines) console.error(line);
        console.error("");
        console.error("Fix:");
        console.error(
          `- Update all workflow TAURI_CLI_VERSION values to match ${releaseWorkflowRelativePath} (${releaseWorkflowPinnedCliVersion}).`,
        );
        process.exit(1);
      }

      if (tauriActionIssues.length > 0) {
        console.error("tauri-apps/tauri-action must use the pinned Cargo-installed Tauri CLI.");
        console.error("Set: tauriScript: cargo tauri");
        console.error("");
        tauriActionIssues.sort();
        for (const line of tauriActionIssues) console.error(line);
        process.exit(1);
      }
    } catch {
      // Ignore: best-effort scan. CI contexts should include these files.
    }
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

  // Keep the pinned CLI patch version aligned with the resolved Tauri crate version in Cargo.lock.
  // This prevents "tauri crates upgraded but CLI not upgraded" drift.
  try {
    const lockText = await readFile(cargoLockPath, "utf8");
    const tauriLockVersion = extractTauriVersionFromCargoLock(lockText);
    if (!tauriLockVersion) {
      console.error(`Failed to locate tauri version in ${cargoLockRelativePath}.`);
      process.exit(1);
    }
    const normalizedPinned = String(pinnedCliVersion).trim().replace(/^v/, "");
    if (normalizedPinned !== tauriLockVersion) {
      console.error("Pinned TAURI_CLI_VERSION does not match the resolved tauri crate version in Cargo.lock.");
      console.error(`- ${cargoLockRelativePath}: tauri version "${tauriLockVersion}"`);
      console.error(`- ${releaseWorkflowRelativePath}: TAURI_CLI_VERSION="${normalizedPinned}"`);
      console.error("");
      console.error("Fix:");
      console.error(`- Update TAURI_CLI_VERSION in ${releaseWorkflowRelativePath} to "${tauriLockVersion}".`);
      console.error(`- Update TAURI_CLI_VERSION in any other workflow/docs snippets to match.`);
      process.exit(1);
    }
  } catch (err) {
    console.error(`Failed to read/parse ${cargoLockRelativePath}.`);
    console.error(err);
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

  // Best-effort docs check: if the release docs pin a TAURI_CLI_VERSION, keep it in sync with the
  // canonical release workflow version so bumping is an explicit PR.
  try {
    const docsText = await readFile(docsReleasePath, "utf8");
    const docsVersions = extractPinnedCliVersionsFromDocs(docsText);
    const mismatches = [...new Set(docsVersions)].filter((v) => v !== pinnedCliVersion);
    if (mismatches.length > 0) {
      console.error("docs/release.md TAURI_CLI_VERSION mismatch detected.");
      console.error(`- ${releaseWorkflowRelativePath}: TAURI_CLI_VERSION="${pinnedCliVersion}"`);
      console.error(
        `- ${docsReleaseRelativePath}: found ${mismatches.length} mismatching value(s): ${mismatches.join(", ")}`,
      );
      console.error("");
      console.error("Fix:");
      console.error(
        `- Update the TAURI_CLI_VERSION assignment(s) in ${docsReleaseRelativePath} to ${pinnedCliVersion}.`,
      );
      process.exit(1);
    }
  } catch {
    // Ignore: the docs file may not be present in some ad-hoc contexts.
  }

  console.log(
    `Pinned Tauri CLI version check passed: tauri=${tauriVersion} (major/minor ${tauriMajorMinor}) matches TAURI_CLI_VERSION=${pinnedCliVersion}.`,
  );
}

const invokedAsScript = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedAsScript) {
  try {
    await main();
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
}
