#!/usr/bin/env node
/**
 * Verifies that a Tauri `latest.json` updater manifest contains entries for every
 * expected OS/architecture target produced by the desktop release workflow.
 *
 * This file intentionally lives under `scripts/` (vs `scripts/ci/`) so it is easy
 * to discover and matches docs/release.md references.
 *
 * Modes:
 *
 * 1) CI (default): `node scripts/verify-tauri-latest-json.mjs <tag>`
 *    Delegates to the more comprehensive validator `scripts/ci/validate-updater-manifest.mjs`,
 *    which downloads latest.json/latest.json.sig from the draft release and validates targets,
 *    signatures, and referenced assets.
 *
 * 2) Local file validation: `node scripts/verify-tauri-latest-json.mjs --manifest latest.json --sig latest.json.sig`
 *    Parses a local manifest and checks that the platforms map contains the required targets.
 */

import { spawnSync } from "node:child_process";
import { access, readFile, stat } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { constants as fsConstants } from "node:fs";
import { validatePlatformEntries } from "./ci/validate-updater-manifest.mjs";

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown>}
 */
function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {unknown} root
 * @returns {{ platforms: Record<string, unknown>; path: string[] } | null}
 */
function findPlatformsObject(root) {
  if (!root || (typeof root !== "object" && !Array.isArray(root))) {
    return null;
  }

  /** @type {{ value: unknown; path: string[] }[]} */
  const queue = [{ value: root, path: [] }];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) break;

    const { value, path: currentPath } = current;
    if (isPlainObject(value)) {
      if (isPlainObject(value.platforms)) {
        return { platforms: value.platforms, path: [...currentPath, "platforms"] };
      }
      for (const [key, child] of Object.entries(value)) {
        if (currentPath.length >= 8) continue;
        if (isPlainObject(child) || Array.isArray(child)) {
          queue.push({ value: child, path: [...currentPath, key] });
        }
      }
      continue;
    }

    if (Array.isArray(value)) {
      if (currentPath.length >= 8) continue;
      for (let i = 0; i < value.length; i += 1) {
        const child = value[i];
        if (isPlainObject(child) || Array.isArray(child)) {
          queue.push({ value: child, path: [...currentPath, String(i)] });
        }
      }
    }
  }

  return null;
}

function usage() {
  return `Usage:
  # CI (downloads from draft GitHub Release via scripts/ci/validate-updater-manifest.mjs)
  node scripts/verify-tauri-latest-json.mjs <tag>

  # Local file validation (no GitHub API calls)
  node scripts/verify-tauri-latest-json.mjs --manifest latest.json --sig latest.json.sig
`;
}

/**
 * @param {string[]} argv
 */
function hasLocalFlags(argv) {
  return argv.some((arg) => arg === "--manifest" || arg.startsWith("--manifest=") || arg === "--sig" || arg.startsWith("--sig=") || arg === "--signature" || arg.startsWith("--signature="));
}

/**
 * @param {string[]} argv
 */
async function runLocal(argv) {
  let manifestPath = "latest.json";
  let sigPath = "latest.json.sig";

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;
    if (arg === "--help" || arg === "-h") {
      console.log(usage());
      return 0;
    }
    if (arg === "--manifest") {
      manifestPath = argv[++i] ?? "";
      continue;
    }
    if (arg.startsWith("--manifest=")) {
      manifestPath = arg.slice("--manifest=".length);
      continue;
    }
    if (arg === "--sig" || arg === "--signature") {
      sigPath = argv[++i] ?? "";
      continue;
    }
    if (arg.startsWith("--sig=")) {
      sigPath = arg.slice("--sig=".length);
      continue;
    }
    if (arg.startsWith("--signature=")) {
      sigPath = arg.slice("--signature=".length);
      continue;
    }
  }

  if (!manifestPath || !sigPath) {
    console.error(usage());
    return 2;
  }

  try {
    await access(sigPath, fsConstants.R_OK);
    const stats = await stat(sigPath);
    if (stats.size === 0) {
      console.error(`Updater manifest verification failed: signature file is empty (${sigPath})`);
      return 1;
    }
  } catch (err) {
    console.error(`Updater manifest verification failed: missing signature file (${sigPath})`);
    console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
    return 1;
  }

  /** @type {any} */
  let manifest;
  try {
    manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  } catch (err) {
    console.error(`Updater manifest verification failed: could not read/parse ${manifestPath}`);
    console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
    return 1;
  }

  const found = findPlatformsObject(manifest);
  if (!found) {
    console.error(
      `Updater manifest verification failed: could not locate a \"platforms\" object in ${manifestPath}`,
    );
    return 1;
  }

  const platformKeys = Object.keys(found.platforms).sort();
  if (platformKeys.length === 0) {
    console.error(`Updater manifest verification failed: platforms map is empty (${manifestPath})`);
    return 1;
  }

  console.log(`Tauri updater manifest platforms found at: ${found.path.join(".")}`);
  console.log(`Platform keys (${platformKeys.length}): ${platformKeys.join(", ")}`);

  /**
   * Offline validation cannot confirm that referenced assets exist on the GitHub Release, but we
   * can still reuse the strict key + per-platform updater artifact type checks from the CI
   * validator by constructing a synthetic asset-name set from the URLs.
   *
   * This keeps `--manifest/--sig` behavior aligned with CI (same expected platform key names).
   */
  const assetNames = new Set();
  for (const entry of Object.values(found.platforms)) {
    if (!entry || typeof entry !== "object") continue;
    const url = /** @type {any} */ (entry).url;
    if (typeof url !== "string" || url.trim().length === 0) continue;
    try {
      const parsed = new URL(url);
      const last = parsed.pathname.split("/").filter(Boolean).pop() ?? "";
      const decoded = decodeURIComponent(last);
      if (decoded) assetNames.add(decoded);
    } catch {
      // Ignore; the validator will report invalid URLs.
    }
  }

  const { errors, invalidTargets, missingAssets } = validatePlatformEntries({
    platforms: found.platforms,
    assetNames,
  });

  // missingAssets should normally be empty (we seed assetNames from the URLs); treat it as an
  // internal consistency failure rather than a user-facing “asset missing” issue.
  if (missingAssets.length > 0) {
    errors.push(
      `Internal error: expected offline assetNames set to cover all platforms, but ${missingAssets.length} entry(s) were missing.`,
    );
  }

  if (invalidTargets.length > 0) {
    errors.push(
      [
        `Invalid platform entries in latest.json:`,
        ...invalidTargets.map((t) => `  - ${t.target}: ${t.message}`),
      ].join("\n"),
    );
  }

  if (errors.length > 0) {
    console.error("");
    console.error(`Updater manifest verification failed:`);
    for (const err of errors) {
      console.error(`- ${err}`);
      console.error("");
    }
    return 1;
  }

  console.log("Updater manifest verification passed.");
  return 0;
}

/**
 * @param {string[]} argv
 */
function runCi(argv) {
  const scriptDir = path.dirname(fileURLToPath(import.meta.url));
  const ciValidatorPath = path.join(scriptDir, "ci", "validate-updater-manifest.mjs");

  const res = spawnSync(process.execPath, [ciValidatorPath, ...argv], { stdio: "inherit" });
  if (res.error) {
    console.error(
      `Failed to run updater manifest validator (${ciValidatorPath}): ${res.error.message}`,
    );
    return 1;
  }
  return res.status ?? 1;
}

const argv = process.argv.slice(2);
if (argv.length === 0) {
  console.error(usage());
  process.exit(2);
}

if (argv.includes("--help") || argv.includes("-h")) {
  console.log(usage());
  process.exit(0);
}

if (hasLocalFlags(argv)) {
  process.exit(await runLocal(argv));
}

process.exit(runCi(argv));
