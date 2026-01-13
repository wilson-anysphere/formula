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

/**
 * Map a Tauri updater "target" key into a canonical {os}-{arch} identifier.
 * Handles common variations between Tauri v1/v2 and Rust target triples.
 *
 * @param {string} key
 * @returns {string | null}
 */
function canonicalizeTargetKey(key) {
  const raw = key.trim();
  if (!raw) return null;
  const lower = raw.toLowerCase();

  let os = null;
  if (
    lower.includes("universal-apple-darwin") ||
    lower.includes("apple-darwin") ||
    lower.includes("darwin") ||
    lower.includes("macos") ||
    lower.includes("osx")
  ) {
    os = "darwin";
  } else if (lower.includes("pc-windows") || lower.includes("windows")) {
    os = "windows";
  } else if (lower.includes("unknown-linux") || lower.includes("linux")) {
    os = "linux";
  }

  if (!os) return null;

  if (lower.includes("universal")) return `${os}-universal`;
  if (lower.includes("aarch64") || lower.includes("arm64")) return `${os}-aarch64`;
  if (lower.includes("x86_64") || lower.includes("amd64") || /\bx64\b/.test(lower)) return `${os}-x86_64`;

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

  /** @type {Map<string, Set<string>>} */
  const canonicalToRaw = new Map();
  for (const key of platformKeys) {
    const canonical = canonicalizeTargetKey(key);
    if (!canonical) continue;
    const set = canonicalToRaw.get(canonical) ?? new Set();
    set.add(key);
    canonicalToRaw.set(canonical, set);
  }

  const missing = [];

  const hasDarwinUniversal = canonicalToRaw.has("darwin-universal");
  const hasDarwinArm64 = canonicalToRaw.has("darwin-aarch64");
  const hasDarwinX64 = canonicalToRaw.has("darwin-x86_64");
  if (!(hasDarwinUniversal || (hasDarwinArm64 && hasDarwinX64))) {
    missing.push("macOS target: darwin-universal (or both darwin-aarch64 + darwin-x86_64)");
  }

  if (!canonicalToRaw.has("windows-x86_64")) missing.push("windows-x86_64");
  if (!canonicalToRaw.has("windows-aarch64")) missing.push("windows-aarch64");
  if (!canonicalToRaw.has("linux-x86_64")) missing.push("linux-x86_64");

  console.log(`Tauri updater manifest platforms found at: ${found.path.join(".")}`);
  console.log(`Platform keys (${platformKeys.length}): ${platformKeys.join(", ")}`);

  if (missing.length > 0) {
    console.error("");
    console.error(`Updater manifest verification failed: missing required platform entries:`);
    for (const m of missing) console.error(`  - ${m}`);
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

if (hasLocalFlags(argv)) {
  process.exit(await runLocal(argv));
}

process.exit(runCi(argv));
