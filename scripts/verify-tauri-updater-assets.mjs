#!/usr/bin/env node
/**
 * Verify that a GitHub Release contains a complete set of Tauri updater assets.
 *
 * This is meant to run as a post-build release guard in CI:
 * - download + parse `latest.json` from the release
 * - ensure each `platforms[<target>].url` references a release asset that exists
 * - ensure a signature is present for each updater asset (inline signature OR `<asset>.sig`)
 * - ensure human-install artifacts exist (DMG / EXE+MSI / AppImage+DEB+RPM)
 *
 * Usage:
 *   node scripts/verify-tauri-updater-assets.mjs <tag>
 *
 * Or:
 *   node scripts/verify-tauri-updater-assets.mjs --tag v0.1.0 --repo owner/repo
 *
 * Required env:
 *   - GITHUB_TOKEN (or GH_TOKEN)
 *   - GITHUB_REPOSITORY (if --repo not provided)
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
const tauriConfigPath = path.join(repoRoot, tauriConfigRelativePath);

const PLACEHOLDER_PUBKEY = "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY";

// Updater platform key mapping is intentionally strict.
//
// Source of truth:
// - docs/desktop-updater-target-mapping.md
//
// If Tauri/tauri-action changes the platform key naming scheme (or which artifact is used for
// self-update), this script should fail loudly so we update the mapping doc + verification logic
// together.
const EXPECTED_PLATFORMS = [
  {
    key: "darwin-universal",
    label: "macOS (universal)",
    expectedUpdaterAsset: {
      description: "macOS updater archive (*.app.tar.gz)",
      matches: (name) => name.endsWith(".app.tar.gz"),
    },
  },
  {
    key: "windows-x86_64",
    label: "Windows (x64)",
    expectedUpdaterAsset: {
      description: "Windows updater installer (*.msi)",
      matches: (name) => name.toLowerCase().endsWith(".msi"),
    },
  },
  {
    key: "windows-aarch64",
    label: "Windows (ARM64)",
    expectedUpdaterAsset: {
      description: "Windows updater installer (*.msi)",
      matches: (name) => name.toLowerCase().endsWith(".msi"),
    },
  },
  {
    key: "linux-x86_64",
    label: "Linux (x86_64)",
    expectedUpdaterAsset: {
      description: "Linux updater bundle (*.AppImage)",
      matches: (name) => name.endsWith(".AppImage"),
    },
  },
];
const EXPECTED_PLATFORM_KEYS = EXPECTED_PLATFORMS.map((p) => p.key);

/**
 * @param {string} message
 */
function fail(message) {
  console.error(message);
  process.exitCode = 1;
}

/**
 * @param {string} heading
 * @param {string[]} details
 */
function failBlock(heading, details) {
  fail(`\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}`);
}

/**
 * @param {string} s
 */
function normalizeTagVersion(s) {
  const trimmed = s.trim();
  return trimmed.startsWith("v") ? trimmed.slice(1) : trimmed;
}

/**
 * @param {unknown} value
 * @returns {value is Record<string, any>}
 */
function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {string} urlOrPath
 */
function filenameFromUrl(urlOrPath) {
  const stripQuery = (s) => s.split("#")[0].split("?")[0];
  try {
    const url = new URL(urlOrPath);
    const raw = url.pathname.split("/").pop() || "";
    const decoded = decodeURIComponent(raw);
    return stripQuery(decoded);
  } catch {
    const raw = stripQuery(urlOrPath).split("/").pop() || "";
    try {
      return decodeURIComponent(raw);
    } catch {
      return raw;
    }
  }
}

/**
 * @param {string} arch
 * @param {string} assetName
 */
function archMatchesAssetName(arch, assetName) {
  const name = assetName.toLowerCase();
  const a = arch.toLowerCase();

  // Put the more specific checks first so we don't match `x86` inside `x86_64`.
  if (a === "x86_64" || a === "amd64") {
    return (
      name.includes("x86_64") ||
      name.includes("amd64") ||
      // common Windows output naming
      name.includes("x64") ||
      // common Tauri target triple string
      name.includes("win64")
    );
  }
  if (a === "aarch64" || a === "arm64") {
    return name.includes("aarch64") || name.includes("arm64");
  }
  if (a === "i686" || a === "x86") {
    return name.includes("i686") || name.includes("ia32") || name.includes("win32") || name.includes("x86");
  }

  // Fallback: try a raw substring match (works for rare arch names).
  return name.includes(a);
}

/**
 * @param {string} baseUrl
 * @param {string} apiPath
 * @param {string} token
 */
async function githubApiJson(baseUrl, apiPath, token) {
  const url = `${baseUrl}${apiPath}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`GitHub API request failed: ${res.status} ${res.statusText} (${url})\n${text}`);
  }
  return res.json();
}

/**
 * Download a release asset by asset id.
 *
 * This works for both published and draft releases (the API grants access to a
 * pre-signed redirect URL even when the asset is not public).
 *
 * @param {string} baseUrl
 * @param {string} repo
 * @param {number} assetId
 * @param {string} token
 */
async function downloadReleaseAsset(baseUrl, repo, assetId, token) {
  const url = `${baseUrl}/repos/${repo}/releases/assets/${assetId}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/octet-stream",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    redirect: "follow",
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to download release asset ${assetId}: ${res.status} ${res.statusText}\n${text}`);
  }
  const buf = Buffer.from(await res.arrayBuffer());
  return buf;
}

/**
 * @param {string} baseUrl
 * @param {string} repo
 * @param {number} releaseId
 * @param {string} token
 */
async function listReleaseAssets(baseUrl, repo, releaseId, token) {
  /** @type {any[]} */
  const assets = [];
  for (let page = 1; page <= 20; page += 1) {
    const pageAssets = await githubApiJson(
      baseUrl,
      `/repos/${repo}/releases/${releaseId}/assets?per_page=100&page=${page}`,
      token,
    );
    if (!Array.isArray(pageAssets)) {
      throw new Error(`Unexpected assets payload for release ${releaseId}: expected array.`);
    }
    assets.push(...pageAssets);
    if (pageAssets.length < 100) {
      break;
    }
  }
  return assets;
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {{tag?: string, repo?: string}} */
  const out = {};

  /** @type {string[]} */
  const positional = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;
    if (!arg.startsWith("-")) {
      positional.push(arg);
      continue;
    }

    const parseValue = () => {
      const next = argv[i + 1];
      if (!next || next.startsWith("-")) {
        throw new Error(`Missing value for ${arg}`);
      }
      i += 1;
      return next;
    };

    if (arg === "--tag") {
      out.tag = parseValue();
    } else if (arg.startsWith("--tag=")) {
      out.tag = arg.slice("--tag=".length);
    } else if (arg === "--repo") {
      out.repo = parseValue();
    } else if (arg.startsWith("--repo=")) {
      out.repo = arg.slice("--repo=".length);
    } else if (arg === "--help" || arg === "-h") {
      out.help = true;
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }

  if (!out.tag && positional.length > 0) {
    out.tag = positional[0];
  }

  return out;
}

async function readUpdaterConfigExpectation() {
  /** @type {{updaterActive: boolean, expectsManifestSignature: boolean}} */
  const result = { updaterActive: false, expectsManifestSignature: false };
  try {
    const text = await readFile(tauriConfigPath, "utf8");
    /** @type {any} */
    const config = JSON.parse(text);
    const updater = config?.plugins?.updater;
    const active = updater?.active === true;
    result.updaterActive = active;
    if (!active) {
      return result;
    }
    const pubkey = typeof updater?.pubkey === "string" ? updater.pubkey.trim() : "";
    if (pubkey && pubkey !== PLACEHOLDER_PUBKEY) {
      result.expectsManifestSignature = true;
    }
  } catch {
    // If we can't read config, keep the default "unknown" expectations; we still
    // validate whatever assets exist in the release.
  }
  return result;
}

/**
 * @param {any} entry
 * @returns {string[]}
 */
function extractUrls(entry) {
  const candidates = [];
  const direct =
    entry?.url ??
    entry?.downloadUrl ??
    entry?.download_url ??
    entry?.downloadURL ??
    entry?.download;
  if (typeof direct === "string") {
    candidates.push(direct);
  } else if (Array.isArray(direct)) {
    candidates.push(...direct.filter((v) => typeof v === "string"));
  }

  const alt = entry?.urls ?? entry?.downloadUrls ?? entry?.download_urls;
  if (typeof alt === "string") {
    candidates.push(alt);
  } else if (Array.isArray(alt)) {
    candidates.push(...alt.filter((v) => typeof v === "string"));
  }

  // De-dupe while preserving order.
  /** @type {string[]} */
  const uniq = [];
  const seen = new Set();
  for (const c of candidates) {
    const trimmed = c.trim();
    if (!trimmed) continue;
    if (seen.has(trimmed)) continue;
    seen.add(trimmed);
    uniq.push(trimmed);
  }
  return uniq;
}

/**
 * @param {any} entry
 */
function hasInlineSignature(entry) {
  const value = entry?.signature ?? entry?.sig ?? entry?.signed ?? entry?.sign;
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * @param {string} target
 * @returns {"macos" | "linux" | "windows" | "other"}
 */
function platformFamilyFromTarget(target) {
  const lower = target.toLowerCase();
  if (lower.includes("darwin") || lower.includes("apple-darwin") || lower.includes("macos")) return "macos";
  if (lower.includes("linux")) return "linux";
  if (lower.includes("windows") || lower.includes("pc-windows")) return "windows";
  return "other";
}

/**
 * @param {string[]} platformKeys
 */
function groupPlatformKeys(platformKeys) {
  /** @type {{windows: string[], linux: string[], macos: string[], other: string[]}} */
  const groups = { windows: [], linux: [], macos: [], other: [] };
  for (const key of platformKeys) {
    const family = platformFamilyFromTarget(key);
    groups[family].push(key);
  }
  return groups;
}

/**
 * @param {string} target
 */
function inferArchFromTarget(target) {
  const lower = target.toLowerCase();
  if (lower.includes("x86_64") || lower.includes("amd64")) return "x86_64";
  if (lower.includes("aarch64")) return "aarch64";
  if (lower.includes("arm64")) return "arm64";
  if (lower.includes("i686")) return "i686";
  if (lower.includes("universal")) return "universal";
  return "unknown";
}

async function main() {
  let args;
  try {
    args = parseArgs(process.argv.slice(2));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    fail(`verify-tauri-updater-assets: ${msg}`);
    fail(`Usage: node scripts/verify-tauri-updater-assets.mjs <tag> [--repo owner/repo]`);
    process.exit(1);
  }

  if (args.help) {
    console.log(`Usage: node scripts/verify-tauri-updater-assets.mjs <tag> [--repo owner/repo]`);
    process.exit(0);
  }

  const tag = args.tag ?? process.env.GITHUB_REF_NAME;
  if (!tag) {
    fail(`Missing tag name. Usage: node scripts/verify-tauri-updater-assets.mjs <tag>`);
    process.exit(1);
  }

  const repo = args.repo ?? process.env.GITHUB_REPOSITORY;
  if (!repo) {
    failBlock(`Missing repo name`, [
      `Set GITHUB_REPOSITORY in the environment (e.g. "owner/repo"), or pass --repo owner/repo.`,
    ]);
    process.exit(1);
  }

  const token = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!token) {
    failBlock(`Missing GitHub token`, [
      `Set GITHUB_TOKEN (recommended for GitHub Actions) or GH_TOKEN (for local runs).`,
    ]);
    process.exit(1);
  }

  const apiBase = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/$/, "");

  /** @type {any} */
  const release = await githubApiJson(apiBase, `/repos/${repo}/releases/tags/${encodeURIComponent(tag)}`, token);
  const releaseId = release?.id;
  if (typeof releaseId !== "number") {
    fail(`Unexpected GitHub API response: missing release id for tag ${tag}.`);
    process.exit(1);
  }

  const assets = await listReleaseAssets(apiBase, repo, releaseId, token);
  /** @type {Map<string, any>} */
  const assetByName = new Map();
  for (const asset of assets) {
    if (asset && typeof asset.name === "string") {
      assetByName.set(asset.name, asset);
    }
  }

  console.log(
    `Release asset verification: ${repo}@${tag} (draft=${Boolean(release?.draft)}, assets=${assets.length})`,
  );

  const manifestAsset = assetByName.get("latest.json");
  if (!manifestAsset) {
    failBlock(`Missing updater manifest`, [
      `Release does not contain "latest.json".`,
      `This asset is expected to be uploaded by tauri-apps/tauri-action.`,
    ]);
    process.exit(1);
  }

  const updaterExpectation = await readUpdaterConfigExpectation();
  const wantsManifestSig = updaterExpectation.expectsManifestSignature;
  const manifestSigAsset = assetByName.get("latest.json.sig");

  if (wantsManifestSig && !manifestSigAsset) {
    failBlock(`Missing updater manifest signature`, [
      `apps/desktop/src-tauri/tauri.conf.json has plugins.updater.active=true and a non-placeholder pubkey,`,
      `but the release is missing "latest.json.sig".`,
    ]);
  }

  // Download latest.json (and signature if present) from the release.
  /** @type {Buffer} */
  let manifestBuf;
  try {
    manifestBuf = await downloadReleaseAsset(apiBase, repo, manifestAsset.id, token);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Failed to download updater manifest`, [`Asset: latest.json`, `Error: ${msg}`]);
    process.exit(1);
  }

  if (manifestSigAsset) {
    try {
      await downloadReleaseAsset(apiBase, repo, manifestSigAsset.id, token);
      console.log(`- Found updater manifest signature: latest.json.sig`);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      failBlock(`Failed to download updater manifest signature`, [`Asset: latest.json.sig`, `Error: ${msg}`]);
    }
  } else {
    console.log(`- No updater manifest signature: latest.json.sig (not found)`);
  }

  /** @type {any} */
  let manifest;
  try {
    manifest = JSON.parse(manifestBuf.toString("utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Invalid latest.json`, [`Failed to parse JSON.`, `Error: ${msg}`]);
    process.exit(1);
  }

  const manifestVersion = typeof manifest?.version === "string" ? manifest.version.trim() : "";
  if (manifestVersion) {
    const tagVersion = normalizeTagVersion(tag);
    if (manifestVersion !== tagVersion) {
      failBlock(`Updater manifest version mismatch`, [
        `Tag: ${tag} (version ${JSON.stringify(tagVersion)})`,
        `latest.json: version ${JSON.stringify(manifestVersion)}`,
      ]);
    }
  }

  const platforms = manifest?.platforms;
  if (!isPlainObject(platforms)) {
    failBlock(`Invalid latest.json`, [
      `Expected latest.json to contain a "platforms" object mapping target → { url, signature, ... }.`,
    ]);
    process.exit(1);
  }

  const platformKeys = Object.keys(platforms).sort();
  if (platformKeys.length === 0) {
    failBlock(`Invalid latest.json`, [`"platforms" is empty; updater has no targets.`]);
    process.exit(1);
  }

  // Enforce stable updater platform key naming (see docs/desktop-updater-target-mapping.md).
  const expectedKeysSorted = EXPECTED_PLATFORM_KEYS.slice().sort();
  const expectedKeySet = new Set(EXPECTED_PLATFORM_KEYS);
  const missingKeys = expectedKeysSorted.filter((k) => !Object.prototype.hasOwnProperty.call(platforms, k));
  const unexpectedKeys = platformKeys.filter((k) => !expectedKeySet.has(k));
  if (missingKeys.length > 0 || unexpectedKeys.length > 0) {
    const formatKeyList = (keys) => keys.map((k) => `    - ${k}`).join("\n");
    failBlock(`Unexpected latest.json.platforms keys (Tauri updater target identifiers)`, [
      `Expected (${expectedKeysSorted.length}):\n${formatKeyList(expectedKeysSorted)}`,
      `Actual (${platformKeys.length}):\n${formatKeyList(platformKeys)}`,
      ...(missingKeys.length > 0 ? [`Missing (${missingKeys.length}):\n${formatKeyList(missingKeys)}`] : []),
      ...(unexpectedKeys.length > 0
        ? [`Unexpected (${unexpectedKeys.length}):\n${formatKeyList(unexpectedKeys)}`]
        : []),
      `If you upgraded Tauri/tauri-action, update docs/desktop-updater-target-mapping.md and scripts/verify-tauri-updater-assets.mjs together.`,
    ]);
  }

  const groups = groupPlatformKeys(platformKeys);

  console.log(`- Platforms in latest.json: ${platformKeys.join(", ")}`);

  /** @type {{platform: string, url: string, assetName: string, inlineSig: boolean, sigAssetName: string, sigAssetPresent: boolean}[]} */
  const updaterEntries = [];

  /** @type {string[]} */
  const missing = [];

  for (const platformKey of platformKeys) {
    const entry = platforms[platformKey];
    const urls = extractUrls(entry);
    if (urls.length === 0) {
      missing.push(`latest.json platforms[${platformKey}] has no url field`);
      continue;
    }

    const inlineSig = hasInlineSignature(entry);
    const expectedAssetType = EXPECTED_PLATFORMS.find((p) => p.key === platformKey)?.expectedUpdaterAsset ?? null;

    for (const url of urls) {
      const assetName = filenameFromUrl(url);
      if (!assetName) {
        missing.push(`latest.json platforms[${platformKey}] has an unparseable url: ${JSON.stringify(url)}`);
        continue;
      }

      if (expectedAssetType && !expectedAssetType.matches(assetName)) {
        missing.push(
          `Updater asset type mismatch for ${platformKey}: ${assetName} (expected ${expectedAssetType.description})`,
        );
      }

      const assetExists = assetByName.has(assetName);
      if (!assetExists) {
        missing.push(`Missing release asset referenced by latest.json: ${platformKey} → ${assetName}`);
      }

      const sigAssetName = `${assetName}.sig`;
      const sigAssetPresent = assetByName.has(sigAssetName);
      if (!inlineSig && !sigAssetPresent) {
        missing.push(
          `Missing updater signature for ${platformKey} → ${assetName} (need inline "signature" field or ${sigAssetName})`,
        );
      }

      updaterEntries.push({
        platform: platformKey,
        url,
        assetName,
        inlineSig,
        sigAssetName,
        sigAssetPresent,
      });
    }
  }

  // Print a concise per-platform summary.
  for (const entry of updaterEntries) {
    const sigStatus = entry.inlineSig ? "inline" : entry.sigAssetPresent ? entry.sigAssetName : "MISSING";
    console.log(`  - ${entry.platform}: ${entry.assetName} (sig: ${sigStatus})`);
  }

  // Verify human-install artifacts exist.
  const assetNames = [...assetByName.keys()];

  const hasAny = (suffix) => assetNames.some((n) => n.toLowerCase().endsWith(suffix.toLowerCase()));

  // macOS: require at least one DMG.
  if (!hasAny(".dmg")) {
    missing.push(`Missing macOS installer: no .dmg asset found in the release`);
  }

  // Windows: require EXE + MSI per arch (arch list derived from updater platforms).
  const windowsArches = groups.windows.map(inferArchFromTarget);
  const uniqueWindowsArches = [...new Set(windowsArches)];
  const windowsExeAssets = assetNames.filter((n) => n.toLowerCase().endsWith(".exe"));
  const windowsMsiAssets = assetNames.filter((n) => n.toLowerCase().endsWith(".msi"));

  if (uniqueWindowsArches.length === 1) {
    if (windowsExeAssets.length === 0) missing.push(`Missing Windows installer: no .exe asset found`);
    if (windowsMsiAssets.length === 0) missing.push(`Missing Windows installer: no .msi asset found`);
  } else {
    if (uniqueWindowsArches.includes("unknown")) {
      missing.push(
        `Cannot infer Windows arch names from latest.json platform keys: ${groups.windows
          .map((k) => JSON.stringify(k))
          .join(", ")}`,
      );
    }
    for (const arch of uniqueWindowsArches) {
      if (arch === "unknown") continue;
      const exeForArch = windowsExeAssets.filter((n) => archMatchesAssetName(arch, n));
      const msiForArch = windowsMsiAssets.filter((n) => archMatchesAssetName(arch, n));
      if (exeForArch.length === 0) missing.push(`Missing Windows .exe installer for arch ${arch}`);
      if (msiForArch.length === 0) missing.push(`Missing Windows .msi installer for arch ${arch}`);
    }
  }

  // Linux: require AppImage + DEB + RPM (per arch if multiple linux arches exist).
  const linuxArches = groups.linux.map(inferArchFromTarget);
  const uniqueLinuxArches = [...new Set(linuxArches)];
  const linuxAppImages = assetNames.filter((n) => n.endsWith(".AppImage"));
  const linuxDebs = assetNames.filter((n) => n.toLowerCase().endsWith(".deb"));
  const linuxRpms = assetNames.filter((n) => n.toLowerCase().endsWith(".rpm"));

  if (uniqueLinuxArches.length === 1) {
    if (linuxAppImages.length === 0) missing.push(`Missing Linux bundle: no .AppImage asset found`);
    if (linuxDebs.length === 0) missing.push(`Missing Linux bundle: no .deb asset found`);
    if (linuxRpms.length === 0) missing.push(`Missing Linux bundle: no .rpm asset found`);
  } else {
    if (uniqueLinuxArches.includes("unknown")) {
      missing.push(
        `Cannot infer Linux arch names from latest.json platform keys: ${groups.linux
          .map((k) => JSON.stringify(k))
          .join(", ")}`,
      );
    }
    for (const arch of uniqueLinuxArches) {
      if (arch === "unknown") continue;
      const appForArch = linuxAppImages.filter((n) => archMatchesAssetName(arch, n));
      const debForArch = linuxDebs.filter((n) => archMatchesAssetName(arch, n));
      const rpmForArch = linuxRpms.filter((n) => archMatchesAssetName(arch, n));
      if (appForArch.length === 0) missing.push(`Missing Linux .AppImage bundle for arch ${arch}`);
      if (debForArch.length === 0) missing.push(`Missing Linux .deb bundle for arch ${arch}`);
      if (rpmForArch.length === 0) missing.push(`Missing Linux .rpm bundle for arch ${arch}`);
    }
  }

  if (missing.length > 0) {
    const unique = [...new Set(missing)];
    failBlock(`Release asset verification failed`, unique);
    process.exit(1);
  }

  // Some validations above call failBlock() without immediately exiting so we can
  // report multiple issues in a single run. If any were triggered, ensure the job
  // still fails (and avoid printing a misleading "passed" message).
  if (process.exitCode) {
    process.exit(1);
  }

  console.log(`Release asset verification passed.`);
}

main().catch((err) => {
  const msg = err instanceof Error ? err.stack || err.message : String(err);
  fail(`verify-tauri-updater-assets: ${msg}`);
  process.exit(1);
});
