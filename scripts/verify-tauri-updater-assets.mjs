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
import { setTimeout as sleep } from "node:timers/promises";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
const tauriConfigPath = path.join(repoRoot, tauriConfigRelativePath);

const PLACEHOLDER_PUBKEY = "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY";

/**
 * macOS updater artifacts are tarballs (often `*.app.tar.gz`, but allow other `*.tar.gz`/`*.tgz`
 * archives). Guard against accidentally matching Linux `*.AppImage.tar.gz` bundles.
 *
 * Keep this in sync with:
 * - scripts/ci/validate-updater-manifest.mjs
 * - scripts/verify-desktop-release-assets.mjs
 *
 * @param {string} name
 * @returns {boolean}
 */
export function isMacUpdaterArchiveAssetName(name) {
  const lower = name.toLowerCase();
  if (lower.endsWith(".appimage.tar.gz") || lower.endsWith(".appimage.tgz")) return false;
  return lower.endsWith(".tar.gz") || lower.endsWith(".tgz");
}

// Updater platform key mapping is intentionally strict.
//
// Source of truth:
// - docs/desktop-updater-target-mapping.md
//
// If Tauri/tauri-action changes the platform key naming scheme (or which artifact is used for
// self-update), this script should fail loudly so we update the mapping doc + verification logic
// together.
const EXPECTED_PLATFORMS = [
  // macOS universal builds are represented under both arch keys (same archive URL).
  {
    key: "darwin-x86_64",
    label: "macOS (x86_64)",
    expectedUpdaterAsset: {
      description: "macOS updater archive (*.app.tar.gz preferred; allow *.tar.gz/*.tgz)",
      matches: isMacUpdaterArchiveAssetName,
    },
  },
  {
    key: "darwin-aarch64",
    label: "macOS (aarch64)",
    expectedUpdaterAsset: {
      description: "macOS updater archive (*.app.tar.gz preferred; allow *.tar.gz/*.tgz)",
      matches: isMacUpdaterArchiveAssetName,
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
  {
    key: "linux-aarch64",
    label: "Linux (ARM64)",
    expectedUpdaterAsset: {
      description: "Linux updater bundle (*.AppImage)",
      matches: (name) => name.endsWith(".AppImage"),
    },
  },
];
const EXPECTED_PLATFORM_KEYS = EXPECTED_PLATFORMS.map((p) => p.key);

/**
 * @param {string} heading
 * @param {string[]} details
 */
function formatBlock(heading, details) {
  return `\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}`;
}

class VerificationFailure extends Error {
  /**
   * @param {{ blocks: Array<{ heading: string; details: string[] }>; retryable: boolean; logLines?: string[] }} args
   */
  constructor({ blocks, retryable, logLines }) {
    super("Release asset verification failed");
    this.blocks = blocks;
    this.retryable = retryable;
    this.logLines = logLines ?? [];
  }

  /**
   * @param {NodeJS.WritableStream} stream
   */
  print(stream) {
    if (this.logLines.length > 0) {
      for (const line of this.logLines) {
        stream.write(`${line}\n`);
      }
    }
    for (const block of this.blocks) {
      stream.write(`${formatBlock(block.heading, block.details)}\n`);
    }
  }
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
 * Best-effort locate a Tauri updater `platforms` map within a JSON payload.
 *
 * Tauri v1/v2 manifests typically have `{ platforms: { ... } }` at the top-level, but we keep this
 * robust in case future versions nest the object.
 *
 * @param {unknown} root
 * @returns {{ platforms: Record<string, any>; path: string[] } | null}
 */
export function findPlatformsObject(root) {
  if (!root || (typeof root !== "object" && !Array.isArray(root))) return null;

  /** @type {{ value: unknown; path: string[] }[]} */
  const queue = [{ value: root, path: [] }];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) break;
    const { value, path: currentPath } = current;

    if (isPlainObject(value)) {
      const platforms = /** @type {any} */ (value).platforms;
      if (isPlainObject(platforms)) {
        return { platforms, path: [...currentPath, "platforms"] };
      }

      if (currentPath.length >= 8) continue;
      for (const [key, child] of Object.entries(value)) {
        if (child && (isPlainObject(child) || Array.isArray(child))) {
          queue.push({ value: child, path: [...currentPath, key] });
        }
      }
      continue;
    }

    if (Array.isArray(value)) {
      if (currentPath.length >= 8) continue;
      for (let i = 0; i < value.length; i += 1) {
        const child = value[i];
        if (child && (isPlainObject(child) || Array.isArray(child))) {
          queue.push({ value: child, path: [...currentPath, String(i)] });
        }
      }
    }
  }

  return null;
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
      name.includes("x86-64") ||
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
  // Use a generous upper bound to avoid infinite loops on unexpected API behavior, but don't
  // artificially cap small/medium releases (some repos upload hundreds of assets).
  const maxPages = 100;
  for (let page = 1; page <= maxPages; page += 1) {
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
    console.error(`verify-tauri-updater-assets: ${msg}`);
    console.error(`Usage: node scripts/verify-tauri-updater-assets.mjs <tag> [--repo owner/repo]`);
    process.exit(1);
  }

  if (args.help) {
    console.log(`Usage: node scripts/verify-tauri-updater-assets.mjs <tag> [--repo owner/repo]`);
    process.exit(0);
  }

  const tag = args.tag ?? process.env.GITHUB_REF_NAME;
  if (!tag) {
    console.error(`Missing tag name. Usage: node scripts/verify-tauri-updater-assets.mjs <tag>`);
    process.exit(1);
  }

  const repo = args.repo ?? process.env.GITHUB_REPOSITORY;
  if (!repo) {
    console.error(
      formatBlock(`Missing repo name`, [
        `Set GITHUB_REPOSITORY in the environment (e.g. "owner/repo"), or pass --repo owner/repo.`,
      ]),
    );
    process.exit(1);
  }

  const token = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!token) {
    console.error(
      formatBlock(`Missing GitHub token`, [
        `Set GITHUB_TOKEN (recommended for GitHub Actions) or GH_TOKEN (for local runs).`,
      ]),
    );
    process.exit(1);
  }

  const apiBase = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/$/, "");

  const updaterExpectation = await readUpdaterConfigExpectation();
  const wantsManifestSig = updaterExpectation.expectsManifestSignature;

  const retryDelaysMs = [2000, 4000, 8000, 12000, 20000];
  const maxAttempts = retryDelaysMs.length + 1;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    if (attempt > 1) {
      const delayMs = retryDelaysMs[attempt - 2] ?? 0;
      console.error(
        `Release assets not fully visible yet; retrying in ${(delayMs / 1000).toFixed(0)}s (attempt ${attempt}/${maxAttempts})...`,
      );
      await sleep(delayMs);
    }

    try {
      const result = await verifyOnce({
        apiBase,
        repo,
        tag,
        token,
        wantsManifestSig,
      });

      // Success.
      for (const line of result.logLines) console.log(line);
      console.log(`Release asset verification passed.`);
      return;
    } catch (err) {
      const isLastAttempt = attempt === maxAttempts;
      if (err instanceof VerificationFailure) {
        if (err.retryable && !isLastAttempt) {
          continue;
        }

        // Final retryable failure (or non-retryable): print the full context.
        err.print(process.stderr);
        process.exit(1);
      }

      const msg = err instanceof Error ? err.stack || err.message : String(err);
      console.error(`verify-tauri-updater-assets: ${msg}`);
      process.exit(1);
    }
  }
}

// Only execute the CLI when invoked as the entrypoint; allow importing this module from node:test.
if (path.resolve(process.argv[1] ?? "") === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    const msg = err instanceof Error ? err.stack || err.message : String(err);
    console.error(`verify-tauri-updater-assets: ${msg}`);
    process.exit(1);
  });
}

/**
 * @param {{ apiBase: string; repo: string; tag: string; token: string; wantsManifestSig: boolean }} opts
 */
async function verifyOnce({ apiBase, repo, tag, token, wantsManifestSig }) {
  /** @type {string[]} */
  const logLines = [];

  /** @type {Array<{ heading: string; details: string[] }>} */
  const blocks = [];

  let release;
  try {
    release = await githubApiJson(apiBase, `/repos/${repo}/releases/tags/${encodeURIComponent(tag)}`, token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      throw new VerificationFailure({
        retryable: true,
        logLines: [],
        blocks: [
          {
            heading: "Failed to fetch GitHub Release",
            details: [`Repo: ${repo}`, `Tag: ${tag}`, `Error: ${msg}`],
          },
        ],
      });
    }

  const releaseId = release?.id;
  if (typeof releaseId !== "number") {
    throw new VerificationFailure({
      retryable: false,
      logLines: [],
      blocks: [
        {
          heading: "Unexpected GitHub API response",
          details: [`Missing numeric release id for tag ${tag}.`],
        },
      ],
    });
  }

  let assets;
  try {
    assets = await listReleaseAssets(apiBase, repo, releaseId, token);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new VerificationFailure({
      retryable: true,
      logLines: [],
      blocks: [
        {
          heading: "Failed to list GitHub Release assets",
          details: [`Repo: ${repo}`, `Tag: ${tag}`, `Release id: ${String(releaseId)}`, `Error: ${msg}`],
        },
      ],
    });
  }

  /** @type {Map<string, any>} */
  const assetByName = new Map();
  for (const asset of assets) {
    if (asset && typeof asset.name === "string") {
      assetByName.set(asset.name, asset);
    }
  }

  logLines.push(`Release asset verification: ${repo}@${tag} (draft=${Boolean(release?.draft)}, assets=${assets.length})`);

  const manifestAsset = assetByName.get("latest.json");
  if (!manifestAsset) {
    throw new VerificationFailure({
      retryable: true,
      logLines,
      blocks: [
        {
          heading: "Missing updater manifest",
          details: [
            `Release does not contain "latest.json".`,
            `This asset is expected to be uploaded by tauri-apps/tauri-action.`,
          ],
        },
      ],
    });
  }

  const manifestSigAsset = assetByName.get("latest.json.sig");
  if (wantsManifestSig && !manifestSigAsset) {
    blocks.push({
      heading: "Missing updater manifest signature",
      details: [
        `apps/desktop/src-tauri/tauri.conf.json has plugins.updater.active=true and a non-placeholder pubkey,`,
        `but the release is missing "latest.json.sig".`,
      ],
    });
  }

  /** @type {Buffer} */
  let manifestBuf;
  try {
    manifestBuf = await downloadReleaseAsset(apiBase, repo, manifestAsset.id, token);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new VerificationFailure({
      retryable: true,
      logLines,
      blocks: [
        {
          heading: "Failed to download updater manifest",
          details: [`Asset: latest.json`, `Error: ${msg}`],
        },
      ],
    });
  }

  if (manifestSigAsset) {
    try {
      await downloadReleaseAsset(apiBase, repo, manifestSigAsset.id, token);
      logLines.push(`- Found updater manifest signature: latest.json.sig`);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      blocks.push({
        heading: "Failed to download updater manifest signature",
        details: [`Asset: latest.json.sig`, `Error: ${msg}`],
      });
    }
  } else {
    logLines.push(`- No updater manifest signature: latest.json.sig (not found)`);
  }

  /** @type {any} */
  let manifest;
  try {
    manifest = JSON.parse(manifestBuf.toString("utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new VerificationFailure({
      retryable: false,
      logLines,
      blocks: [
        {
          heading: "Invalid latest.json",
          details: [`Failed to parse JSON.`, `Error: ${msg}`],
        },
      ],
    });
  }

  const manifestVersion = typeof manifest?.version === "string" ? manifest.version.trim() : "";
  if (manifestVersion) {
    const tagVersion = normalizeTagVersion(tag);
    if (manifestVersion !== tagVersion) {
      throw new VerificationFailure({
        retryable: false,
        logLines,
        blocks: [
          {
            heading: "Updater manifest version mismatch",
            details: [
              `Tag: ${tag} (version ${JSON.stringify(tagVersion)})`,
              `latest.json: version ${JSON.stringify(manifestVersion)}`,
            ],
          },
        ],
      });
    }
  }

  const platformsFound = findPlatformsObject(manifest);
  const platforms = platformsFound?.platforms;
  if (!platformsFound || !platforms) {
    throw new VerificationFailure({
      retryable: false,
      logLines,
      blocks: [
        {
          heading: "Invalid latest.json",
          details: [
            `Expected latest.json to contain a "platforms" object mapping target → { url, signature, ... }.`,
          ],
        },
      ],
    });
  }

  const platformKeys = Object.keys(platforms).sort();
  if (platformKeys.length === 0) {
    throw new VerificationFailure({
      retryable: false,
      logLines,
      blocks: [
        {
          heading: "Invalid latest.json",
          details: [`"platforms" is empty; updater has no targets.`],
        },
      ],
    });
  }

  // Enforce stable updater platform key naming (see docs/desktop-updater-target-mapping.md).
  const expectedKeysSorted = EXPECTED_PLATFORM_KEYS.slice().sort();
  const expectedKeySet = new Set(EXPECTED_PLATFORM_KEYS);
  const missingKeys = expectedKeysSorted.filter((k) =>
    !Object.prototype.hasOwnProperty.call(platforms, k),
  );
  const otherKeys = platformKeys.filter((k) => !expectedKeySet.has(k));
  if (missingKeys.length > 0) {
    const formatKeyList = (keys) => keys.map((k) => `    - ${k}`).join("\n");
    blocks.push({
      heading: "Missing required latest.json.platforms keys (Tauri updater target identifiers)",
      details: [
        `Expected (${expectedKeysSorted.length}):\n${formatKeyList(expectedKeysSorted)}`,
        `Actual (${platformKeys.length}):\n${formatKeyList(platformKeys)}`,
        `Missing (${missingKeys.length}):\n${formatKeyList(missingKeys)}`,
        ...(otherKeys.length > 0
          ? [`Other keys present (${otherKeys.length}):\n${formatKeyList(otherKeys)}`]
          : []),
        `If you upgraded Tauri/tauri-action, update docs/desktop-updater-target-mapping.md and scripts/verify-tauri-updater-assets.mjs together.`,
      ],
    });
  }

  const groups = groupPlatformKeys(platformKeys);
  const platformsPath = platformsFound.path.join(".");
  logLines.push(`- Platforms in latest.json (${platformKeys.length}) at ${platformsPath}: ${platformKeys.join(", ")}`);

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
    const expectedArchForAsset =
      platformKey === "windows-x86_64" || platformKey === "linux-x86_64"
        ? "x86_64"
        : platformKey === "windows-aarch64" || platformKey === "linux-aarch64"
          ? "aarch64"
          : null;

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
      if (expectedArchForAsset && !archMatchesAssetName(expectedArchForAsset, assetName)) {
        const archHint =
          expectedArchForAsset === "x86_64"
            ? "x86_64/x64/amd64"
            : expectedArchForAsset === "aarch64"
              ? "arm64/aarch64"
              : expectedArchForAsset;
        missing.push(
          `Updater asset arch mismatch for ${platformKey}: ${assetName} (expected ${archHint} token in filename)`,
        );
      }

      if (!assetByName.has(assetName)) {
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

  for (const entry of updaterEntries) {
    const sigStatus = entry.sigAssetPresent
      ? entry.inlineSig
        ? `${entry.sigAssetName} + inline`
        : entry.sigAssetName
      : entry.inlineSig
        ? "inline (no .sig asset)"
        : "MISSING";
    logLines.push(`  - ${entry.platform}: ${entry.assetName} (sig: ${sigStatus})`);
  }

  const assetNames = [...assetByName.keys()];

  const hasAny = (suffix) => assetNames.some((n) => n.toLowerCase().endsWith(suffix.toLowerCase()));

  if (!hasAny(".dmg")) {
    missing.push(`Missing macOS installer: no .dmg asset found in the release`);
  }

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

  const requireArtifactSigAssets = wantsManifestSig;
  if (requireArtifactSigAssets) {
    const signedKinds = [
      { label: "macOS installer (.dmg)", matches: (name) => name.toLowerCase().endsWith(".dmg") },
      {
        label: "macOS updater archive (.tar.gz/.tgz)",
        matches: isMacUpdaterArchiveAssetName,
      },
      { label: "Windows installer (.msi)", matches: (name) => name.toLowerCase().endsWith(".msi") },
      { label: "Windows installer (.exe)", matches: (name) => name.toLowerCase().endsWith(".exe") },
      { label: "Linux bundle (.AppImage)", matches: (name) => name.toLowerCase().endsWith(".appimage") },
      { label: "Linux package (.deb)", matches: (name) => name.toLowerCase().endsWith(".deb") },
      { label: "Linux package (.rpm)", matches: (name) => name.toLowerCase().endsWith(".rpm") },
    ];

    /** @type {string[]} */
    const missingSigAssets = [];
    for (const name of assetNames) {
      const kind = signedKinds.find((k) => k.matches(name));
      if (!kind) continue;
      const sigName = `${name}.sig`;
      if (!assetByName.has(sigName)) {
        missingSigAssets.push(`Missing signature asset for ${kind.label}: ${sigName}`);
      }
    }
    missing.push(...missingSigAssets);
  }

  if (missing.length > 0) {
    blocks.push({
      heading: "Release asset verification failed",
      details: [...new Set(missing)],
    });
  }

  if (blocks.length > 0) {
    throw new VerificationFailure({ blocks, retryable: true, logLines });
  }

  return { logLines };
}
