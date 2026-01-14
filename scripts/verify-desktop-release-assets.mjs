#!/usr/bin/env node
/**
 * Verify GitHub Release assets produced by `.github/workflows/release.yml` for
 * the desktop (Tauri) app.
 *
 * Validates:
 * - `latest.json` updater manifest structure + version correctness
 * - every `platforms[*].url` in `latest.json` points at an uploaded asset
 * - `latest.json.sig` exists and verifies against the updater public key in
 *   `apps/desktop/src-tauri/tauri.conf.json` (`plugins.updater.pubkey`)
 * - optional per-bundle signature coverage (either `signature` in JSON or a
 *   sibling `<bundle>.sig` release asset)
 *
 * Generates:
 * - `SHA256SUMS.txt` for primary bundle artifacts (excluding `.sig` files by default).
 *   Use `--all-assets` to hash every release asset and `--include-sigs` to include `.sig` files.
 *
 * Usage:
 *   node scripts/verify-desktop-release-assets.mjs --tag v0.1.0 --repo owner/repo
 *
 * Required env:
 *   GITHUB_TOKEN (or GH_TOKEN) - token with access to the repo's releases
 */

import { createHash, verify } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { Readable } from "node:stream";
import { setTimeout as sleep } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import {
  ed25519PublicKeyFromRaw,
  parseTauriUpdaterPubkey,
  parseTauriUpdaterSignature,
} from "./ci/tauri-minisign.mjs";

class ActionableError extends Error {
  /**
   * @param {string} heading
   * @param {string[]} [details]
   */
  constructor(heading, details = []) {
    const body = details.length > 0 ? `\n${details.map((d) => `- ${d}`).join("\n")}` : "";
    super(`${heading}${body}`);
    this.name = "ActionableError";
  }
}

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfigRelativePath = "apps/desktop/src-tauri/tauri.conf.json";
const tauriConfigPath = path.join(repoRoot, tauriConfigRelativePath);
const PLACEHOLDER_PUBKEY = "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY";

/**
 * @param {string} value
 */
function normalizeVersion(value) {
  const trimmed = value.trim();
  return trimmed.startsWith("v") ? trimmed.slice(1) : trimmed;
}

/**
 * @param {string} maybeUrl
 * @returns {string | null}
 */
function filenameFromUrl(maybeUrl) {
  const trimmed = maybeUrl.trim();
  if (!trimmed) return null;
  try {
    const url = new URL(trimmed);
    const last = url.pathname.split("/").filter(Boolean).at(-1) ?? "";
    return last ? decodeURIComponent(last) : null;
  } catch {
    // Be liberal: fall back to extracting via string ops if the URL isn't fully qualified.
    const noQuery = trimmed.split("?")[0].split("#")[0];
    const last = noQuery.split("/").filter(Boolean).at(-1) ?? "";
    return last || null;
  }
}

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown>}
 */
function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Best-effort locate a Tauri updater `platforms` map within a JSON payload.
 * (Some formats nest the object; keep this liberal.)
 *
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

      if (currentPath.length >= 8) continue;
      for (const [key, child] of Object.entries(value)) {
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
 * @param {unknown} value
 * @param {string} context
 */
function requireNonEmptyString(value, context) {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new ActionableError(`${context} must be a non-empty string.`);
  }
  return value.trim();
}

/**
 * @param {number} bytes
 */
function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return `${bytes}`;
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KiB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MiB`;
  const gb = mb / 1024;
  return `${gb.toFixed(2)} GiB`;
}

// --- Optional expectations layer ------------------------------------------------

const EXPECT_FLAG_MAP = new Map([
  ["--expect-windows-x64", "windows-x64"],
  ["--expect-windows-arm64", "windows-arm64"],
  ["--expect-macos-x64", "macos-x64"],
  ["--expect-macos-arm64", "macos-arm64"],
  ["--expect-macos-universal", "macos-universal"],
  ["--expect-linux-x64", "linux-x64"],
  ["--expect-linux-arm64", "linux-arm64"],
]);

const ARCH_TOKENS = {
  x64: ["x64", "x86_64", "amd64"],
  arm64: ["arm64", "aarch64"],
  universal: ["universal"],
};

/** @typedef {"windows" | "macos" | "linux"} DesktopOs */
/** @typedef {"x64" | "arm64" | "universal"} DesktopArch */
/**
 * @typedef {{
 *   id: string,
 *   os: DesktopOs,
 *   arch: DesktopArch,
 *   installerExts: string[],
 *   updaterPlatformKeys: string[],
 *   allowMissingArchInInstallerName?: boolean,
 * }} ExpectedTarget
 */

/** @type {Record<string, ExpectedTarget>} */
const EXPECTED_TARGETS = {
  "windows-x64": {
    id: "windows-x64",
    os: "windows",
    arch: "x64",
    installerExts: [".msi", ".exe"],
    updaterPlatformKeys: ["windows-x86_64", "windows-x64"],
  },
  "windows-arm64": {
    id: "windows-arm64",
    os: "windows",
    arch: "arm64",
    installerExts: [".msi", ".exe"],
    updaterPlatformKeys: ["windows-aarch64", "windows-arm64"],
  },
  "macos-x64": {
    id: "macos-x64",
    os: "macos",
    arch: "x64",
    installerExts: [".dmg", ".pkg"],
    updaterPlatformKeys: ["darwin-x86_64", "darwin-x64", "macos-x86_64"],
  },
  "macos-arm64": {
    id: "macos-arm64",
    os: "macos",
    arch: "arm64",
    installerExts: [".dmg", ".pkg"],
    updaterPlatformKeys: ["darwin-aarch64", "darwin-arm64", "macos-aarch64", "macos-arm64"],
  },
  "macos-universal": {
    id: "macos-universal",
    os: "macos",
    arch: "universal",
    installerExts: [".dmg", ".pkg"],
    // Formula's release workflow produces a single universal macOS updater archive, but Tauri's
    // updater target at runtime is still per-arch. We therefore expect the universal archive to
    // appear under BOTH `darwin-x86_64` and `darwin-aarch64` keys in `latest.json.platforms`.
    //
    // `darwin-universal` may also be present if tauri-action's `updaterJsonKeepUniversal` is enabled.
    updaterPlatformKeys: [
      "darwin-x86_64",
      "darwin-aarch64",
      "darwin-universal",
      "macos-universal",
    ],
    // Some universal builds ship a single installer that omits the arch token (because the
    // installer itself is universal). This is allowed only when `--expect-macos-universal` is set.
    allowMissingArchInInstallerName: true,
  },
  "linux-x64": {
    id: "linux-x64",
    os: "linux",
    arch: "x64",
    installerExts: [".deb", ".rpm", ".AppImage", ".appimage"],
    updaterPlatformKeys: ["linux-x86_64", "linux-x64"],
  },
  "linux-arm64": {
    id: "linux-arm64",
    os: "linux",
    arch: "arm64",
    installerExts: [".deb", ".rpm", ".AppImage", ".appimage"],
    updaterPlatformKeys: ["linux-aarch64", "linux-arm64"],
  },
};

/**
 * @param {string} value
 */
function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Boundary-aware "token exists" regex builder.
 *
 * @param {string[]} tokens
 */
function tokenRegex(tokens) {
  const inner = tokens.map((t) => escapeRegex(t)).join("|");
  // Treat non-alphanumerics as boundaries so `_x86_64` and `-arm64` are matched as tokens.
  return new RegExp(`(?:^|[^a-z0-9])(?:${inner})(?:[^a-z0-9]|$)`, "i");
}

/**
 * @param {string} name
 * @param {string[]} exts
 */
function hasAnyExtension(name, exts) {
  const lower = name.toLowerCase();
  return exts.some((ext) => lower.endsWith(ext.toLowerCase()));
}

/**
 * @param {string} name
 * @param {string} version
 */
function containsVersion(name, version) {
  return name.includes(version) || name.includes(`v${version}`);
}

/**
 * @param {string} name
 */
function containsAnyArchToken(name) {
  return tokenRegex([...ARCH_TOKENS.x64, ...ARCH_TOKENS.arm64, ...ARCH_TOKENS.universal]).test(name);
}

/**
 * @param {string} name
 * @param {DesktopArch} arch
 */
function containsArchToken(name, arch) {
  return tokenRegex(ARCH_TOKENS[arch]).test(name);
}

/**
 * @param {DesktopOs} os
 */
function osArtifactRegex(os) {
  switch (os) {
    case "windows":
      // Includes updater artifacts like `.msi.zip`, `.exe.sig`, etc.
      return /(?:\.msi|\.exe)(?:\.|$)/i;
    case "macos":
      // Includes `.dmg`, `.pkg`, updater `.app.tar.gz`, etc.
      return /(?:\.dmg|\.pkg|\.app\.tar\.gz)(?:\.|$)/i;
    case "linux":
      // Includes `.AppImage.tar.gz`, `.deb.sig`, etc.
      return /(?:\.appimage|\.deb|\.rpm)(?:\.|$)/i;
  }
}

/**
 * @param {string[]} assetNames
 * @param {DesktopOs} os
 */
function listOsArtifacts(assetNames, os) {
  const re = osArtifactRegex(os);
  return assetNames.filter((name) => re.test(name));
}

/**
 * @param {string[]} assetNames
 * @param {ExpectedTarget} target
 * @param {string} version
 */
function findInstallerAssetsForTarget(assetNames, target, version) {
  return assetNames
    .filter((name) => hasAnyExtension(name, target.installerExts))
    .filter((name) => containsVersion(name, version))
    .filter((name) => {
      if (containsArchToken(name, target.arch)) return true;
      if (target.allowMissingArchInInstallerName && !containsAnyArchToken(name)) return true;
      return false;
    });
}

/**
 * @param {string[]} assetNames
 * @param {DesktopOs} os
 * @param {DesktopArch[]} expectedArchsForOs
 * @param {string} version
 * @param {boolean} allowMissingArchToken
 */
function findAmbiguousAssetsForOs(assetNames, os, expectedArchsForOs, version, allowMissingArchToken) {
  const tokens = expectedArchsForOs.flatMap((arch) => ARCH_TOKENS[arch]);
  const re = tokenRegex(tokens);
  const artifactNames = listOsArtifacts(assetNames, os).filter((name) => containsVersion(name, version));

  return artifactNames.filter((name) => {
    if (re.test(name)) return false;
    if (allowMissingArchToken && !containsAnyArchToken(name)) return false;
    return true;
  });
}

/**
 * @param {unknown} value
 */
function isStringArray(value) {
  return Array.isArray(value) && value.every((v) => typeof v === "string");
}

/**
 * @param {string} filePath
 * @returns {Promise<{ targets: string[] }>}
 */
async function loadExpectationsFile(filePath) {
  const resolvedPath = path.isAbsolute(filePath) ? filePath : path.join(repoRoot, filePath);
  let raw;
  try {
    raw = await readFile(resolvedPath, "utf8");
  } catch (err) {
    throw new ActionableError(`Failed to read expectations file ${filePath}.`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }

  /** @type {any} */
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    throw new ActionableError(`Failed to parse expectations file ${filePath} as JSON.`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }

  if (isStringArray(parsed)) {
    return { targets: parsed };
  }

  if (parsed && typeof parsed === "object") {
    if (isStringArray(parsed.expect)) return { targets: parsed.expect };
    if (isStringArray(parsed.targets)) return { targets: parsed.targets };
  }

  throw new ActionableError(`Invalid expectations file ${filePath}.`, [
    `Expected an array of strings or { "expect": string[] }.`,
  ]);
}

function usage() {
  const cmd = "node scripts/verify-desktop-release-assets.mjs";
  console.log(
    [
      "Verify desktop GitHub Release assets + generate SHA256SUMS.txt.",
      "",
      "Options:",
      "  --tag <tag>        Release tag (default: env GITHUB_REF_NAME)",
      "  --repo <owner/repo> GitHub repo (default: env GITHUB_REPOSITORY)",
      "  --out <path>       Output path for SHA256SUMS.txt (default: ./SHA256SUMS.txt)",
      "  --all-assets       Hash all release assets (still excludes .sig by default)",
      "  --include-sigs     Include .sig assets in SHA256SUMS (use with --all-assets to match CI)",
      "  --dry-run          Validate manifest/assets only (skip bundle hashing)",
      "  --verify-assets    Download updater assets referenced in latest.json and verify their signatures (slow)",
      "",
      "Expectations (optional; off by default):",
      "  --expectations <file>      Load expected targets from a JSON file",
      "                             (see scripts/release-asset-expectations.json for a template)",
      "  --expect-windows-x64",
      "  --expect-windows-arm64",
      "  --expect-macos-universal",
      "  --expect-macos-x64",
      "  --expect-macos-arm64",
      "  --expect-linux-x64",
      "  --expect-linux-arm64",
      "  -h, --help         Show help",
      "",
      "Env:",
      "  GITHUB_TOKEN       GitHub token with permission to read release assets",
      "  GH_TOKEN           Alternative token env var (supported for local runs)",
      "",
      "Examples:",
      `  ${cmd} --tag v0.2.3 --repo wilson/formula`,
      `  ${cmd} --dry-run --tag v0.2.3`,
    ].join("\n")
  );
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {{ tag?: string; repo?: string; out?: string; dryRun: boolean; verifyAssets: boolean; help: boolean; expectationsFile?: string; expectedTargets: string[]; includeSigs: boolean; allAssets: boolean }} */
  const parsed = {
    dryRun: false,
    verifyAssets: false,
    help: false,
    expectedTargets: [],
    includeSigs: false,
    allAssets: false,
  };

  const takeValue = (i, flag) => {
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) {
      throw new ActionableError(`Missing value for ${flag}.`, [
        `Usage: ${flag} <value>`,
        "Run with --help for full usage.",
      ]);
    }
    return value;
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--dry-run") {
      parsed.dryRun = true;
      continue;
    }
    if (arg === "--verify-assets") {
      parsed.verifyAssets = true;
      continue;
    }
    if (arg === "--include-sigs") {
      parsed.includeSigs = true;
      continue;
    }
    if (arg === "--all-assets") {
      parsed.allAssets = true;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      parsed.help = true;
      continue;
    }
    if (arg === "--tag") {
      parsed.tag = takeValue(i, "--tag");
      i++;
      continue;
    }
    if (arg.startsWith("--tag=")) {
      parsed.tag = arg.slice("--tag=".length);
      continue;
    }
    if (arg === "--repo") {
      parsed.repo = takeValue(i, "--repo");
      i++;
      continue;
    }
    if (arg.startsWith("--repo=")) {
      parsed.repo = arg.slice("--repo=".length);
      continue;
    }
    if (arg === "--out") {
      parsed.out = takeValue(i, "--out");
      i++;
      continue;
    }
    if (arg.startsWith("--out=")) {
      parsed.out = arg.slice("--out=".length);
      continue;
    }
    if (arg === "--expectations") {
      parsed.expectationsFile = takeValue(i, "--expectations");
      i++;
      continue;
    }
    if (arg.startsWith("--expectations=")) {
      parsed.expectationsFile = arg.slice("--expectations=".length);
      continue;
    }

    const mappedExpectation = EXPECT_FLAG_MAP.get(arg);
    if (mappedExpectation) {
      parsed.expectedTargets.push(mappedExpectation);
      continue;
    }

    throw new ActionableError(`Unknown argument: ${arg}`, [
      "Run with --help for full usage.",
    ]);
  }

  return parsed;
}

/**
 * @param {string} repo
 * @param {string} tag
 */
function buildReleaseByTagApiUrl(repo, tag) {
  const normalizedTag = tag.startsWith("refs/tags/") ? tag.slice("refs/tags/".length) : tag;
  // Tag is a path component; encode for safety (e.g. v1.2.3-beta).
  return `https://api.github.com/repos/${repo}/releases/tags/${encodeURIComponent(normalizedTag)}`;
}

/**
 * @param {string} repo
 * @param {number} releaseId
 * @param {number} page
 */
function buildReleaseAssetsApiUrl(repo, releaseId, page) {
  return `https://api.github.com/repos/${repo}/releases/${releaseId}/assets?per_page=100&page=${page}`;
}

/**
 * @param {string} url
 * @param {{ token: string; accept?: string }} opts
 */
async function fetchGitHub(url, { token, accept }) {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: accept ?? "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "formula-release-assets-verifier",
    },
  });
  return res;
}

/**
 * @param {string} url
 * @param {{ token: string }} opts
 */
async function fetchGitHubJson(url, { token }) {
  const res = await fetchGitHub(url, { token });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`GitHub API request failed: ${res.status} ${res.statusText}`, [
      `GET ${url}`,
      body ? `Response: ${body}` : "Response body was empty.",
    ]);
  }
  return /** @type {any} */ (await res.json());
}

/**
 * @param {Response} res
 */
async function safeReadBody(res) {
  try {
    const text = await res.text();
    return text.trim().length > 0 ? text.trim().slice(0, 1000) : "";
  } catch {
    return "";
  }
}

/**
 * @param {string} repo
 * @param {string} tag
 * @param {string} token
 */
async function getReleaseByTag(repo, tag, token) {
  const url = buildReleaseByTagApiUrl(repo, tag);
  const release = await fetchGitHubJson(url, { token });
  if (!release || typeof release !== "object") {
    throw new ActionableError(`Unexpected GitHub API response for release tag ${tag}.`);
  }
  if (typeof release.id !== "number") {
    throw new ActionableError(`GitHub API response missing "id" for release tag ${tag}.`);
  }
  return release;
}

/**
 * @param {string} repo
 * @param {number} releaseId
 * @param {string} token
 */
async function listAllReleaseAssets(repo, releaseId, token) {
  const assets = [];
  for (let page = 1; page <= 50; page++) {
    const url = buildReleaseAssetsApiUrl(repo, releaseId, page);
    const pageAssets = await fetchGitHubJson(url, { token });
    if (!Array.isArray(pageAssets)) {
      throw new ActionableError("Unexpected GitHub API response for release assets list.", [
        `Expected an array, got ${typeof pageAssets}.`,
      ]);
    }

    for (const asset of pageAssets) {
      assets.push(asset);
    }

    if (pageAssets.length < 100) break;
  }
  return assets;
}

/**
 * Download a release asset using the GitHub REST API `asset.url` endpoint
 * (not `browser_download_url`) so it works for private repos too.
 *
 * @param {{ name: string; url?: string; browser_download_url?: string }} asset
 * @param {string} token
 */
async function downloadReleaseAssetText(asset, token) {
  const downloadUrl =
    typeof asset.url === "string" && asset.url.length > 0
      ? asset.url
      : asset.browser_download_url;
  if (typeof downloadUrl !== "string" || downloadUrl.trim().length === 0) {
    throw new ActionableError(`Release asset "${asset.name}" is missing a download URL.`, [
      `Expected GitHub API to provide "url" or "browser_download_url".`,
    ]);
  }

  const res = await fetchGitHub(downloadUrl, {
    token,
    accept: "application/octet-stream",
  });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`Failed to download release asset "${asset.name}".`, [
      `${res.status} ${res.statusText}`,
      body ? `Response: ${body}` : "Response body was empty.",
      `URL: ${downloadUrl}`,
    ]);
  }
  return await res.text();
}

/**
 * Download a release asset and return the raw bytes.
 *
 * Note: this buffers the entire asset in memory (required for Ed25519 verification). Keep usage
 * behind explicit flags for large binaries.
 *
 * @param {{ name: string; url?: string; browser_download_url?: string; size?: number }} asset
 * @param {string} token
 */
async function downloadReleaseAssetBytes(asset, token) {
  const downloadUrl =
    typeof asset.url === "string" && asset.url.length > 0
      ? asset.url
      : asset.browser_download_url;
  if (typeof downloadUrl !== "string" || downloadUrl.trim().length === 0) {
    throw new ActionableError(`Release asset "${asset.name}" is missing a download URL.`, [
      `Expected GitHub API to provide "url" or "browser_download_url".`,
    ]);
  }

  const res = await fetchGitHub(downloadUrl, {
    token,
    accept: "application/octet-stream",
  });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`Failed to download release asset "${asset.name}".`, [
      `${res.status} ${res.statusText}`,
      body ? `Response: ${body}` : "Response body was empty.",
      `URL: ${downloadUrl}`,
    ]);
  }
  return Buffer.from(await res.arrayBuffer());
}

/**
 * @param {{ name: string; size?: number; url?: string; browser_download_url?: string }} asset
 * @param {string} token
 */
async function sha256OfReleaseAsset(asset, token) {
  const downloadUrl =
    typeof asset.url === "string" && asset.url.length > 0
      ? asset.url
      : asset.browser_download_url;
  if (typeof downloadUrl !== "string" || downloadUrl.trim().length === 0) {
    throw new ActionableError(`Release asset "${asset.name}" is missing a download URL.`, [
      `Expected GitHub API to provide "url" or "browser_download_url".`,
    ]);
  }

  const res = await fetchGitHub(downloadUrl, {
    token,
    accept: "application/octet-stream",
  });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`Failed to download asset for hashing: "${asset.name}".`, [
      `${res.status} ${res.statusText}`,
      body ? `Response: ${body}` : "Response body was empty.",
      `URL: ${downloadUrl}`,
    ]);
  }
  if (!res.body) {
    throw new ActionableError(`No response body while downloading "${asset.name}" for hashing.`);
  }

  const hash = createHash("sha256");
  // `res.body` is a WHATWG ReadableStream. Convert to a Node stream so we can
  // reliably async-iterate across Node versions.
  const nodeStream = Readable.fromWeb(res.body);
  for await (const chunk of nodeStream) {
    hash.update(chunk);
  }
  return hash.digest("hex");
}

/**
 * @param {string} name
 */
function isPrimaryBundleAssetName(name) {
  const lower = name.toLowerCase();
  // `.app.tar.gz` is covered by `.tar.gz`.
  const suffixes = [
    ".dmg",
    ".tar.gz",
    ".msi",
    ".exe",
    ".appimage",
    ".deb",
    ".rpm",
    ".zip",
    ".pkg",
  ];
  return suffixes.some((s) => lower.endsWith(s));
}

/**
 * @param {string} name
 * @param {{ includeSigs: boolean }} opts
 */
function isPrimaryBundleOrSig(name, { includeSigs }) {
  if (isPrimaryBundleAssetName(name)) return true;
  if (!includeSigs) return false;
  if (!name.endsWith(".sig")) return false;
  const base = name.slice(0, -".sig".length);
  return isPrimaryBundleAssetName(base);
}

/**
 * @param {any} manifest
 * @param {string} expectedVersion
 * @param {Map<string, any>} assetsByName
 */
function validateLatestJson(manifest, expectedVersion, assetsByName) {
  /** @type {string[]} */
  const errors = [];

  if (!manifest || typeof manifest !== "object") {
    throw new ActionableError("latest.json did not parse into an object.");
  }

  const manifestVersionRaw = manifest.version;
  if (typeof manifestVersionRaw !== "string" || manifestVersionRaw.trim().length === 0) {
    errors.push(`latest.json is missing a non-empty top-level "version" string.`);
  } else {
    const expected = normalizeVersion(expectedVersion);
    const actual = normalizeVersion(manifestVersionRaw);
    if (expected !== actual) {
      errors.push(
        `latest.json version mismatch: expected ${JSON.stringify(expected)} (from ${tauriConfigRelativePath}), got ${JSON.stringify(actual)}.`
      );
    }
  }

  const foundPlatforms = findPlatformsObject(manifest);
  if (!foundPlatforms) {
    errors.push(`latest.json is missing a "platforms" object.`);
  } else {
    const { platforms } = foundPlatforms;
    const keys = Object.keys(platforms);
    const requiredPlatformKeys = [
      // Keep this list in sync with docs/desktop-updater-target-mapping.md and
      // scripts/ci/validate-updater-manifest.mjs (which is the stricter CI validator).
      // macOS universal bundles are written under both arch keys (same updater archive URL).
      "darwin-x86_64",
      "darwin-aarch64",
      "windows-x86_64",
      "windows-aarch64",
      "linux-x86_64",
      "linux-aarch64",
    ];
    for (const requiredKey of requiredPlatformKeys) {
      if (!Object.prototype.hasOwnProperty.call(platforms, requiredKey)) {
        errors.push(
          `latest.json platforms is missing required key ${JSON.stringify(requiredKey)}. Present keys: ${keys.length > 0 ? keys.join(", ") : "(none)"}`
        );
      }
    }

    for (const [platformKey, platformEntry] of Object.entries(platforms)) {
      if (!platformEntry || typeof platformEntry !== "object") {
        errors.push(`latest.json platforms[${JSON.stringify(platformKey)}] must be an object.`);
        continue;
      }

      const url = platformEntry.url;
      if (typeof url !== "string" || url.trim().length === 0) {
        errors.push(
          `latest.json platforms[${JSON.stringify(platformKey)}] is missing a non-empty "url" string.`
        );
        continue;
      }

      const filename = filenameFromUrl(url);
      if (!filename) {
        errors.push(
          `latest.json platforms[${JSON.stringify(platformKey)}].url did not contain a valid filename: ${JSON.stringify(url)}`
        );
        continue;
      }

      if (!assetsByName.has(filename)) {
        errors.push(
          `latest.json platforms[${JSON.stringify(platformKey)}].url references "${filename}", but that asset is not present in the GitHub Release.`
        );
      }

      const signature = platformEntry.signature;
      const hasInlineSignature = typeof signature === "string" && signature.trim().length > 0;
      if (!hasInlineSignature) {
        const sigAssetName = `${filename}.sig`;
        if (!assetsByName.has(sigAssetName)) {
          errors.push(
            `latest.json platforms[${JSON.stringify(platformKey)}] is missing a non-empty "signature" AND "${sigAssetName}" was not found in the release assets.`
          );
        }
      }
    }
  }

  if (errors.length > 0) {
    throw new ActionableError("Desktop release asset verification failed: latest.json validation errors.", errors);
  }
}

/**
 * (Optional) Verify updater asset signatures for each platform entry in latest.json.
 *
 * This downloads the referenced updater artifacts (can be large). Keep it behind an explicit flag.
 *
 * @param {{
 *   manifest: any,
 *   assetsByName: Map<string, any>,
 *   token: string,
 *   publicKey: crypto.KeyObject,
 *   pubkeyKeyId: string | null,
 * }} opts
 */
async function verifyUpdaterPlatformAssetSignatures({ manifest, assetsByName, token, publicKey, pubkeyKeyId }) {
  const platforms = manifest?.platforms;
  if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
    throw new ActionableError(
      "latest.json is missing a 'platforms' object; cannot verify updater asset signatures.",
    );
  }

  const entries = Object.entries(platforms);
  if (entries.length === 0) {
    throw new ActionableError(
      "latest.json 'platforms' object is empty; cannot verify updater asset signatures.",
    );
  }

  console.log(`Verifying updater asset signatures for ${entries.length} platform(s)...`);
  for (const [platformKey, entry] of entries) {
    if (!entry || typeof entry !== "object") continue;

    const url = typeof entry.url === "string" ? entry.url.trim() : "";
    if (!url) continue;

    const assetName = filenameFromUrl(url);
    if (!assetName) continue;

    const asset = assetsByName.get(assetName);
    if (!asset) {
      throw new ActionableError(`Updater asset is missing from the GitHub Release.`, [
        `platform: ${platformKey}`,
        `asset: ${assetName}`,
        `url: ${url}`,
      ]);
    }

    const inlineSig = typeof entry.signature === "string" ? entry.signature.trim() : "";
    let signatureText = inlineSig;
    let signatureSource = inlineSig ? "inline" : "asset";

    if (!signatureText) {
      const sigAssetName = `${assetName}.sig`;
      const sigAsset = assetsByName.get(sigAssetName);
      if (!sigAsset) {
        // This should have been caught by validateLatestJson, but keep this check defensive.
        throw new ActionableError(`Missing updater asset signature.`, [
          `platform: ${platformKey}`,
          `asset: ${assetName}`,
          `expected either platforms[${JSON.stringify(platformKey)}].signature or a release asset "${sigAssetName}"`,
        ]);
      }
      signatureText = await downloadReleaseAssetText(sigAsset, token);
      signatureSource = sigAssetName;
    }

    if (signatureText.trim().length === 0) {
      throw new ActionableError(`Updater asset signature file is empty.`, [
        `platform: ${platformKey}`,
        `asset: ${assetName}`,
        `signature source: ${signatureSource}`,
      ]);
    }

    let parsedSig;
    try {
      parsedSig = parseTauriUpdaterSignature(signatureText, `${platformKey}.signature`);
    } catch (err) {
      throw new ActionableError(`Failed to parse updater asset signature.`, [
        `platform: ${platformKey}`,
        `asset: ${assetName}`,
        `signature source: ${signatureSource}`,
        err instanceof Error ? err.message : String(err),
      ]);
    }

    if (parsedSig.keyId && pubkeyKeyId && parsedSig.keyId !== pubkeyKeyId) {
      throw new ActionableError(`Updater asset signature key id mismatch.`, [
        `platform: ${platformKey}`,
        `asset: ${assetName}`,
        `signature key id: ${parsedSig.keyId}`,
        `updater pubkey key id: ${pubkeyKeyId}`,
      ]);
    }

    const size = typeof asset.size === "number" ? asset.size : undefined;
    console.log(
      `- ${platformKey}: downloading ${assetName}${
        size ? ` (${formatBytes(size)})` : ""
      } for signature verification...`,
    );
    const assetBytes = await downloadReleaseAssetBytes(asset, token);
    const ok = verify(null, assetBytes, publicKey, parsedSig.signatureBytes);
    if (!ok) {
      throw new ActionableError(`Updater asset signature verification failed.`, [
        `platform: ${platformKey}`,
        `asset: ${assetName}`,
        `url: ${url}`,
        `signature source: ${signatureSource}`,
        `This usually means the asset or signature was tampered with, or TAURI_PRIVATE_KEY does not match plugins.updater.pubkey.`,
      ]);
    }
  }

  console.log("Verified updater asset signatures.");
}

/**
 * Verify that `latest.json.sig` matches `latest.json` under the updater public key embedded in
 * `apps/desktop/src-tauri/tauri.conf.json` (`plugins.updater.pubkey`).
 *
 * @param {Buffer} latestJsonBytes
 * @param {string} latestSigText
 * @param {string} updaterPubkeyBase64
 */
function verifyUpdaterManifestSignature(latestJsonBytes, latestSigText, updaterPubkeyBase64) {
  let pubkey;
  try {
    pubkey = parseTauriUpdaterPubkey(updaterPubkeyBase64);
  } catch (err) {
    throw new ActionableError(`Invalid updater public key (plugins.updater.pubkey).`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }

  let signature;
  try {
    signature = parseTauriUpdaterSignature(latestSigText, "latest.json.sig");
  } catch (err) {
    throw new ActionableError(`Invalid updater manifest signature file (latest.json.sig).`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }

  if (signature.keyId && pubkey.keyId && signature.keyId !== pubkey.keyId) {
    throw new ActionableError(`Updater manifest signature key id mismatch.`, [
      `latest.json.sig key id: ${signature.keyId}`,
      `plugins.updater.pubkey key id: ${pubkey.keyId}`,
    ]);
  }

  const publicKey = ed25519PublicKeyFromRaw(pubkey.publicKeyBytes);
  const ok = verify(null, latestJsonBytes, publicKey, signature.signatureBytes);
  if (!ok) {
    throw new ActionableError(`Updater manifest signature mismatch.`, [
      `latest.json.sig does not verify latest.json with the configured updater public key.`,
      `This typically means latest.json and latest.json.sig were uploaded/generated inconsistently (race/overwrite), or TAURI_PRIVATE_KEY does not match plugins.updater.pubkey.`,
    ]);
  }
}

/**
 * Optional expectations layer.
 *
 * When enabled via `--expect-*` flags or `--expectations <file>`, asserts:
 * - each expected (os, arch) has at least one installer asset whose filename includes version + arch
 * - latest.json.platforms contains an entry for each expected updater target
 * - no ambiguous multi-arch artifacts whose name omits any arch/universal discriminator
 *
 * @param {{ manifest: any; expectedVersion: string; assetNames: string[]; expectedTargets: ExpectedTarget[] }} opts
 */
function validateReleaseExpectations({ manifest, expectedVersion, assetNames, expectedTargets }) {
  if (!expectedTargets || expectedTargets.length === 0) return;

  const version = normalizeVersion(expectedVersion);

  /** @type {string[]} */
  const errors = [];

  // 1) Installer presence per expected target.
  for (const target of expectedTargets) {
    const matches = findInstallerAssetsForTarget(assetNames, target, version);
    if (matches.length > 0) continue;

    const archTokens =
      target.arch === "universal"
        ? [
            ...ARCH_TOKENS.universal,
            ...(target.allowMissingArchInInstallerName ? ["(or no arch token; universal installer)"] : []),
          ]
        : ARCH_TOKENS[target.arch];

    const osInstallers = assetNames
      .filter((name) => hasAnyExtension(name, target.installerExts))
      .filter((name) => containsVersion(name, version));

    errors.push(`[${target.id}] Missing installer asset.`);
    errors.push(
      `  Looked for: version "${version}" + arch token (${archTokens.join(", ")}) + extension (${target.installerExts.join(
        ", ",
      )})`,
    );
    errors.push(
      `  Found ${target.os} installer-like assets (version match): ${
        osInstallers.length > 0 ? osInstallers.join(", ") : "(none)"
      }`,
    );
  }

  // 2) Updater targets must be present in latest.json.platforms.
  const foundPlatforms = findPlatformsObject(manifest);
  const platforms = foundPlatforms?.platforms;
  if (!platforms) {
    errors.push(`latest.json is missing a "platforms" object; cannot validate expected updater targets.`);
  } else {
    const platformKeys = Object.keys(platforms).sort();
    for (const target of expectedTargets) {
      const hasKey = target.updaterPlatformKeys.some((key) =>
        Object.prototype.hasOwnProperty.call(platforms, key),
      );
      if (hasKey) continue;

      errors.push(`[${target.id}] Missing updater platform entry in latest.json.`);
      errors.push(`  Expected one of: ${target.updaterPlatformKeys.join(", ")}`);
      errors.push(`  Found keys: ${platformKeys.length > 0 ? platformKeys.join(", ") : "(none)"}`);
    }
  }

  // 3) Ambiguous asset names on multi-arch platforms.
  /** @type {Map<DesktopOs, DesktopArch[]>} */
  const expectedArchsByOs = new Map();
  for (const target of expectedTargets) {
    const current = expectedArchsByOs.get(target.os) ?? [];
    current.push(target.arch);
    expectedArchsByOs.set(target.os, current);
  }

  for (const [os, archs] of expectedArchsByOs.entries()) {
    const uniqueArchs = Array.from(new Set(archs));
    const isMultiArch = uniqueArchs.length > 1 || uniqueArchs.includes("universal");
    if (!isMultiArch) continue;

    const allowMissingArchToken = expectedTargets.some(
      (t) => t.os === os && t.allowMissingArchInInstallerName === true,
    );
    const ambiguous = findAmbiguousAssetsForOs(assetNames, os, uniqueArchs, version, allowMissingArchToken);
    if (ambiguous.length === 0) continue;

    errors.push(`[${os}] Ambiguous artifacts detected (missing arch/universal discriminator).`);
    errors.push(
      `  Multi-arch enabled for ${os}: ${uniqueArchs.join(", ")}. Expected assets to include one of: ${uniqueArchs
        .flatMap((a) => ARCH_TOKENS[a])
        .join(", ")}`,
    );
    if (allowMissingArchToken) {
      errors.push(
        `  Note: ${os} allows a missing arch token for universal installers, but assets that include *some* arch token must still match one of the expected tokens.`,
      );
    }
    for (const name of ambiguous) {
      errors.push(`  ${name}`);
    }
  }

  if (errors.length > 0) {
    throw new ActionableError("Desktop release asset verification failed: expectation errors.", errors);
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    usage();
    return;
  }

  const ghToken = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!ghToken) {
    throw new ActionableError("Missing env var: GITHUB_TOKEN / GH_TOKEN", [
      "Set GITHUB_TOKEN (recommended) or GH_TOKEN to a token that can read the repo's releases.",
      "In GitHub Actions, you can usually use: env: { GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }} }.",
    ]);
  }

  const tag = args.tag ?? process.env.GITHUB_REF_NAME;
  const repo = args.repo ?? process.env.GITHUB_REPOSITORY;
  if (!tag) {
    throw new ActionableError("Missing release tag.", [
      "Pass --tag <tag> (example: --tag v0.2.3) or set env GITHUB_REF_NAME.",
    ]);
  }
  if (!repo) {
    throw new ActionableError("Missing GitHub repo.", [
      "Pass --repo <owner/repo> (example: --repo wilson/formula) or set env GITHUB_REPOSITORY.",
    ]);
  }

  /** @type {any} */
  let config;
  try {
    config = JSON.parse(await readFile(tauriConfigPath, "utf8"));
  } catch (err) {
    throw new ActionableError(`Failed to read/parse ${tauriConfigRelativePath}.`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }
  const expectedVersion = requireNonEmptyString(
    config?.version,
    `${tauriConfigRelativePath} → "version"`
  );
  const updaterPubkeyBase64 = requireNonEmptyString(
    config?.plugins?.updater?.pubkey,
    `${tauriConfigRelativePath} → plugins.updater.pubkey`
  );
  if (updaterPubkeyBase64 === PLACEHOLDER_PUBKEY || updaterPubkeyBase64.includes("REPLACE_WITH")) {
    throw new ActionableError(`Invalid updater public key (placeholder).`, [
      `Expected ${tauriConfigRelativePath} → plugins.updater.pubkey to be the real updater public key (base64).`,
    ]);
  }

  /** @type {{ publicKeyBytes: Buffer; keyId: string | null }} */
  let parsedPubkey;
  try {
    parsedPubkey = parseTauriUpdaterPubkey(updaterPubkeyBase64);
  } catch (err) {
    throw new ActionableError(`Invalid updater public key.`, [
      `Failed to parse ${tauriConfigRelativePath} → plugins.updater.pubkey.`,
      err instanceof Error ? err.message : String(err),
    ]);
  }

  let publicKey;
  try {
    publicKey = ed25519PublicKeyFromRaw(parsedPubkey.publicKeyBytes);
  } catch (err) {
    throw new ActionableError(`Invalid updater public key.`, [
      `Failed to construct an Ed25519 public key from plugins.updater.pubkey.`,
      err instanceof Error ? err.message : String(err),
    ]);
  }

  // Resolve opt-in expectations (off by default; safe for forks with fewer targets).
  /** @type {string[]} */
  const expectedTargetIds = [...args.expectedTargets];
  if (args.expectationsFile) {
    const fileConfig = await loadExpectationsFile(args.expectationsFile);
    expectedTargetIds.push(...fileConfig.targets);
  }

  const normalizedTargetIds = expectedTargetIds.map((t) => t.trim()).filter(Boolean);
  const uniqueTargetIds = Array.from(new Set(normalizedTargetIds));
  const unknownTargets = uniqueTargetIds.filter((t) => !EXPECTED_TARGETS[t]);
  if (unknownTargets.length > 0) {
    throw new ActionableError("Unknown expectation target(s).", [
      ...unknownTargets,
      "Supported targets:",
      ...Object.keys(EXPECTED_TARGETS).map((t) => `  ${t}`),
    ]);
  }

  /** @type {ExpectedTarget[]} */
  const expectedTargets = uniqueTargetIds.map((t) => EXPECTED_TARGETS[t]);
  const expectationsEnabled = expectedTargets.length > 0;
  if (expectationsEnabled) {
    console.log(`Expectations enabled: ${expectedTargets.map((t) => t.id).join(", ")}`);

    // Guard: macOS cannot be both universal and per-arch at the same time.
    const macosUniversal = expectedTargets.some((t) => t.id === "macos-universal");
    const macosPerArch = expectedTargets.some((t) => t.id === "macos-x64" || t.id === "macos-arm64");
    if (macosUniversal && macosPerArch) {
      throw new ActionableError("Invalid expectations.", [
        "Cannot combine macos-universal with macos-x64/macos-arm64.",
        "Choose either a universal build (--expect-macos-universal) or per-arch builds (--expect-macos-x64 + --expect-macos-arm64).",
      ]);
    }
  }

  /** @type {any} */
  let release;
  /** @type {any[]} */
  let assets = [];
  /** @type {Map<string, any>} */
  let assetsByName = new Map();
  /** @type {string[]} */
  let assetNames = [];
  /** @type {any} */
  let manifest;

  // GitHub Releases can take a moment to reflect newly-uploaded assets (especially
  // during CI where platform jobs upload in parallel). Retry a few times so the
  // verifier can be run immediately after a release workflow completes.
  const retryDelaysMs = [2000, 4000, 8000, 12000, 20000];
  let lastError = null;

  for (let attempt = 0; attempt <= retryDelaysMs.length; attempt += 1) {
    try {
      release = await getReleaseByTag(repo, tag, ghToken);
      const releaseId = release.id;

      assets = await listAllReleaseAssets(repo, releaseId, ghToken);
      assetsByName = new Map();
      for (const asset of assets) {
        if (asset && typeof asset === "object" && typeof asset.name === "string") {
          assetsByName.set(asset.name, asset);
        }
      }

      assetNames = Array.from(assetsByName.keys()).sort();
      if (assetNames.length === 0) {
        throw new ActionableError(`No assets found on the GitHub Release for tag ${tag}.`, [
          "Ensure `.github/workflows/release.yml` completed successfully and uploaded artifacts.",
        ]);
      }

      const latestJsonAsset = assetsByName.get("latest.json");
      if (!latestJsonAsset) {
        throw new ActionableError(`Release is missing required asset: latest.json`, [
          `Tag: ${tag}`,
          `Repo: ${repo}`,
          `Assets present (${assetNames.length}): ${assetNames.join(", ")}`,
        ]);
      }

      const latestJsonSigAsset = assetsByName.get("latest.json.sig");
      if (!latestJsonSigAsset) {
        throw new ActionableError(`Release is missing required asset: latest.json.sig`, [
          "The Tauri updater manifest must be signed; ensure TAURI_PRIVATE_KEY/TAURI_KEY_PASSWORD are set in CI.",
          `Assets present (${assetNames.length}): ${assetNames.join(", ")}`,
        ]);
      }

      const latestJsonBytes = await downloadReleaseAssetBytes(latestJsonAsset, ghToken);
      const latestJsonText = latestJsonBytes.toString("utf8");
      try {
        manifest = JSON.parse(latestJsonText);
      } catch (err) {
        throw new ActionableError(`Failed to parse latest.json as JSON.`, [
          err instanceof Error ? err.message : String(err),
        ]);
      }

      // Download latest.json.sig to ensure it's actually readable (not just present in the listing).
      const latestSigText = await downloadReleaseAssetText(latestJsonSigAsset, ghToken);
      if (latestSigText.trim().length === 0) {
        throw new ActionableError(`latest.json.sig downloaded successfully but was empty.`, [
          "This likely indicates an upstream signing/upload failure.",
        ]);
      }

      /** @type {{ signatureBytes: Buffer; keyId: string | null }} */
      let parsedSig;
      try {
        parsedSig = parseTauriUpdaterSignature(latestSigText, "latest.json.sig");
      } catch (err) {
        throw new ActionableError(`Failed to parse latest.json.sig as a Tauri updater signature.`, [
          err instanceof Error ? err.message : String(err),
        ]);
      }

      if (parsedSig.keyId && parsedPubkey.keyId && parsedSig.keyId !== parsedPubkey.keyId) {
        throw new ActionableError(`Updater manifest signature key id mismatch.`, [
          `latest.json.sig uses key id ${parsedSig.keyId}, but plugins.updater.pubkey is ${parsedPubkey.keyId}.`,
          `This usually means TAURI_PRIVATE_KEY does not correspond to the committed plugins.updater.pubkey.`,
        ]);
      }

      const signatureOk = verify(null, latestJsonBytes, publicKey, parsedSig.signatureBytes);
      if (!signatureOk) {
        throw new ActionableError(`Updater manifest signature verification failed.`, [
          `latest.json.sig does not verify latest.json using the updater public key embedded in ${tauriConfigRelativePath}.`,
          `This usually means the manifest/signature were generated with a different key, or assets were tampered with.`,
        ]);
      }
      validateLatestJson(manifest, expectedVersion, assetsByName);

      if (expectationsEnabled) {
        validateReleaseExpectations({
          manifest,
          expectedVersion,
          assetNames,
          expectedTargets,
        });
      }

      // If we get here, manifest + asset cross-check has passed.
      console.log(
        `Verified latest.json against desktop version ${expectedVersion} and release assets (${assetNames.length} total).`
      );
      lastError = null;
      break;
    } catch (err) {
      lastError = err;
      if (attempt === retryDelaysMs.length) break;

      const ms = retryDelaysMs[attempt];
      const brief = err instanceof Error ? err.message.split("\n")[0] : String(err);
      console.log(
        `Release assets not ready yet (${attempt + 1}/${retryDelaysMs.length + 1}): ${brief}`
      );
      console.log(`Retrying in ${Math.round(ms / 1000)}s...`);
      await sleep(ms);
    }
  }

  if (lastError) {
    throw lastError;
  }

  if (args.verifyAssets) {
    await verifyUpdaterPlatformAssetSignatures({
      manifest,
      assetsByName,
      token: ghToken,
      publicKey,
      pubkeyKeyId: parsedPubkey.keyId,
    });
  }

  if (args.dryRun) {
    console.log("Dry-run enabled: skipping SHA256SUMS generation.");
    return;
  }

  const outPath = path.resolve(process.cwd(), args.out ?? "SHA256SUMS.txt");

  const primaryAssets = assets
    .filter((asset) => asset && typeof asset === "object")
    .filter(
      (asset) =>
        typeof asset.name === "string" &&
        (typeof asset.url === "string" || typeof asset.browser_download_url === "string")
    )
    .filter((asset) => isPrimaryBundleOrSig(asset.name, { includeSigs: args.includeSigs }))
    .sort((a, b) => a.name.localeCompare(b.name));

  if (primaryAssets.length === 0) {
    throw new ActionableError("No primary bundle assets found to hash.", [
      "Expected at least one of: .dmg, .tar.gz, .msi, .exe, .AppImage, .deb, .rpm, .zip, .pkg",
      `Assets present (${assetNames.length}): ${assetNames.join(", ")}`,
    ]);
  }

  const primaryBundleCount = primaryAssets.filter((a) => isPrimaryBundleAssetName(a.name)).length;
  if (primaryBundleCount === 0) {
    throw new ActionableError("No primary bundle assets found to hash.", [
      "Only signature files were found (.sig).",
      "Expected at least one of: .dmg, .tar.gz, .msi, .exe, .AppImage, .deb, .rpm, .zip, .pkg",
      `Assets present (${assetNames.length}): ${assetNames.join(", ")}`,
    ]);
  }

  const assetsToHash = args.allAssets
    ? assets
        .filter((asset) => asset && typeof asset === "object")
        .filter(
          (asset) =>
            typeof asset.name === "string" &&
            (typeof asset.url === "string" || typeof asset.browser_download_url === "string")
        )
        .filter((asset) => asset.name !== "SHA256SUMS.txt")
        .filter((asset) => args.includeSigs || !asset.name.endsWith(".sig"))
        .sort((a, b) => a.name.localeCompare(b.name))
    : primaryAssets;

  /** @type {string[]} */
  const lines = [];
  for (let i = 0; i < assetsToHash.length; i++) {
    const asset = assetsToHash[i];
    const size = typeof asset.size === "number" ? asset.size : undefined;
    console.log(
      `Hashing (${i + 1}/${assetsToHash.length}) ${asset.name}${size ? ` (${formatBytes(size)})` : ""}...`
    );
    const digest = await sha256OfReleaseAsset(asset, ghToken);
    lines.push(`${digest}  ${asset.name}`);
  }

  await mkdir(path.dirname(outPath), { recursive: true });
  await writeFile(outPath, `${lines.join("\n")}\n`, "utf8");

  console.log(`Wrote SHA256SUMS for ${assetsToHash.length} assets → ${outPath}`);
  console.log("Desktop release asset verification passed.");
}

const isMainModule =
  typeof process.argv[1] === "string" &&
  path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);

if (isMainModule) {
  try {
    await main();
  } catch (err) {
    if (err instanceof ActionableError) {
      console.error(`\n${err.message}\n`);
    } else {
      console.error(err);
    }
    process.exit(1);
  }
}

export {
  ActionableError,
  filenameFromUrl,
  findPlatformsObject,
  isPrimaryBundleOrSig,
  isPrimaryBundleAssetName,
  normalizeVersion,
  validateReleaseExpectations,
  verifyUpdaterManifestSignature,
  validateLatestJson,
};
