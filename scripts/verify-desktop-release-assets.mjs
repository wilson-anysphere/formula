#!/usr/bin/env node
/**
 * Verify GitHub Release assets produced by `.github/workflows/release.yml` for
 * the desktop (Tauri) app.
 *
 * Validates:
 * - `latest.json` updater manifest structure + version correctness
 * - every `platforms[*].url` in `latest.json` points at an uploaded asset
 * - `latest.json.sig` exists (manifest signing)
 * - optional per-bundle signature coverage (either `signature` in JSON or a
 *   sibling `<bundle>.sig` release asset)
 *
 * Generates:
 * - `SHA256SUMS.txt` for primary bundle artifacts (excluding `.sig` files)
 *
 * Usage:
 *   node scripts/verify-desktop-release-assets.mjs --tag v0.1.0 --repo owner/repo
 *
 * Required env:
 *   GITHUB_TOKEN - token with access to the repo's releases
 */

import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

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
      "  --dry-run          Validate manifest/assets only (skip bundle hashing)",
      "  -h, --help         Show help",
      "",
      "Env:",
      "  GITHUB_TOKEN       GitHub token with permission to read release assets",
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
  /** @type {{ tag?: string; repo?: string; out?: string; dryRun: boolean; help: boolean }} */
  const parsed = { dryRun: false, help: false };

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
 * @param {{ name: string; browser_download_url: string }} asset
 * @param {string} token
 */
async function downloadReleaseAssetText(asset, token) {
  const res = await fetchGitHub(asset.browser_download_url, {
    token,
    accept: "application/octet-stream",
  });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`Failed to download release asset "${asset.name}".`, [
      `${res.status} ${res.statusText}`,
      body ? `Response: ${body}` : "Response body was empty.",
      `URL: ${asset.browser_download_url}`,
    ]);
  }
  return await res.text();
}

/**
 * @param {{ name: string; size?: number; browser_download_url: string }} asset
 * @param {string} token
 */
async function sha256OfReleaseAsset(asset, token) {
  const res = await fetchGitHub(asset.browser_download_url, {
    token,
    accept: "application/octet-stream",
  });
  if (!res.ok) {
    const body = await safeReadBody(res);
    throw new ActionableError(`Failed to download asset for hashing: "${asset.name}".`, [
      `${res.status} ${res.statusText}`,
      body ? `Response: ${body}` : "Response body was empty.",
      `URL: ${asset.browser_download_url}`,
    ]);
  }
  if (!res.body) {
    throw new ActionableError(`No response body while downloading "${asset.name}" for hashing.`);
  }

  const hash = createHash("sha256");
  // `res.body` is a WHATWG ReadableStream in Node 18+; it supports async iteration.
  for await (const chunk of res.body) {
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

  const platforms = manifest.platforms;
  if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
    errors.push(`latest.json is missing a "platforms" object.`);
  } else {
    const keys = Object.keys(platforms);
    const requiredOsSubstrings = ["linux", "windows", "darwin"];
    for (const os of requiredOsSubstrings) {
      if (!keys.some((k) => k.toLowerCase().includes(os))) {
        errors.push(
          `latest.json platforms is missing an entry containing "${os}". Present keys: ${keys.length > 0 ? keys.join(", ") : "(none)"}`
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

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    usage();
    return;
  }

  const token = process.env.GITHUB_TOKEN;
  if (!token) {
    throw new ActionableError("Missing env var: GITHUB_TOKEN", [
      "Set GITHUB_TOKEN to a token that can read the repo's releases.",
      "In GitHub Actions, you can usually use: env: { GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }} }",
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

  const release = await getReleaseByTag(repo, tag, token);
  const releaseId = release.id;

  const assets = await listAllReleaseAssets(repo, releaseId, token);
  const assetsByName = new Map();
  for (const asset of assets) {
    if (asset && typeof asset === "object" && typeof asset.name === "string") {
      assetsByName.set(asset.name, asset);
    }
  }

  const assetNames = Array.from(assetsByName.keys()).sort();
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

  const latestJsonText = await downloadReleaseAssetText(latestJsonAsset, token);
  /** @type {any} */
  let manifest;
  try {
    manifest = JSON.parse(latestJsonText);
  } catch (err) {
    throw new ActionableError(`Failed to parse latest.json as JSON.`, [
      err instanceof Error ? err.message : String(err),
    ]);
  }

  // Download latest.json.sig to ensure it's actually readable (not just present in the listing).
  const latestSigText = await downloadReleaseAssetText(latestJsonSigAsset, token);
  if (latestSigText.trim().length === 0) {
    throw new ActionableError(`latest.json.sig downloaded successfully but was empty.`, [
      "This likely indicates an upstream signing/upload failure.",
    ]);
  }

  validateLatestJson(manifest, expectedVersion, assetsByName);

  // If we get here, manifest + asset cross-check has passed.
  console.log(
    `Verified latest.json against desktop version ${expectedVersion} and release assets (${assetNames.length} total).`
  );

  if (args.dryRun) {
    console.log("Dry-run enabled: skipping SHA256SUMS generation.");
    return;
  }

  const outPath = path.resolve(process.cwd(), args.out ?? "SHA256SUMS.txt");

  const primaryAssets = assets
    .filter((asset) => asset && typeof asset === "object")
    .filter((asset) => typeof asset.name === "string" && typeof asset.browser_download_url === "string")
    .filter((asset) => !asset.name.endsWith(".sig"))
    .filter((asset) => isPrimaryBundleAssetName(asset.name))
    .sort((a, b) => a.name.localeCompare(b.name));

  if (primaryAssets.length === 0) {
    throw new ActionableError("No primary bundle assets found to hash.", [
      "Expected at least one of: .dmg, .tar.gz, .msi, .exe, .AppImage, .deb, .rpm, .zip, .pkg",
      `Assets present (${assetNames.length}): ${assetNames.join(", ")}`,
    ]);
  }

  /** @type {string[]} */
  const lines = [];
  for (let i = 0; i < primaryAssets.length; i++) {
    const asset = primaryAssets[i];
    const size = typeof asset.size === "number" ? asset.size : undefined;
    console.log(
      `Hashing (${i + 1}/${primaryAssets.length}) ${asset.name}${size ? ` (${formatBytes(size)})` : ""}...`
    );
    const digest = await sha256OfReleaseAsset(asset, token);
    lines.push(`${digest}  ${asset.name}`);
  }

  await mkdir(path.dirname(outPath), { recursive: true });
  await writeFile(outPath, `${lines.join("\n")}\n`, "utf8");

  console.log(`Wrote SHA256SUMS for ${primaryAssets.length} assets → ${outPath}`);
  console.log("Desktop release asset verification passed.");
}

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

