#!/usr/bin/env node
/**
 * CI guard for the desktop release workflow:
 * - Downloads the combined Tauri updater manifest (latest.json + latest.json.sig) from the draft release.
 * - Ensures the manifest version matches the git tag.
 * - Ensures the manifest contains updater entries for all expected targets.
 * - Ensures each updater entry references an asset that exists on the GitHub Release.
 * - Ensures each target references the correct *self-updatable* artifact type:
 *   - macOS: `.app.tar.gz` updater archive (not the `.dmg`)
 *   - Windows: `.msi` (Windows Installer; updater runs this)
 *   - Linux: `.AppImage`
 * - Ensures required `{os}-{arch}` targets do not unexpectedly collide on the same updater URL
 *   (macOS universal uses a single updater archive referenced by both `darwin-x86_64` and `darwin-aarch64`).
 *
 * This catches "last writer wins" / merge regressions where one platform build overwrites latest.json
 * and ships an incomplete updater manifest.
 */
import { readFileSync, writeFileSync, statSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { setTimeout as sleep } from "node:timers/promises";
import { URL, fileURLToPath } from "node:url";
import crypto from "node:crypto";
import {
  ed25519PublicKeyFromRaw,
  parseTauriUpdaterPubkey,
  parseTauriUpdaterSignature,
} from "./tauri-minisign.mjs";

/**
 * @param {string} message
 */
function fatal(message) {
  console.error(message);
  process.exit(1);
}

/**
 * @param {string} url
 * @returns {string}
 */
function assetNameFromUrl(url) {
  const parsed = new URL(url);
  const last = parsed.pathname.split("/").filter(Boolean).pop() ?? "";
  return decodeURIComponent(last);
}

/**
 * @param {{ target: string; assetName: string }[]} rows
 */
function formatTargetAssetTable(rows) {
  const sorted = rows.slice().sort((a, b) => a.target.localeCompare(b.target));
  const targetWidth = Math.max("target".length, ...sorted.map((r) => r.target.length));
  const assetWidth = Math.max("asset".length, ...sorted.map((r) => r.assetName.length));
  return [
    `${"target".padEnd(targetWidth)}  ${"asset".padEnd(assetWidth)}`,
    `${"-".repeat(targetWidth)}  ${"-".repeat(assetWidth)}`,
    ...sorted.map((r) => `${r.target.padEnd(targetWidth)}  ${r.assetName.padEnd(assetWidth)}`),
  ].join("\n");
}

/**
 * @param {{ target: string; assetName: string }[]} rows
 */
function formatTargetAssetMarkdownTable(rows) {
  const sorted = rows.slice().sort((a, b) => a.target.localeCompare(b.target));
  return [
    `| target | asset |`,
    `| --- | --- |`,
    ...sorted.map((r) => `| \`${r.target}\` | \`${r.assetName}\` |`),
  ].join("\n");
}

/**
 * Best-effort locate a Tauri updater `platforms` map within a JSON payload.
 *
 * Tauri v1/v2 manifests typically have `{ platforms: { ... } }` at the top-level, but we keep this
 * robust in case future versions nest the object.
 *
 * @param {unknown} root
 * @returns {{ platforms: Record<string, unknown>; path: string[] } | null}
 */
function findPlatformsObject(root) {
  if (!root || (typeof root !== "object" && !Array.isArray(root))) return null;

  /** @type {{ value: unknown; path: string[] }[]} */
  const queue = [{ value: root, path: [] }];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) break;

    const { value, path: currentPath } = current;
    if (value && typeof value === "object" && !Array.isArray(value)) {
      const obj = /** @type {Record<string, unknown>} */ (value);
      const platforms = obj.platforms;
      if (platforms && typeof platforms === "object" && !Array.isArray(platforms)) {
        return {
          platforms: /** @type {Record<string, unknown>} */ (platforms),
          path: [...currentPath, "platforms"],
        };
      }

      if (currentPath.length >= 8) continue;
      for (const [key, child] of Object.entries(obj)) {
        if (child && (typeof child === "object" || Array.isArray(child))) {
          queue.push({ value: child, path: [...currentPath, key] });
        }
      }
      continue;
    }

    if (Array.isArray(value)) {
      if (currentPath.length >= 8) continue;
      for (let i = 0; i < value.length; i += 1) {
        const child = value[i];
        if (child && (typeof child === "object" || Array.isArray(child))) {
          queue.push({ value: child, path: [...currentPath, String(i)] });
        }
      }
    }
  }

  return null;
}

// Platform key mapping is intentionally strict for the required `{os}-{arch}` updater keys.
//
// Source of truth:
// - docs/desktop-updater-target-mapping.md
//
// Do NOT accept "alias" keys here (like Rust target triples). If Tauri/tauri-action changes
// the platform key naming, we want this validator to fail loudly with an expected vs actual
// diff so we can update the docs + verification logic together.
const EXPECTED_PLATFORMS = [
  // macOS universal builds are published as a single `*.app.tar.gz`, but tauri-action writes it
  // under both arch keys so the updater can resolve updates on both Intel and Apple Silicon.
  {
    key: "darwin-x86_64",
    label: "macOS (x86_64)",
    expectedAsset: {
      description: `macOS updater archive (*.app.tar.gz)`,
      matches: (assetName) => assetName.toLowerCase().endsWith(".app.tar.gz"),
    },
  },
  {
    key: "darwin-aarch64",
    label: "macOS (aarch64)",
    expectedAsset: {
      description: `macOS updater archive (*.app.tar.gz)`,
      matches: (assetName) => assetName.toLowerCase().endsWith(".app.tar.gz"),
    },
  },
  {
    key: "windows-x86_64",
    label: "Windows (x64)",
    expectedAsset: {
      description: `Windows updater installer (*.msi)`,
      matches: (assetName) => assetName.toLowerCase().endsWith(".msi"),
    },
  },
  {
    key: "windows-aarch64",
    label: "Windows (ARM64)",
    expectedAsset: {
      description: `Windows updater installer (*.msi)`,
      matches: (assetName) => assetName.toLowerCase().endsWith(".msi"),
    },
  },
  {
    key: "linux-x86_64",
    label: "Linux (x86_64)",
    expectedAsset: {
      description: `Linux updater bundle (*.AppImage; not .deb/.rpm)`,
      matches: (assetName) => assetName.endsWith(".AppImage"),
    },
  },
  {
    key: "linux-aarch64",
    label: "Linux (ARM64)",
    expectedAsset: {
      description: `Linux updater bundle (*.AppImage; not .deb/.rpm)`,
      matches: (assetName) => assetName.endsWith(".AppImage"),
    },
  },
];

const EXPECTED_PLATFORM_KEYS = EXPECTED_PLATFORMS.map((p) => p.key);

/**
 * Validate the `platforms` section of a Tauri updater manifest (`latest.json`).
 *
 * Exported so node:test can cover the per-platform artifact type checks + collision guards without
 * needing to hit the GitHub API.
 *
 * Note: this function does not attempt to refresh/re-fetch release assets for eventual consistency.
 * The release workflow validator does that in `main()`.
 *
 * @param {{
 *   platforms: any,
 *   assetNames: Set<string>,
 * }} opts
 * @returns {{
 *   errors: string[],
 *   missingAssets: Array<{ target: string; url: string; assetName: string }>,
 *   invalidTargets: Array<{ target: string; message: string }>,
 *   validatedTargets: Array<{ target: string; url: string; assetName: string }>,
 * }}
 */
export function validatePlatformEntries({ platforms, assetNames }) {
  /** @type {string[]} */
  const errors = [];
  /** @type {Array<{ target: string; url: string; assetName: string }>} */
  const missingAssets = [];
  /** @type {Array<{ target: string; message: string }>} */
  const invalidTargets = [];
  /** @type {Array<{ target: string; url: string; assetName: string }>} */
  const validatedTargets = [];

  if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
    errors.push(`latest.json missing required "platforms" object.`);
    return { errors, missingAssets, invalidTargets, validatedTargets };
  }

  for (const [target, entry] of Object.entries(platforms)) {
    if (!entry || typeof entry !== "object") {
      invalidTargets.push({ target, message: "platform entry is not an object" });
      continue;
    }

    try {
      expectNonEmptyString(`${target}.url`, /** @type {any} */ (entry).url);
      expectNonEmptyString(`${target}.signature`, /** @type {any} */ (entry).signature);
    } catch (err) {
      invalidTargets.push({
        target,
        message: err instanceof Error ? err.message : String(err),
      });
      continue;
    }

    let assetName = "";
    try {
      assetName = assetNameFromUrl(/** @type {any} */ (entry).url);
    } catch (err) {
      invalidTargets.push({
        target,
        message: `url is not a valid URL (${err instanceof Error ? err.message : String(err)})`,
      });
      continue;
    }

    if (!assetNames.has(assetName)) {
      missingAssets.push({
        target,
        url: /** @type {any} */ (entry).url,
        assetName,
      });
    }

    validatedTargets.push({
      target,
      url: /** @type {any} */ (entry).url,
      assetName,
    });
  }

  const expectedKeySet = new Set(EXPECTED_PLATFORM_KEYS);

  // Ensure the manifest contains the required `{os}-{arch}` platform keys.
  // Note: `latest.json` may also contain additional installer-specific keys of the form
  // `{os}-{arch}-{bundle}` (for example `windows-x86_64-msi`). Those are allowed.
  const actualPlatformKeys = Object.keys(platforms).slice().sort();
  const expectedSortedKeys = EXPECTED_PLATFORM_KEYS.slice().sort();
  const missingKeys = expectedSortedKeys.filter(
    (k) => !Object.prototype.hasOwnProperty.call(platforms, k),
  );
  const otherKeys = actualPlatformKeys.filter((k) => !expectedKeySet.has(k));
  if (missingKeys.length > 0) {
    errors.push(
      [
        `Missing required latest.json.platforms keys (Tauri updater target identifiers).`,
        ``,
        `Required (${expectedSortedKeys.length}):`,
        ...expectedSortedKeys.map((k) => `  - ${k}`),
        ``,
        `Actual (${actualPlatformKeys.length}):`,
        ...actualPlatformKeys.map((k) => `  - ${k}`),
        ``,
        `Missing (${missingKeys.length}):`,
        ...missingKeys.map((k) => `  - ${k}`),
        ``,
        ...(otherKeys.length > 0
          ? [`Other keys present (${otherKeys.length}):`, ...otherKeys.map((k) => `  - ${k}`), ``]
          : []),
        `If you upgraded Tauri/tauri-action, update docs/desktop-updater-target-mapping.md and scripts/ci/validate-updater-manifest.mjs together.`,
      ].join("\n"),
    );
  }

  // Collision guards (required targets only).
  //
  // `latest.json` can legitimately contain multiple keys that point at the *same* asset URL:
  // - tauri-action writes both `{os}-{arch}` and `{os}-{arch}-{bundle}` variants
  //   (e.g. `windows-x86_64` and `windows-x86_64-msi` may point at the same `.msi`).
  // - macOS universal builds use a single updater archive but it is referenced by both
  //   `darwin-x86_64` and `darwin-aarch64`.
  //
  // We therefore only treat duplicates as an error when *different* required `{os}-{arch}` keys
  // (other than the macOS universal pair) collide.
  const allowedMacUniversalTargets = new Set(["darwin-x86_64", "darwin-aarch64"]);
  const requiredTargets = validatedTargets.filter((t) => expectedKeySet.has(t.target));

  const urlToRequiredTargets = new Map();
  for (const { target, url } of requiredTargets) {
    const list = urlToRequiredTargets.get(url) ?? [];
    list.push(target);
    urlToRequiredTargets.set(url, list);
  }

  const duplicateRequiredUrls = [...urlToRequiredTargets.entries()].filter(
    ([, targets]) => targets.length > 1,
  );
  const unexpectedDuplicateRequiredUrls = duplicateRequiredUrls.filter(([, targets]) => {
    return !targets.every((t) => allowedMacUniversalTargets.has(t));
  });
  if (unexpectedDuplicateRequiredUrls.length > 0) {
    errors.push(
      [
        `Duplicate updater URLs across required targets in latest.json (unexpected collision):`,
        ...unexpectedDuplicateRequiredUrls
          .slice()
          .sort((a, b) => a[0].localeCompare(b[0]))
          .map(([url, targets]) => `  - ${targets.slice().sort().join(", ")} → ${url}`),
      ].join("\n"),
    );
  }

  // Asset-name collision guard: same intent as URL check, but catches querystring/encoding differences.
  const assetNameToRequiredTargets = new Map();
  for (const { target, assetName } of requiredTargets) {
    const list = assetNameToRequiredTargets.get(assetName) ?? [];
    list.push(target);
    assetNameToRequiredTargets.set(assetName, list);
  }
  const duplicateRequiredAssets = [...assetNameToRequiredTargets.entries()].filter(
    ([, targets]) => targets.length > 1,
  );
  const unexpectedDuplicateRequiredAssets = duplicateRequiredAssets.filter(([, targets]) => {
    return !targets.every((t) => allowedMacUniversalTargets.has(t));
  });
  if (unexpectedDuplicateRequiredAssets.length > 0) {
    errors.push(
      [
        `Duplicate updater assets across required targets in latest.json (unexpected collision):`,
        ...unexpectedDuplicateRequiredAssets
          .slice()
          .sort((a, b) => a[0].localeCompare(b[0]))
          .map(([asset, targets]) => `  - ${targets.slice().sort().join(", ")} → ${asset}`),
      ].join("\n"),
    );
  }

  // Per-platform updater artifact type checks.
  const validatedByTarget = new Map(validatedTargets.map((t) => [t.target, t]));
  /** @type {Array<{ target: string; url: string; assetName: string; expected: string }>} */
  const wrongAssetTypes = [];
  for (const expected of EXPECTED_PLATFORMS) {
    const validated = validatedByTarget.get(expected.key);
    if (!validated) continue;
    if (!expected.expectedAsset.matches(validated.assetName)) {
      wrongAssetTypes.push({
        target: expected.key,
        url: validated.url,
        assetName: validated.assetName,
        expected: expected.expectedAsset.description,
      });
    }
  }

  if (wrongAssetTypes.length > 0) {
    errors.push(
      [
        `Updater asset type mismatch in latest.json.platforms:`,
        ...wrongAssetTypes
          .slice()
          .sort((a, b) => a.target.localeCompare(b.target))
          .map(
            (t) =>
              `  - ${t.target}: ${t.assetName} (from ${JSON.stringify(t.url)}; expected ${t.expected})`,
          ),
      ].join("\n"),
    );
  }

  // Guardrail for multi-arch Windows releases: ensure the updater entries reference *arch-specific*
  // assets (x64 vs arm64) and that the filenames include an arch token. This prevents multi-target
  // runs from clobbering assets on the GitHub Release (same name uploaded twice) and prevents
  // shipping a manifest that points the x64 updater target at an arm64 installer (or vice versa).
  const winX64Token = /(x64|x86[_-]64|amd64|win64)/i;
  const winArm64Token = /(arm64|aarch64)/i;
  /** @type {Array<{ target: string; assetName: string; expected: string }>} */
  const wrongWindowsAssetNames = [];

  for (const key of ["windows-x86_64", "windows-aarch64"]) {
    const validated = validatedByTarget.get(key);
    if (!validated) continue;
    const name = validated.assetName;
    const hasX64 = winX64Token.test(name);
    const hasArm64 = winArm64Token.test(name);

    if (key === "windows-x86_64") {
      if (!hasX64) {
        wrongWindowsAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename contains x64 token (x64/x86_64/x86-64/amd64/win64)`,
        });
      } else if (hasArm64) {
        wrongWindowsAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename does not contain arm64 token (arm64/aarch64)`,
        });
      }
    } else if (key === "windows-aarch64") {
      if (!hasArm64) {
        wrongWindowsAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename contains arm64 token (arm64/aarch64)`,
        });
      } else if (hasX64) {
        wrongWindowsAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename does not contain x64 token (x64/x86_64/x86-64/amd64/win64)`,
        });
      }
    }
  }

  if (wrongWindowsAssetNames.length > 0) {
    errors.push(
      [
        `Invalid Windows updater asset naming in latest.json.platforms (expected arch token in filename):`,
        ...wrongWindowsAssetNames
          .slice()
          .sort((a, b) => a.target.localeCompare(b.target))
          .map((t) => `  - ${t.target}: ${t.assetName} (${t.expected})`),
      ].join("\n"),
    );
  }

  // Guardrail for multi-arch Linux releases: ensure the updater entries reference *arch-specific*
  // assets (x86_64 vs aarch64) and that the filenames include an arch token. This prevents
  // multi-target runs from clobbering assets on the GitHub Release (same name uploaded twice) and
  // prevents shipping a manifest that points the x86_64 updater target at an ARM64 AppImage (or
  // vice versa).
  const linuxX64Token = /(x64|x86[_-]64|amd64)/i;
  const linuxArm64Token = /(arm64|aarch64)/i;
  /** @type {Array<{ target: string; assetName: string; expected: string }>} */
  const wrongLinuxAssetNames = [];

  for (const key of ["linux-x86_64", "linux-aarch64"]) {
    const validated = validatedByTarget.get(key);
    if (!validated) continue;
    const name = validated.assetName;
    const hasX64 = linuxX64Token.test(name);
    const hasArm64 = linuxArm64Token.test(name);

    if (key === "linux-x86_64") {
      if (!hasX64) {
        wrongLinuxAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename contains x86_64 token (x86_64/amd64/x64)`,
        });
      } else if (hasArm64) {
        wrongLinuxAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename does not contain arm64 token (arm64/aarch64)`,
        });
      }
    } else if (key === "linux-aarch64") {
      if (!hasArm64) {
        wrongLinuxAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename contains arm64 token (arm64/aarch64)`,
        });
      } else if (hasX64) {
        wrongLinuxAssetNames.push({
          target: key,
          assetName: name,
          expected: `filename does not contain x86_64 token (x86_64/amd64/x64)`,
        });
      }
    }
  }

  if (wrongLinuxAssetNames.length > 0) {
    errors.push(
      [
        `Invalid Linux updater asset naming in latest.json.platforms (expected arch token in filename):`,
        ...wrongLinuxAssetNames
          .slice()
          .sort((a, b) => a.target.localeCompare(b.target))
          .map((t) => `  - ${t.target}: ${t.assetName} (${t.expected})`),
      ].join("\n"),
    );
  }

  return { errors, missingAssets, invalidTargets, validatedTargets };
}

/**
 * @param {string} version
 * @returns {string}
 */
function normalizeVersion(version) {
  const trimmed = version.trim();
  return trimmed.startsWith("v") ? trimmed.slice(1) : trimmed;
}

/**
 * @param {{ repo: string; tag: string; token: string }}
 */
async function fetchRelease({ repo, tag, token }) {
  const apiBase = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/$/, "");
  const url = `${apiBase}/repos/${repo}/releases/tags/${encodeURIComponent(tag)}`;
  const res = await fetch(url, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`GET ${url} failed (${res.status}): ${text}`);
  }
  return /** @type {any} */ (await res.json());
}

/**
 * @param {{ repo: string; releaseId: number; token: string }}
 */
async function fetchAllReleaseAssets({ repo, releaseId, token }) {
  /** @type {any[]} */
  const assets = [];
  const apiBase = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/$/, "");
  const perPage = 100;
  let page = 1;
  while (true) {
    const url = `${apiBase}/repos/${repo}/releases/${releaseId}/assets?per_page=${perPage}&page=${page}`;
    const res = await fetch(url, {
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": "2022-11-28",
      },
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`GET ${url} failed (${res.status}): ${text}`);
    }
    const pageAssets = /** @type {any[]} */ (await res.json());
    assets.push(...pageAssets);
    const link = res.headers.get("link") ?? "";
    if (!link.includes('rel="next"') || pageAssets.length < perPage) {
      break;
    }
    page += 1;
  }
  return assets;
}

/**
 * @param {any[]} assets
 */
function indexAssetsByName(assets) {
  const map = new Map(
    assets
      .filter((a) => a && typeof a.name === "string")
      .map((a) => /** @type {[string, any]} */ ([a.name, a])),
  );
  return { map, names: new Set(map.keys()) };
}

/**
 * Download a GitHub Release asset via the GitHub API (supports draft releases).
 *
 * @param {{ asset: any; fileName: string; token: string }}
 */
async function downloadReleaseAsset({ asset, fileName, token }) {
  const assetUrl = typeof asset.url === "string" ? asset.url : "";
  if (!assetUrl) {
    throw new Error(`Release asset ${fileName} missing API url; cannot download.`);
  }

  const res = await fetch(assetUrl, {
    headers: {
      Accept: "application/octet-stream",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
    redirect: "follow",
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to download ${fileName} (${res.status}): ${text}`);
  }

  const bytes = Buffer.from(await res.arrayBuffer());
  writeFileSync(fileName, bytes);
}

/**
 * @param {string} key
 * @param {unknown} value
 */
function expectNonEmptyString(key, value) {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`Expected ${key} to be a non-empty string, got ${String(value)}`);
  }
}

/**
 * @param {Buffer} latestJsonBytes
 * @param {string} signatureText
 * @param {string} pubkeyBase64
 */
function verifyLatestJsonSignature(latestJsonBytes, signatureText, pubkeyBase64) {
  const { signatureBytes, keyId: sigKeyId } = parseTauriUpdaterSignature(signatureText, "latest.json.sig");
  const { publicKeyBytes, keyId: pubKeyId } = parseTauriUpdaterPubkey(pubkeyBase64);

  if (sigKeyId && pubKeyId && sigKeyId !== pubKeyId) {
    throw new Error(
      `latest.json.sig key id mismatch: signature uses ${sigKeyId}, but updater pubkey is ${pubKeyId}.`,
    );
  }

  const publicKey = ed25519PublicKeyFromRaw(publicKeyBytes);
  const ok = crypto.verify(null, latestJsonBytes, publicKey, signatureBytes);
  if (!ok) {
    throw new Error(
      `latest.json.sig does not verify latest.json with the configured updater public key. This typically means latest.json and latest.json.sig were generated/uploaded inconsistently (race/overwrite).`,
    );
  }
}

async function main() {
  const refName = process.argv[2] ?? process.env.GITHUB_REF_NAME;
  if (!refName) {
    fatal(
      "Missing tag name. Usage: node scripts/ci/validate-updater-manifest.mjs <tag> (example: v0.2.3)",
    );
  }

  const normalizedRefName = refName.startsWith("refs/tags/")
    ? refName.slice("refs/tags/".length)
    : refName;
  const tag = normalizedRefName;
  const expectedVersion = normalizeVersion(normalizedRefName);

  const repo = process.env.GITHUB_REPOSITORY;
  if (!repo) {
    fatal("Missing GITHUB_REPOSITORY (expected to run inside GitHub Actions).");
  }

  const token = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!token) {
    fatal("Missing GITHUB_TOKEN / GH_TOKEN (required to query/download draft release assets).");
  }

  const retryDelaysMs = [2000, 4000, 8000, 12000, 20000];
  /** @type {any | undefined} */
  let release;
  /** @type {any[] | undefined} */
  let assets;
  /** @type {number | undefined} */
  let releaseId;

  /**
   * Linux release requirements:
   * - Tagged releases must ship .AppImage + .deb + .rpm artifacts.
   * - Each artifact must have a corresponding Tauri updater signature file: <artifact>.sig
   *
   * Note: these `.sig` files are *not* RPM/DEB GPG signatures; they are Ed25519 signatures used by
   * Tauri's updater.
   */
  const linuxRequiredArtifacts = [
    { ext: ".AppImage", label: "Linux bundle (.AppImage)" },
    { ext: ".deb", label: "Linux package (.deb)" },
    { ext: ".rpm", label: "Linux package (.rpm)" },
  ];

  for (let attempt = 0; attempt <= retryDelaysMs.length; attempt += 1) {
    try {
      release = await fetchRelease({ repo, tag, token });
      releaseId = /** @type {number} */ (release?.id);
      if (!releaseId) {
        throw new Error(`Release payload missing id.`);
      }
      assets = await fetchAllReleaseAssets({ repo, releaseId, token });

      const names = new Set(assets.map((a) => a?.name).filter((n) => typeof n === "string"));
      const nameList = Array.from(names);
      const linuxMissing = [];
      const linuxMissingSigs = [];
      for (const req of linuxRequiredArtifacts) {
        const matches = nameList.filter((n) => n.endsWith(req.ext));
        if (matches.length === 0) {
          linuxMissing.push(req.label);
          continue;
        }
        for (const match of matches) {
          if (!names.has(`${match}.sig`)) {
            linuxMissingSigs.push(`${match}.sig`);
          }
        }
      }

      if (
        names.has("latest.json") &&
        names.has("latest.json.sig") &&
        linuxMissing.length === 0 &&
        linuxMissingSigs.length === 0
      ) {
        break;
      }

      if (attempt === retryDelaysMs.length) {
        break;
      }
    } catch (err) {
      if (attempt === retryDelaysMs.length) {
        throw err;
      }
    }

    await sleep(retryDelaysMs[attempt]);
  }

  if (!release || !assets) {
    fatal(`Failed to fetch release info for tag ${tag}.`);
  }
  if (!releaseId) {
    fatal(`Failed to determine release id for tag ${tag}.`);
  }

  let { map: assetByName, names: assetNames } = indexAssetsByName(assets);

  const latestAsset = assetByName.get("latest.json");
  const latestSigAsset = assetByName.get("latest.json.sig");

  if (!latestAsset || !latestSigAsset) {
    const available = Array.from(assetNames).sort();
    fatal(
      [
        `Updater manifest validation failed: release ${tag} is missing updater manifest assets.`,
        "",
        `Expected assets:`,
        `  - latest.json`,
        `  - latest.json.sig`,
        "",
        `Available assets (${available.length}):`,
        ...available.map((name) => `  - ${name}`),
        "",
        `If this is a freshly-created draft release, a platform build may have failed before uploading the updater manifest.`,
      ].join("\n"),
    );
  }

  // Guardrail: ensure Linux release artifacts are present on the GitHub Release (not just produced
  // locally in the Linux build job).
  /** @type {string[]} */
  const linuxArtifactFailures = [];
  for (const req of linuxRequiredArtifacts) {
    const matches = Array.from(assetNames).filter((n) => n.endsWith(req.ext));
    if (matches.length === 0) {
      linuxArtifactFailures.push(`Missing ${req.label} asset.`);
      continue;
    }
    for (const match of matches) {
      const sigName = `${match}.sig`;
      if (!assetNames.has(sigName)) {
        linuxArtifactFailures.push(`Missing signature asset: ${sigName}`);
      }
    }
  }
  if (linuxArtifactFailures.length > 0) {
    const available = Array.from(assetNames).sort();
    fatal(
      [
        `Linux release artifact validation failed for release ${tag}.`,
        "",
        ...linuxArtifactFailures.map((m) => `- ${m}`),
        "",
        `Release assets (${available.length}):`,
        ...available.map((name) => `  - ${name}`),
      ].join("\n"),
    );
  }

  // Download the manifest + signature using the GitHub API (works for draft releases and private repos).
  for (const [asset, fileName] of [
    [latestAsset, "latest.json"],
    [latestSigAsset, "latest.json.sig"],
  ]) {
    await downloadReleaseAsset({ asset, fileName, token });

    try {
      const stats = statSync(fileName);
      if (stats.size === 0) {
        fatal(`Downloaded ${fileName} is empty (0 bytes).`);
      }
    } catch (err) {
      fatal(`Failed to stat downloaded ${fileName}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  /** @type {any} */
  let manifest;
  try {
    manifest = JSON.parse(readFileSync("latest.json", "utf8"));
  } catch (err) {
    fatal(
      [
        `Updater manifest validation failed: could not parse latest.json from release ${tag}.`,
        `Error: ${err instanceof Error ? err.message : String(err)}`,
      ].join("\n"),
    );
  }

  /** @type {string[]} */
  const errors = [];
  /** @type {{ keyId: string | null; publicKeyBytes: Buffer } | null} */
  let updaterPubkey = null;

  const manifestVersion = typeof manifest?.version === "string" ? manifest.version : "";
  if (!manifestVersion) {
    errors.push(`latest.json missing required "version" field.`);
  } else if (normalizeVersion(manifestVersion) !== expectedVersion) {
    errors.push(
      `latest.json version mismatch: expected ${JSON.stringify(expectedVersion)} (from tag ${tag}), got ${JSON.stringify(manifestVersion)}.`,
    );
  }

  // Verify the manifest signature file matches latest.json. This catches a particularly nasty
  // failure mode where concurrent jobs upload mismatched latest.json/latest.json.sig pairs.
  try {
    const tauriConfigPath = "apps/desktop/src-tauri/tauri.conf.json";
    const tauriConfig = JSON.parse(readFileSync(tauriConfigPath, "utf8"));
    const pubkey = tauriConfig?.plugins?.updater?.pubkey;
    if (typeof pubkey !== "string" || pubkey.trim().length === 0) {
      errors.push(
        `Cannot verify latest.json.sig: missing plugins.updater.pubkey in ${tauriConfigPath}.`,
      );
    } else if (pubkey.trim().includes("REPLACE_WITH")) {
      // This should already be guarded by scripts/check-updater-config.mjs, but keep the
      // validator robust to future config changes.
      errors.push(
        `Cannot verify latest.json.sig: updater pubkey looks like a placeholder value in ${tauriConfigPath}.`,
      );
    } else {
      updaterPubkey = parseTauriUpdaterPubkey(pubkey);
      verifyLatestJsonSignature(
        readFileSync("latest.json"),
        readFileSync("latest.json.sig", "utf8"),
        pubkey,
      );
    }
  } catch (err) {
    errors.push(
      `latest.json.sig verification failed: ${err instanceof Error ? err.message : String(err)}`,
    );
  }

  const platformsFound = findPlatformsObject(manifest);
  const platforms = platformsFound?.platforms;

  // Validate the per-platform updater entries (strict target keys + asset type checks).
  // This is extracted so node:test can cover the tricky parts without GitHub API calls.
  let platformValidation = validatePlatformEntries({ platforms, assetNames });

  // GitHub release assets can be eventually consistent right after upload. If the manifest
  // references an asset we can't see yet, re-fetch the asset list a few times before failing.
  if (platformValidation.missingAssets.length > 0) {
    const refreshDelaysMs = [2000, 4000, 8000];
    for (const delay of refreshDelaysMs) {
      await sleep(delay);
      try {
        assets = await fetchAllReleaseAssets({ repo, releaseId, token });
        ({ map: assetByName, names: assetNames } = indexAssetsByName(assets));
      } catch {
        // Ignore transient API errors; we'll fall back to the last-seen asset list.
      }

      platformValidation = validatePlatformEntries({ platforms, assetNames });
      if (platformValidation.missingAssets.length === 0) {
        break;
      }
    }
  }

  errors.push(...platformValidation.errors);

  // Optional: sanity-check the format of each per-platform signature string (base64 minisign / Ed25519).
  /** @type {Array<{ target: string; message: string }>} */
  const signatureFormatErrors = [];
  if (platforms && typeof platforms === "object" && !Array.isArray(platforms)) {
    for (const [target, entry] of Object.entries(platforms)) {
      if (!entry || typeof entry !== "object") continue;
      const signature = /** @type {any} */ (entry).signature;
      if (typeof signature !== "string") continue;
      try {
        const { keyId } = parseTauriUpdaterSignature(signature, `${target}.signature`);
        if (keyId && updaterPubkey?.keyId && keyId !== updaterPubkey.keyId) {
          signatureFormatErrors.push({
            target,
            message: `signature key id mismatch: expected ${updaterPubkey.keyId} but got ${keyId}`,
          });
        }
      } catch (err) {
        signatureFormatErrors.push({
          target,
          message: err instanceof Error ? err.message : String(err),
        });
      }
    }
  }

  const invalidTargets = [...platformValidation.invalidTargets, ...signatureFormatErrors];

  if (invalidTargets.length > 0) {
    errors.push(
      [
        `Invalid platform entries in latest.json:`,
        ...invalidTargets.map((t) => `  - ${t.target}: ${t.message}`),
      ].join("\n"),
    );
  }

  if (platformValidation.missingAssets.length > 0) {
    errors.push(
      [
        `latest.json references assets that are not present on the GitHub Release:`,
        ...platformValidation.missingAssets.map(
          (a) => `  - ${a.target}: ${a.assetName} (from ${JSON.stringify(a.url)})`,
        ),
      ].join("\n"),
    );
  }

  // Success summary.
  if (
    errors.length === 0 &&
    invalidTargets.length === 0 &&
    platformValidation.missingAssets.length === 0
  ) {
    const allRows = platformValidation.validatedTargets.map((t) => ({
      target: t.target,
      assetName: t.assetName,
    }));
    console.log(`Updater manifest validation passed for ${tag} (version ${expectedVersion}).`);
    console.log(`\nUpdater manifest target → asset:\n${formatTargetAssetTable(allRows)}\n`);

    const stepSummaryPath = process.env.GITHUB_STEP_SUMMARY;
    if (stepSummaryPath) {
      const sha = crypto.createHash("sha256").update(readFileSync("latest.json")).digest("hex");
      const sigSha = crypto
        .createHash("sha256")
        .update(readFileSync("latest.json.sig"))
        .digest("hex");

      const md = [
        `## Updater manifest validation`,
        ``,
        `- Tag: \`${tag}\``,
        `- Manifest version: \`${expectedVersion}\``,
        `- latest.json sha256: \`${sha}\``,
        `- latest.json.sig sha256: \`${sigSha}\``,
        ``,
        `### Targets`,
        ``,
        formatTargetAssetMarkdownTable(allRows),
        ``,
      ].join("\n");
      // Overwrite the step summary rather than append (the job is dedicated to validation).
      writeFileSync(stepSummaryPath, md, "utf8");
    }

    return;
  }

  if (errors.length > 0) {
    const available = Array.from(assetNames).sort();
    const platformDebugLines =
      platforms && typeof platforms === "object" && !Array.isArray(platforms)
        ? [
            `Manifest platforms (${Object.keys(platforms).length}):`,
            ...(platformValidation.validatedTargets.length > 0
              ? [
                  ``,
                  `Target → asset:`,
                  formatTargetAssetTable(
                    platformValidation.validatedTargets.map((t) => ({
                      target: t.target,
                      assetName: assetNames.has(t.assetName)
                        ? t.assetName
                        : `${t.assetName} (missing asset)`,
                    })),
                  ),
                ]
              : []),
            ...(invalidTargets.length > 0
              ? [
                  ``,
                  `Invalid target entries:`,
                  ...invalidTargets
                    .slice()
                    .sort((a, b) => a.target.localeCompare(b.target))
                    .map((t) => `  - ${t.target}: ${t.message}`),
                ]
              : []),
            ``,
          ]
        : [];
    fatal(
      [
        `Updater manifest validation failed for release ${tag}.`,
        "",
        ...errors.map((e) => `- ${e}`),
        "",
        ...platformDebugLines,
        `Release assets (${available.length}):`,
        ...available.map((name) => `  - ${name}`),
        "",
        `This usually indicates that one platform build overwrote latest.json instead of merging it.`,
        `Verify the release workflow uploads a combined updater manifest for all targets.`,
      ].join("\n"),
    );
  }
}

// Only execute the CLI when invoked as the entrypoint; allow importing this module from node:test.
if (path.resolve(process.argv[1] ?? "") === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    fatal(err instanceof Error ? err.stack ?? err.message : String(err));
  });
}
