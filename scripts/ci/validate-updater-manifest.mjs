#!/usr/bin/env node
/**
 * CI guard for the desktop release workflow:
 * - Downloads the combined Tauri updater manifest (latest.json + latest.json.sig) from the draft release.
 * - Ensures the manifest version matches the git tag.
 * - Ensures the manifest contains updater entries for all expected targets.
 * - Ensures each updater entry references an asset that exists on the GitHub Release.
 *
 * This catches "last writer wins" / merge regressions where one platform build overwrites latest.json
 * and ships an incomplete updater manifest.
 */
import { spawnSync } from "node:child_process";
import { readFileSync, writeFileSync, statSync } from "node:fs";
import process from "node:process";
import { setTimeout as sleep } from "node:timers/promises";
import { URL } from "node:url";
import crypto from "node:crypto";

/**
 * @param {string} message
 */
function fatal(message) {
  console.error(message);
  process.exit(1);
}

/**
 * @param {string[]} args
 * @returns {string}
 */
function gh(args) {
  const res = spawnSync("gh", args, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
  });
  if (res.error) {
    throw res.error;
  }
  if (res.status !== 0) {
    const cmd = ["gh", ...args].join(" ");
    throw new Error(`${cmd} failed (exit ${res.status}).\n${res.stderr}`.trim());
  }
  return res.stdout;
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
  const url = `https://api.github.com/repos/${repo}/releases/tags/${encodeURIComponent(tag)}`;
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
  const perPage = 100;
  let page = 1;
  while (true) {
    const url = `https://api.github.com/repos/${repo}/releases/${releaseId}/assets?per_page=${perPage}&page=${page}`;
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
 * @param {string} key
 * @param {unknown} value
 */
function expectNonEmptyString(key, value) {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`Expected ${key} to be a non-empty string, got ${String(value)}`);
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

  // At minimum (per platform requirements). We allow a couple of aliases in case Tauri changes
  // how it formats target identifiers in the updater JSON.
  const expectedTargets = [
    {
      id: "darwin-universal",
      label: "macOS (universal)",
      keys: ["darwin-universal", "universal-apple-darwin"],
    },
    {
      id: "windows-x86_64",
      label: "Windows (x86_64)",
      keys: ["windows-x86_64", "x86_64-pc-windows-msvc"],
    },
    {
      id: "windows-arm64",
      label: "Windows (ARM64)",
      keys: ["windows-aarch64", "windows-arm64", "aarch64-pc-windows-msvc"],
    },
    {
      id: "linux-x86_64",
      label: "Linux (x86_64)",
      keys: ["linux-x86_64", "x86_64-unknown-linux-gnu"],
    },
  ];

  const retryDelaysMs = [2000, 4000, 8000, 12000, 20000];
  /** @type {any | undefined} */
  let release;
  /** @type {any[] | undefined} */
  let assets;

  for (let attempt = 0; attempt <= retryDelaysMs.length; attempt += 1) {
    try {
      release = await fetchRelease({ repo, tag, token });
      const releaseId = /** @type {number} */ (release?.id);
      if (!releaseId) {
        throw new Error(`Release payload missing id.`);
      }
      assets = await fetchAllReleaseAssets({ repo, releaseId, token });

      const names = new Set(assets.map((a) => a?.name).filter((n) => typeof n === "string"));
      if (names.has("latest.json") && names.has("latest.json.sig")) {
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

  const assetByName = new Map(
    assets
      .filter((a) => a && typeof a.name === "string")
      .map((a) => /** @type {[string, any]} */ ([a.name, a])),
  );
  const assetNames = new Set(assetByName.keys());

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

  // Download the manifest + signature using the GitHub API (works for draft releases and private repos).
  for (const [asset, fileName] of [
    [latestAsset, "latest.json"],
    [latestSigAsset, "latest.json.sig"],
  ]) {
    const assetUrl = typeof asset.url === "string" ? asset.url : "";
    if (!assetUrl) {
      fatal(`Release asset ${fileName} missing API url; cannot download.`);
    }

    const apiPath = new URL(assetUrl).pathname.replace(/^\/+/, "");
    gh(["api", "-H", "Accept: application/octet-stream", apiPath, "--output", fileName]);

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

  const manifestVersion = typeof manifest?.version === "string" ? manifest.version : "";
  if (!manifestVersion) {
    errors.push(`latest.json missing required "version" field.`);
  } else if (normalizeVersion(manifestVersion) !== expectedVersion) {
    errors.push(
      `latest.json version mismatch: expected ${JSON.stringify(expectedVersion)} (from tag ${tag}), got ${JSON.stringify(manifestVersion)}.`,
    );
  }

  const platforms = manifest?.platforms;
  if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
    errors.push(`latest.json missing required "platforms" object.`);
  }

  /** @type {Array<{ target: string; url: string; assetName: string }>} */
  const missingAssets = [];
  /** @type {Array<{ label: string; expectedKeys: string[] }>} */
  const missingTargets = [];
  /** @type {Array<{ target: string; message: string }>} */
  const invalidTargets = [];

  if (platforms && typeof platforms === "object" && !Array.isArray(platforms)) {
    /** @type {Array<{ target: string; url: string; assetName: string }>} */
    const validatedTargets = [];

    // Validate *every* platform entry, not just the required ones. If the manifest contains
    // stale/invalid targets we want to catch them too.
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

    const validatedByTarget = new Map(validatedTargets.map((t) => [t.target, t]));

    /** @type {Array<{ label: string; key: string; url: string; assetName: string }>} */
    const summaryRows = [];

    for (const expected of expectedTargets) {
      const foundKey = expected.keys.find((k) => Object.prototype.hasOwnProperty.call(platforms, k));
      if (!foundKey) {
        missingTargets.push({ label: expected.label, expectedKeys: expected.keys });
        continue;
      }

      const validated = validatedByTarget.get(foundKey);
      if (!validated) {
        // Entry exists (foundKey) but was invalid, so it will already be present in invalidTargets.
        continue;
      }
      summaryRows.push({
        label: expected.label,
        key: foundKey,
        url: validated.url,
        assetName: validated.assetName,
      });
    }

    // Print additional targets (if any) in the success summary.
    const expectedKeySet = new Set(expectedTargets.flatMap((t) => t.keys));
    const otherTargets = validatedTargets
      .filter((t) => !expectedKeySet.has(t.target))
      .sort((a, b) => a.target.localeCompare(b.target));

    // Success: print a short summary (also write to the GitHub Actions step summary if available).
    if (
      errors.length === 0 &&
      missingTargets.length === 0 &&
      invalidTargets.length === 0 &&
      missingAssets.length === 0
    ) {
      const summaryLines = [
        `Updater manifest validation passed for ${tag} (version ${expectedVersion}).`,
        `Targets present (${summaryRows.length}):`,
        ...summaryRows.map((row) => `  - ${row.key} → ${row.assetName}`),
        ...(otherTargets.length > 0
          ? [
              `Other targets present (${otherTargets.length}):`,
              ...otherTargets.map((t) => `  - ${t.target} → ${t.assetName}`),
            ]
          : []),
      ];

      console.log(summaryLines.join("\n"));

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
          ...summaryRows.map((row) => `- \`${row.key}\` → \`${row.assetName}\``),
          ...(otherTargets.length > 0
            ? [
                ``,
                `### Other targets`,
                ``,
                ...otherTargets.map((t) => `- \`${t.target}\` → \`${t.assetName}\``),
              ]
            : []),
          ``,
        ].join("\n");
        // Overwrite the step summary rather than append (the job is dedicated to validation).
        writeFileSync(stepSummaryPath, md, "utf8");
      }

      return;
    }
  }

  if (missingTargets.length > 0) {
    errors.push(
      [
        `Missing required platform targets in latest.json:`,
        ...missingTargets.map(
          (t) =>
            `  - ${t.label} (expected one of: ${t.expectedKeys.map((k) => JSON.stringify(k)).join(", ")})`,
        ),
      ].join("\n"),
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

  if (missingAssets.length > 0) {
    errors.push(
      [
        `latest.json references assets that are not present on the GitHub Release:`,
        ...missingAssets.map(
          (a) => `  - ${a.target}: ${a.assetName} (from ${JSON.stringify(a.url)})`,
        ),
      ].join("\n"),
    );
  }

  if (errors.length > 0) {
    const available = Array.from(assetNames).sort();
    fatal(
      [
        `Updater manifest validation failed for release ${tag}.`,
        "",
        ...errors.map((e) => `- ${e}`),
        "",
        `Release assets (${available.length}):`,
        ...available.map((name) => `  - ${name}`),
        "",
        `This usually indicates that one platform build overwrote latest.json instead of merging it.`,
        `Verify the release workflow uploads a combined updater manifest for all targets.`,
      ].join("\n"),
    );
  }
}

main().catch((err) => {
  fatal(err instanceof Error ? err.stack ?? err.message : String(err));
});
