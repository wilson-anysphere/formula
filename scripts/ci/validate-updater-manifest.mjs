#!/usr/bin/env node
/**
 * CI guard for the desktop release workflow:
 * - Downloads the combined Tauri updater manifest (latest.json + latest.json.sig) from the draft release.
 * - Ensures the manifest version matches the git tag.
 * - Ensures the manifest contains updater entries for all expected targets.
 * - Ensures each updater entry references an asset that exists on the GitHub Release.
 * - Ensures each target references the correct *updatable* artifact type (macOS .app.tar.gz, Linux
 *   .AppImage, Windows .msi/.exe).
 * - Ensures all platform URLs are unique (no two targets colliding on the same asset URL).
 *
 * This catches "last writer wins" / merge regressions where one platform build overwrites latest.json
 * and ships an incomplete updater manifest.
 */
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
 * @param {string} url
 * @returns {string}
 */
function assetNameFromUrl(url) {
  const parsed = new URL(url);
  const last = parsed.pathname.split("/").filter(Boolean).pop() ?? "";
  return decodeURIComponent(last);
}

/**
 * @param {string} target
 * @returns {"macos" | "linux" | "windows" | null}
 */
function platformFamilyFromTarget(target) {
  const lower = target.toLowerCase();
  if (lower.includes("darwin") || lower.includes("apple-darwin") || lower.includes("macos")) {
    return "macos";
  }
  if (lower.includes("linux")) return "linux";
  if (lower.includes("windows") || lower.includes("pc-windows")) return "windows";
  return null;
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
 * Creates a Node.js public key object for an Ed25519 public key stored as raw bytes.
 *
 * Tauri stores updater keys as base64 strings, which decode to a 32-byte Ed25519 public key.
 * Node's crypto APIs expect a SPKI wrapper, so we construct:
 *   SubjectPublicKeyInfo  ::=  SEQUENCE  {
 *     algorithm         AlgorithmIdentifier,
 *     subjectPublicKey  BIT STRING
 *   }
 * where AlgorithmIdentifier is OID 1.3.101.112 (Ed25519).
 *
 * @param {Uint8Array} rawKey32
 */
function ed25519PublicKeyFromRaw(rawKey32) {
  if (rawKey32.length !== 32) {
    throw new Error(`Expected 32-byte Ed25519 public key, got ${rawKey32.length} bytes.`);
  }
  const spkiPrefix = Buffer.from("302a300506032b6570032100", "hex");
  const spkiDer = Buffer.concat([spkiPrefix, Buffer.from(rawKey32)]);
  return crypto.createPublicKey({ key: spkiDer, format: "der", type: "spki" });
}

/**
 * @param {Buffer} latestJsonBytes
 * @param {string} signatureText
 * @param {string} pubkeyBase64
 */
function verifyLatestJsonSignature(latestJsonBytes, signatureText, pubkeyBase64) {
  const signatureB64 = signatureText.trim();
  if (!signatureB64) {
    throw new Error("latest.json.sig is empty.");
  }

  /** @type {Buffer} */
  let signatureBytes;
  try {
    signatureBytes = Buffer.from(signatureB64, "base64");
  } catch (err) {
    throw new Error(
      `latest.json.sig is not valid base64 (${err instanceof Error ? err.message : String(err)}).`,
    );
  }

  if (signatureBytes.length !== 64) {
    throw new Error(
      `latest.json.sig decoded to ${signatureBytes.length} bytes (expected 64 for Ed25519 signature).`,
    );
  }

  /** @type {Buffer} */
  let pubkeyBytes;
  try {
    pubkeyBytes = Buffer.from(pubkeyBase64.trim(), "base64");
  } catch (err) {
    throw new Error(
      `Updater pubkey is not valid base64 (${err instanceof Error ? err.message : String(err)}).`,
    );
  }

  const publicKey = ed25519PublicKeyFromRaw(pubkeyBytes);
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

  // At minimum (per platform requirements). We allow a couple of aliases in case Tauri changes
  // how it formats target identifiers in the updater JSON.
  const macUniversalKeys = ["darwin-universal", "universal-apple-darwin"];
  const macX64Keys = ["darwin-x86_64", "x86_64-apple-darwin"];
  const macArm64Keys = ["darwin-aarch64", "aarch64-apple-darwin"];

  const expectedTargets = [
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
      const releaseId = /** @type {number} */ (release?.id);
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

  const platforms = manifest?.platforms;
  if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
    errors.push(`latest.json missing required "platforms" object.`);
  }

  /** @type {Array<{ target: string; url: string; assetName: string }>} */
  const missingAssets = [];
  /** @type {Array<{ label: string; expectation: string }>} */
  const missingTargets = [];
  /** @type {Array<{ target: string; message: string }>} */
  const invalidTargets = [];
  /** @type {Array<{ target: string; url: string; assetName: string }>} */
  const validatedTargets = [];

  if (platforms && typeof platforms === "object" && !Array.isArray(platforms)) {
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

    // Ensure platform URLs are unique (prevents collisions where multiple targets point at the same asset).
    const urlToTargets = new Map();
    for (const { target, url } of validatedTargets) {
      const list = urlToTargets.get(url) ?? [];
      list.push(target);
      urlToTargets.set(url, list);
    }

    const duplicateUrls = [...urlToTargets.entries()].filter(([, targets]) => targets.length > 1);
    if (duplicateUrls.length > 0) {
      errors.push(
        [
          `Duplicate platform URLs in latest.json (target collision):`,
          ...duplicateUrls
            .slice()
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([url, targets]) => `  - ${targets.slice().sort().join(", ")} → ${url}`),
        ].join("\n"),
      );
    }

    // Ensure each platform points at a self-updatable artifact type.
    /** @type {Array<{ target: string; url: string; expected: string }>} */
    const invalidMac = [];
    /** @type {Array<{ target: string; url: string; expected: string }>} */
    const invalidLinux = [];
    /** @type {Array<{ target: string; url: string; expected: string }>} */
    const invalidWindows = [];

    for (const { target, url } of validatedTargets) {
      const family = platformFamilyFromTarget(target);
      if (!family) continue;

      if (family === "macos") {
        if (!(url.endsWith(".app.tar.gz") || url.endsWith(".tar.gz"))) {
          invalidMac.push({
            target,
            url,
            expected: `ends with ".app.tar.gz" (preferred) or ".tar.gz"`,
          });
        }
      } else if (family === "linux") {
        if (!url.endsWith(".AppImage")) {
          invalidLinux.push({
            target,
            url,
            expected: `ends with ".AppImage" (Linux auto-update requires AppImage; .deb/.rpm are typically not self-updatable)`,
          });
        }
      } else if (family === "windows") {
        const lower = url.toLowerCase();
        if (!(lower.endsWith(".msi") || lower.endsWith(".exe"))) {
          invalidWindows.push({
            target,
            url,
            expected: `ends with ".msi" or ".exe"`,
          });
        }
      }
    }

    if (invalidMac.length > 0) {
      errors.push(
        [
          `Invalid macOS updater URLs in latest.json (expected updater archive):`,
          ...invalidMac
            .slice()
            .sort((a, b) => a.target.localeCompare(b.target))
            .map((t) => `  - ${t.target}: ${t.url} (${t.expected})`),
        ].join("\n"),
      );
    }

    if (invalidLinux.length > 0) {
      errors.push(
        [
          `Invalid Linux updater URLs in latest.json (expected .AppImage):`,
          ...invalidLinux
            .slice()
            .sort((a, b) => a.target.localeCompare(b.target))
            .map((t) => `  - ${t.target}: ${t.url} (${t.expected})`),
        ].join("\n"),
      );
    }

    if (invalidWindows.length > 0) {
      errors.push(
        [
          `Invalid Windows updater URLs in latest.json (expected .msi or .exe):`,
          ...invalidWindows
            .slice()
            .sort((a, b) => a.target.localeCompare(b.target))
            .map((t) => `  - ${t.target}: ${t.url} (${t.expected})`),
        ].join("\n"),
      );
    }

    const validatedByTarget = new Map(validatedTargets.map((t) => [t.target, t]));

    /** @type {Array<{ label: string; key: string; url: string; assetName: string }>} */
    const summaryRows = [];

    // macOS: either a dedicated universal key is present, OR both per-arch keys exist.
    const macUniversalKey = macUniversalKeys.find((k) =>
      Object.prototype.hasOwnProperty.call(platforms, k),
    );
    const macUniversalValidatedKey = macUniversalKeys.find((k) => validatedByTarget.has(k));

    if (macUniversalKey && !macUniversalValidatedKey) {
      // Key exists but was invalid; the invalid entry error should be enough, but keep the
      // missing-target message actionable about what we were expecting.
      missingTargets.push({
        label: "macOS (universal)",
        expectation: `expected a valid universal entry (${macUniversalKeys.map((k) => JSON.stringify(k)).join(", ")}) or both per-arch entries (${macX64Keys[0]} + ${macArm64Keys[0]}).`,
      });
    } else if (macUniversalValidatedKey) {
      const validated = validatedByTarget.get(macUniversalValidatedKey);
      if (validated) {
        summaryRows.push({
          label: "macOS (universal)",
          key: macUniversalValidatedKey,
          url: validated.url,
          assetName: validated.assetName,
        });
      }
    } else {
      const macX64Key = macX64Keys.find((k) => Object.prototype.hasOwnProperty.call(platforms, k));
      const macArm64Key = macArm64Keys.find((k) =>
        Object.prototype.hasOwnProperty.call(platforms, k),
      );

      const macX64ValidatedKey = macX64Keys.find((k) => validatedByTarget.has(k));
      const macArm64ValidatedKey = macArm64Keys.find((k) => validatedByTarget.has(k));

      if (!macX64Key || !macArm64Key) {
        missingTargets.push({
          label: "macOS (universal)",
          expectation: `expected ${macUniversalKeys.map((k) => JSON.stringify(k)).join(" or ")} or both ${macX64Keys.map((k) => JSON.stringify(k)).join(" / ")} and ${macArm64Keys.map((k) => JSON.stringify(k)).join(" / ")}.`,
        });
      } else if (!macX64ValidatedKey || !macArm64ValidatedKey) {
        // Keys exist but at least one was invalid.
        missingTargets.push({
          label: "macOS (universal)",
          expectation: `expected valid per-arch entries for both macOS x86_64 and arm64 (found ${JSON.stringify(macX64Key)} and ${JSON.stringify(macArm64Key)}).`,
        });
      } else {
        const validatedX64 = validatedByTarget.get(macX64ValidatedKey);
        const validatedArm64 = validatedByTarget.get(macArm64ValidatedKey);
        if (validatedX64) {
          summaryRows.push({
            label: "macOS (x86_64)",
            key: macX64ValidatedKey,
            url: validatedX64.url,
            assetName: validatedX64.assetName,
          });
        }
        if (validatedArm64) {
          summaryRows.push({
            label: "macOS (arm64)",
            key: macArm64ValidatedKey,
            url: validatedArm64.url,
            assetName: validatedArm64.assetName,
          });
        }
      }
    }

    for (const expected of expectedTargets) {
      const foundKey = expected.keys.find((k) => Object.prototype.hasOwnProperty.call(platforms, k));
      if (!foundKey) {
        missingTargets.push({
          label: expected.label,
          expectation: `expected one of: ${expected.keys.map((k) => JSON.stringify(k)).join(", ")}`,
        });
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
    const expectedKeySet = new Set([
      ...macUniversalKeys,
      ...macX64Keys,
      ...macArm64Keys,
      ...expectedTargets.flatMap((t) => t.keys),
    ]);
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
      const allRows = [
        ...summaryRows.map((row) => ({ target: row.key, assetName: row.assetName })),
        ...otherTargets.map((t) => ({ target: t.target, assetName: t.assetName })),
      ];
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
  }

  if (missingTargets.length > 0) {
    errors.push(
      [
        `Missing required platform targets in latest.json:`,
        ...missingTargets.map((t) => `  - ${t.label} (${t.expectation})`),
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
    const platformDebugLines =
      platforms && typeof platforms === "object" && !Array.isArray(platforms)
        ? [
            `Manifest platforms (${Object.keys(platforms).length}):`,
            ...validatedTargets
              .slice()
              .sort((a, b) => a.target.localeCompare(b.target))
              .map(
                (t) =>
                  `  - ${t.target} → ${t.assetName}${assetNames.has(t.assetName) ? "" : " (missing asset)"}`,
              ),
            ...invalidTargets
              .slice()
              .sort((a, b) => a.target.localeCompare(b.target))
              .map((t) => `  - ${t.target} → INVALID (${t.message})`),
            "",
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

main().catch((err) => {
  fatal(err instanceof Error ? err.stack ?? err.message : String(err));
});
