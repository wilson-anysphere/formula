import crypto from "node:crypto";
import { isDeepStrictEqual } from "node:util";

/**
 * Normalize a version string so `vX.Y.Z`, `refs/tags/vX.Y.Z`, and `X.Y.Z` compare equal.
 *
 * @param {string} raw
 * @returns {string}
 */
export function normalizeVersion(raw) {
  if (typeof raw !== "string") {
    throw new TypeError(`Expected version to be a string, got ${typeof raw}`);
  }

  let v = raw.trim();
  if (v.startsWith("refs/tags/")) v = v.slice("refs/tags/".length);

  // Only strip a leading "v" when it looks like a semver prefix.
  if (v.startsWith("v") && /\d/.test(v.slice(1, 2))) v = v.slice(1);

  return v;
}

/**
 * @typedef {Record<string, any>} JsonObject
 *
 * @typedef {{
 *   version: string,
 *   notes?: string,
 *   pub_date?: string,
 *   platforms: Record<string, { url: string, signature?: string, [k: string]: any }>,
 *   [k: string]: any,
 * }} TauriUpdaterManifest
 */

/**
 * Merge multiple Tauri updater `latest.json` manifests into one multi-platform manifest.
 *
 * - Versions must match after normalization (`vX.Y.Z` vs `X.Y.Z`).
 * - Platform entries are merged by union of keys.
 * - Conflicting duplicate platform keys fail the merge.
 *
 * @param {TauriUpdaterManifest[]} manifests
 * @returns {TauriUpdaterManifest}
 */
export function mergeTauriUpdaterManifests(manifests) {
  if (!Array.isArray(manifests) || manifests.length === 0) {
    throw new Error("Expected one or more updater manifests to merge.");
  }

  const first = manifests[0];
  const firstVersion = normalizeVersion(first?.version ?? "");
  if (!firstVersion) throw new Error("Input manifest is missing a non-empty `version` field.");

  /** @type {TauriUpdaterManifest} */
  const merged = {
    ...structuredClone(first),
    // Normalize output version to avoid propagating `v` prefixes into updater metadata.
    version: firstVersion,
    platforms: {},
  };

  for (const [idx, manifest] of manifests.entries()) {
    if (!manifest || typeof manifest !== "object") {
      throw new Error(`Manifest[${idx}] is not an object.`);
    }

    const v = normalizeVersion(manifest.version ?? "");
    if (!v) throw new Error(`Manifest[${idx}] is missing a non-empty \`version\` field.`);
    if (v !== firstVersion) {
      throw new Error(
        `Manifest version mismatch: expected ${JSON.stringify(firstVersion)} but got ${JSON.stringify(v)} in manifest[${idx}].`,
      );
    }

    const platforms = manifest.platforms;
    if (!platforms || typeof platforms !== "object") {
      throw new Error(`Manifest[${idx}] is missing a \`platforms\` object.`);
    }

    for (const [platformKey, entry] of Object.entries(platforms)) {
      const existing = merged.platforms[platformKey];
      if (!existing) {
        merged.platforms[platformKey] = structuredClone(entry);
        continue;
      }

      if (!isDeepStrictEqual(existing, entry)) {
        const existingUrl = typeof existing?.url === "string" ? existing.url : String(existing?.url);
        const entryUrl = typeof entry?.url === "string" ? entry.url : String(entry?.url);
        throw new Error(
          `Conflicting platform entry for ${JSON.stringify(platformKey)}: existing url=${JSON.stringify(existingUrl)} vs new url=${JSON.stringify(entryUrl)}`,
        );
      }
    }
  }

  return merged;
}

/**
 * @typedef {{
 *   expectedVersion?: string,
 *   requiredPlatforms?: string[],
 * }} ValidateManifestOptions
 */

/**
 * Validate a merged Tauri updater manifest against our desktop release expectations.
 *
 * This intentionally encodes platform/asset-type rules so CI catches regressions without
 * requiring a live GitHub Release.
 *
 * @param {TauriUpdaterManifest} manifest
 * @param {ValidateManifestOptions} [opts]
 */
export function validateTauriUpdaterManifest(manifest, opts = {}) {
  /** @type {string[]} */
  const errors = [];

  if (!manifest || typeof manifest !== "object") {
    throw new Error("Manifest must be a JSON object.");
  }

  const rawVersion = typeof manifest.version === "string" ? manifest.version.trim() : "";
  if (!rawVersion) errors.push("Manifest is missing a non-empty `version` field.");

  const normalizedVersion = rawVersion ? normalizeVersion(rawVersion) : "";
  if (opts.expectedVersion) {
    const expected = normalizeVersion(opts.expectedVersion);
    if (normalizedVersion && normalizedVersion !== expected) {
      errors.push(
        `Manifest version mismatch: expected ${JSON.stringify(expected)} but got ${JSON.stringify(normalizedVersion)}.`,
      );
    }
  }

  const platforms = manifest.platforms;
  if (!platforms || typeof platforms !== "object") {
    errors.push("Manifest is missing a `platforms` object.");
  } else if (Object.keys(platforms).length === 0) {
    errors.push("Manifest `platforms` object is empty.");
  }

  if (platforms && typeof platforms === "object") {
    const required = Array.isArray(opts.requiredPlatforms) ? opts.requiredPlatforms : [];
    for (const key of required) {
      if (!(key in platforms)) errors.push(`Manifest is missing required platform ${JSON.stringify(key)}.`);
    }

    for (const [platformKey, entry] of Object.entries(platforms)) {
      const url = typeof entry?.url === "string" ? entry.url.trim() : "";
      if (!url) {
        errors.push(`platforms[${JSON.stringify(platformKey)}].url must be a non-empty string.`);
        continue;
      }

      const path = urlPathname(url);
      const lower = path.toLowerCase();
      const os = inferOsFromPlatformKey(platformKey);

      if (os === "darwin") {
        if (lower.endsWith(".dmg")) {
          errors.push(
            `macOS updater artifact must be an update archive (.app.tar.gz), not a DMG installer: ${JSON.stringify(url)}`,
          );
        } else if (!lower.endsWith(".app.tar.gz")) {
          errors.push(`macOS updater artifact must end with .app.tar.gz: ${JSON.stringify(url)}`);
        }
      } else if (os === "linux") {
        if (lower.endsWith(".deb")) {
          errors.push(
            `Linux updater artifact must be an AppImage archive (.AppImage.tar.gz), not a .deb installer: ${JSON.stringify(url)}`,
          );
        } else if (!lower.endsWith(".appimage.tar.gz")) {
          errors.push(`Linux updater artifact must end with .AppImage.tar.gz: ${JSON.stringify(url)}`);
        }
      } else if (os === "windows") {
        if (!(lower.endsWith(".msi.zip") || lower.endsWith(".exe.zip"))) {
          errors.push(`Windows updater artifact must end with .msi.zip or .exe.zip: ${JSON.stringify(url)}`);
        }
      }
    }
  }

  if (errors.length > 0) {
    throw new Error(`Updater manifest validation failed:\n${errors.map((e) => `- ${e}`).join("\n")}`);
  }
}

/**
 * Verify an Ed25519 signature over a UTF-8 manifest string using a base64 raw public key.
 *
 * The public key format matches `plugins.updater.pubkey` in `tauri.conf.json` (base64-encoded
 * raw Ed25519 public key bytes).
 *
 * @param {string} manifestText
 * @param {string} signatureBase64
 * @param {string} publicKeyBase64
 */
export function verifyTauriManifestSignature(manifestText, signatureBase64, publicKeyBase64) {
  if (typeof manifestText !== "string") throw new TypeError("manifestText must be a string");
  if (typeof signatureBase64 !== "string") throw new TypeError("signatureBase64 must be a string");
  if (typeof publicKeyBase64 !== "string") throw new TypeError("publicKeyBase64 must be a string");

  const signature = Buffer.from(signatureBase64.trim(), "base64");
  if (signature.length !== 64) {
    throw new Error(`Expected Ed25519 signature to be 64 bytes, got ${signature.length}.`);
  }

  const pubBytes = Buffer.from(publicKeyBase64.trim(), "base64");
  if (pubBytes.length !== 32) {
    throw new Error(`Expected Ed25519 public key to be 32 bytes (base64), got ${pubBytes.length}.`);
  }

  const jwk = { kty: "OKP", crv: "Ed25519", x: pubBytes.toString("base64url") };
  const key = crypto.createPublicKey({ key: jwk, format: "jwk" });

  return crypto.verify(null, Buffer.from(manifestText, "utf8"), key, signature);
}

/**
 * @param {string} platformKey
 * @returns {"darwin" | "linux" | "windows" | "unknown"}
 */
function inferOsFromPlatformKey(platformKey) {
  const lower = platformKey.toLowerCase();
  if (lower.startsWith("darwin-") || lower === "darwin") return "darwin";
  if (lower.startsWith("linux-") || lower === "linux") return "linux";
  if (lower.startsWith("windows-") || lower === "windows") return "windows";
  return "unknown";
}

/**
 * @param {string} url
 * @returns {string}
 */
function urlPathname(url) {
  try {
    return new URL(url).pathname;
  } catch {
    return url;
  }
}

