#!/usr/bin/env node
/**
 * Publish a combined Tauri updater manifest (latest.json + latest.json.sig) for a
 * release tag.
 *
 * Why:
 * - `tauri-apps/tauri-action` generates and uploads `latest.json` from each matrix
 *   job. When jobs run in parallel, these uploads can race ("last writer wins"),
 *   leaving `latest.json` missing targets.
 * - This script merges the per-job manifests (uploaded as workflow artifacts),
 *   then uploads a single combined `latest.json` and matching `latest.json.sig`.
 *
 * Usage:
 *   node scripts/ci/publish-updater-manifest.mjs <tag> <manifests-dir>
 *
 * Required env:
 *   - GITHUB_REPOSITORY
 *   - GITHUB_TOKEN (or GH_TOKEN)
 *   - TAURI_PRIVATE_KEY
 *   - TAURI_KEY_PASSWORD (required if TAURI_PRIVATE_KEY is encrypted)
 */
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import crypto from "node:crypto";
import {
  ed25519PrivateKeyFromSeed,
  parseMinisignSecretKeyPayload,
  parseMinisignSecretKeyText,
} from "./tauri-minisign.mjs";

/**
 * @param {string} message
 */
function fatal(message) {
  console.error(message);
  process.exit(1);
}

/**
 * @param {string} version
 */
function normalizeVersion(version) {
  const trimmed = version.trim();
  return trimmed.startsWith("v") ? trimmed.slice(1) : trimmed;
}

/**
 * Returns the decoded bytes if the input looks like a base64 string, otherwise `null`.
 * Supports base64url and unpadded base64.
 * @param {string} value
 */
function decodeBase64(value) {
  const normalized = value.trim().replace(/\s+/g, "");
  if (normalized.length === 0) return null;
  let base64 = normalized.replace(/-/g, "+").replace(/_/g, "/");
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) return null;
  const mod = base64.length % 4;
  if (mod === 1) return null;
  if (mod !== 0) base64 += "=".repeat(4 - mod);
  return Buffer.from(base64, "base64");
}

/**
 * Parse TAURI_PRIVATE_KEY formats supported by the release workflow:
 * - PEM PKCS#8 (encrypted or not)
 * - base64 PKCS#8 DER (encrypted or not)
 * - raw Ed25519 private key (32/64 bytes, base64/base64url)
 * - minisign secret key (raw text, base64-encoded text, or base64 payload line)
 *
 * @param {string} privateKeyText
 * @param {string} password
 */
function loadEd25519PrivateKey(privateKeyText, password) {
  const trimmed = privateKeyText.trim();
  if (!trimmed) throw new Error("TAURI_PRIVATE_KEY is empty.");

  const passphrase = password.trim().length > 0 ? password : undefined;

  if (trimmed.includes("-----BEGIN")) {
    return crypto.createPrivateKey({ key: trimmed, format: "pem", passphrase });
  }

  // Support minisign secret keys (as printed by `cargo tauri signer generate`). These are Ed25519
  // keys; for unencrypted keys we can derive a PKCS#8 Ed25519 private key from the 32-byte seed.
  if (trimmed.toLowerCase().includes("minisign secret key")) {
    const parsed = parseMinisignSecretKeyText(trimmed);
    if (parsed.encrypted) {
      throw new Error(
        `Encrypted minisign secret keys are not supported by publish-updater-manifest.mjs. Convert your key to PKCS#8 or use an unencrypted minisign secret key.`,
      );
    }
    return ed25519PrivateKeyFromSeed(parsed.secretKeyBytes.subarray(0, 32));
  }

  const decoded = decodeBase64(trimmed);
  if (!decoded) {
    throw new Error("TAURI_PRIVATE_KEY is not valid base64/base64url and is not PEM.");
  }

  // base64-encoded minisign secret key file
  {
    const decodedText = decoded.toString("utf8");
    if (decodedText.toLowerCase().includes("minisign secret key")) {
      const parsed = parseMinisignSecretKeyText(decodedText);
      if (parsed.encrypted) {
        throw new Error(
          `Encrypted minisign secret keys are not supported by publish-updater-manifest.mjs. Convert your key to PKCS#8 or use an unencrypted minisign secret key.`,
        );
      }
      return ed25519PrivateKeyFromSeed(parsed.secretKeyBytes.subarray(0, 32));
    }
  }

  // base64-encoded minisign secret key binary payload (starts with "Ed")
  if (decoded.length >= 74 && decoded[0] === 0x45 && decoded[1] === 0x64) {
    const parsed = parseMinisignSecretKeyPayload(trimmed);
    if (parsed.encrypted) {
      throw new Error(
        `Encrypted minisign secret keys are not supported by publish-updater-manifest.mjs. Convert your key to PKCS#8 or use an unencrypted minisign secret key.`,
      );
    }
    return ed25519PrivateKeyFromSeed(parsed.secretKeyBytes.subarray(0, 32));
  }

  // Raw Ed25519 secret key (seed) or seed+public (libsodium style).
  if (decoded.length === 32 || decoded.length === 64) {
    const seed = decoded.subarray(0, 32);
    return ed25519PrivateKeyFromSeed(seed);
  }

  // Assume DER-encoded PKCS#8.
  return crypto.createPrivateKey({ key: decoded, format: "der", type: "pkcs8", passphrase });
}

/**
 * @param {string} dir
 * @returns {string[]}
 */
function findJsonFiles(dir) {
  /** @type {string[]} */
  const out = [];
  /** @type {string[]} */
  const stack = [dir];
  while (stack.length > 0) {
    const cur = stack.pop();
    if (!cur) break;
    let entries;
    try {
      entries = fs.readdirSync(cur, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const ent of entries) {
      const full = path.join(cur, ent.name);
      if (ent.isDirectory()) {
        stack.push(full);
      } else if (ent.isFile() && ent.name.endsWith(".json")) {
        out.push(full);
      }
    }
  }
  return out.slice().sort();
}

/**
 * @param {string} repo
 * @param {string} tag
 * @param {string} token
 */
async function fetchRelease(repo, tag, token) {
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
 * @param {string} repo
 * @param {number} releaseId
 * @param {string} token
 */
async function fetchAllReleaseAssets(repo, releaseId, token) {
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
    if (!link.includes('rel="next"') || pageAssets.length < perPage) break;
    page += 1;
  }
  return assets;
}

/**
 * @param {string} repo
 * @param {number} assetId
 * @param {string} token
 */
async function deleteReleaseAsset(repo, assetId, token) {
  const url = `https://api.github.com/repos/${repo}/releases/assets/${assetId}`;
  const res = await fetch(url, {
    method: "DELETE",
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`DELETE ${url} failed (${res.status}): ${text}`);
  }
}

/**
 * @param {string} uploadUrlTemplate
 * @param {string} name
 */
function releaseUploadUrl(uploadUrlTemplate, name) {
  const base = uploadUrlTemplate.replace(/\{.*$/, "");
  return `${base}?name=${encodeURIComponent(name)}`;
}

/**
 * @param {{ uploadUrl: string; name: string; bytes: Buffer; contentType: string; token: string }}
 */
async function uploadReleaseAsset({ uploadUrl, name, bytes, contentType, token }) {
  const url = releaseUploadUrl(uploadUrl, name);
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType,
      "X-GitHub-Api-Version": "2022-11-28",
    },
    body: bytes,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Upload ${name} failed (${res.status}): ${text}`);
  }
}

async function main() {
  const refName = process.argv[2] ?? process.env.GITHUB_REF_NAME;
  const manifestsDir = process.argv[3];

  if (!refName || !manifestsDir) {
    fatal(
      "Usage: node scripts/ci/publish-updater-manifest.mjs <tag> <manifests-dir> (example: v0.2.3 updater-manifests)",
    );
  }

  const normalizedRefName = refName.startsWith("refs/tags/")
    ? refName.slice("refs/tags/".length)
    : refName;
  const tag = normalizedRefName;
  const expectedVersion = normalizeVersion(tag);

  const repo = process.env.GITHUB_REPOSITORY;
  if (!repo) fatal("Missing GITHUB_REPOSITORY (expected to run inside GitHub Actions).");

  const token = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!token) fatal("Missing GITHUB_TOKEN / GH_TOKEN (required to update the GitHub Release).");

  const tauriPrivateKey = process.env.TAURI_PRIVATE_KEY ?? "";
  const tauriKeyPassword = process.env.TAURI_KEY_PASSWORD ?? "";

  const jsonFiles = findJsonFiles(manifestsDir);
  if (jsonFiles.length === 0) {
    fatal(`No manifest JSON files found under: ${manifestsDir}`);
  }

  /** @type {Record<string, { url: string; signature: string }>} */
  const mergedPlatforms = {};

  for (const file of jsonFiles) {
    /** @type {any} */
    let manifest;
    try {
      manifest = JSON.parse(fs.readFileSync(file, "utf8"));
    } catch (err) {
      throw new Error(`Failed to parse ${file}: ${err instanceof Error ? err.message : String(err)}`);
    }

    const v = typeof manifest?.version === "string" ? manifest.version : "";
    if (!v) throw new Error(`Manifest ${file} missing "version"`);
    if (normalizeVersion(v) !== expectedVersion) {
      throw new Error(
        `Manifest ${file} version mismatch: expected ${expectedVersion}, got ${JSON.stringify(v)}`,
      );
    }

    const platforms = manifest?.platforms;
    if (!platforms || typeof platforms !== "object" || Array.isArray(platforms)) {
      throw new Error(`Manifest ${file} missing "platforms" object`);
    }

    for (const [target, entry] of Object.entries(platforms)) {
      const url = typeof entry?.url === "string" ? entry.url.trim() : "";
      const signature = typeof entry?.signature === "string" ? entry.signature.trim() : "";
      if (!url) throw new Error(`Manifest ${file} missing ${target}.url`);
      if (!signature) throw new Error(`Manifest ${file} missing ${target}.signature`);

      const existing = mergedPlatforms[target];
      if (existing) {
        if (existing.url !== url || existing.signature !== signature) {
          throw new Error(
            `Conflicting platform entry for ${target} across manifests.\n` +
              `- existing: ${JSON.stringify(existing)}\n` +
              `- new:      ${JSON.stringify({ url, signature })}\n` +
              `Source file: ${file}`,
          );
        }
        continue;
      }

      mergedPlatforms[target] = { url, signature };
    }
  }

  // Deterministic output: sort platform keys so the merged `latest.json` is stable even when
  // the input manifests contain multiple targets or are discovered in different orders.
  const sortedPlatforms = Object.fromEntries(
    Object.keys(mergedPlatforms)
      .sort((a, b) => a.localeCompare(b))
      .map((key) => [key, mergedPlatforms[key]]),
  );

  const combined = {
    version: expectedVersion,
    notes: `Automated build for ${tag}.`,
    pub_date: new Date().toISOString(),
    platforms: sortedPlatforms,
  };

  const latestJsonText = `${JSON.stringify(combined, null, 2)}\n`;
  const latestJsonBytes = Buffer.from(latestJsonText, "utf8");

  const privateKey = loadEd25519PrivateKey(tauriPrivateKey, tauriKeyPassword);
  const signatureBytes = crypto.sign(null, latestJsonBytes, privateKey);
  if (signatureBytes.length !== 64) {
    throw new Error(`Unexpected Ed25519 signature length: ${signatureBytes.length} bytes (expected 64).`);
  }
  const latestSigText = `${signatureBytes.toString("base64")}\n`;
  const latestSigBytes = Buffer.from(latestSigText, "utf8");

  fs.writeFileSync("latest.json", latestJsonText);
  fs.writeFileSync("latest.json.sig", latestSigText);

  const release = await fetchRelease(repo, tag, token);
  const releaseId = /** @type {number} */ (release?.id);
  if (!releaseId) throw new Error("Release payload missing id.");

  const uploadUrl = /** @type {string} */ (release?.upload_url);
  if (!uploadUrl) throw new Error("Release payload missing upload_url.");

  const assets = await fetchAllReleaseAssets(repo, releaseId, token);
  const assetsByName = new Map(
    assets
      .filter((a) => a && typeof a.name === "string" && typeof a.id === "number")
      .map((a) => /** @type {[string, any]} */ ([a.name, a])),
  );

  for (const name of ["latest.json", "latest.json.sig"]) {
    const existing = assetsByName.get(name);
    if (existing) {
      console.log(`Deleting existing release asset: ${name} (id=${existing.id})`);
      await deleteReleaseAsset(repo, existing.id, token);
    }
  }

  console.log(`Uploading combined updater manifest assets to release ${tag}...`);
  await uploadReleaseAsset({
    uploadUrl,
    name: "latest.json",
    bytes: latestJsonBytes,
    contentType: "application/json",
    token,
  });
  await uploadReleaseAsset({
    uploadUrl,
    name: "latest.json.sig",
    bytes: latestSigBytes,
    contentType: "text/plain",
    token,
  });

  console.log(`publish-updater-manifest: uploaded latest.json + latest.json.sig for ${tag}`);
}

main().catch((err) => {
  fatal(err instanceof Error ? err.stack ?? err.message : String(err));
});
