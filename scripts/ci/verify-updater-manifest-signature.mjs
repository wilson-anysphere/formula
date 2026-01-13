#!/usr/bin/env node
/**
 * Verify that `latest.json.sig` matches `latest.json` under the updater public key embedded in
 * apps/desktop/src-tauri/tauri.conf.json (plugins.updater.pubkey).
 *
 * Tauri's updater uses Ed25519 signatures. The updater public key in `tauri.conf.json` is a base64
 * string produced by `tauri signer generate`. Depending on the Tauri/minisign version, the base64
 * may represent:
 *   - a minisign public key file (base64 of a text file containing a base64 payload line), or
 *   - the raw minisign payload ("Ed" + key id + raw Ed25519 public key), or
 *   - the raw 32-byte Ed25519 public key (less common).
 *
 * Usage:
 *   node scripts/ci/verify-updater-manifest-signature.mjs path/to/latest.json path/to/latest.json.sig
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createPublicKey, verify } from "node:crypto";
import process from "node:process";

const repoRoot = path.resolve(fileURLToPath(new URL("../..", import.meta.url)));
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
// Test/debug hook: allow overriding the Tauri config path so unit tests can operate on a temp copy
// instead of mutating the real repo config.
const configPath = process.env.FORMULA_TAURI_CONF_PATH
  ? path.resolve(repoRoot, process.env.FORMULA_TAURI_CONF_PATH)
  : defaultConfigPath;
const relativeConfigPath = path.relative(repoRoot, configPath);

const PLACEHOLDER_PUBKEY = "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY";

/**
 * @param {string} message
 */
function fail(message) {
  process.exitCode = 1;
  console.error(message);
}

/**
 * @param {string} heading
 * @param {string[]} details
 */
function failBlock(heading, details) {
  fail(`\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}\n`);
}

/**
 * Strict base64 decoder with support for unpadded base64 and base64url.
 *
 * Node's `Buffer.from(str, "base64")` is permissive and will silently ignore invalid characters.
 * For CI/security checks, we'd rather fail fast with a clear error.
 *
 * @param {string} label
 * @param {string} value
 */
function decodeBase64Strict(label, value) {
  const normalized = value.trim().replace(/\s+/g, "");
  if (normalized.length === 0) {
    throw new Error(`${label} is empty.`);
  }

  // Support base64url by normalizing to standard base64.
  let base64 = normalized.replace(/-/g, "+").replace(/_/g, "/");

  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) {
    throw new Error(`${label} is not valid base64.`);
  }

  const mod = base64.length % 4;
  if (mod === 1) {
    throw new Error(`${label} has invalid base64 padding.`);
  }
  if (mod !== 0) {
    base64 += "=".repeat(4 - mod);
  }

  return Buffer.from(base64, "base64");
}

/**
 * @param {Uint8Array} rawKey32
 */
function ed25519PublicKeyFromRaw(rawKey32) {
  if (rawKey32.length !== 32) {
    throw new Error(`Expected 32-byte Ed25519 public key, got ${rawKey32.length} bytes.`);
  }
  const ed25519SpkiPrefix = Buffer.from("302a300506032b6570032100", "hex");
  const spkiDer = Buffer.concat([ed25519SpkiPrefix, Buffer.from(rawKey32)]);
  return createPublicKey({ key: spkiDer, format: "der", type: "spki" });
}

/**
 * Parse the updater public key stored in `tauri.conf.json` (`plugins.updater.pubkey`).
 *
 * @param {string} pubkeyBase64
 */
function parseTauriUpdaterPubkey(pubkeyBase64) {
  /** @type {Buffer} */
  let decoded;
  try {
    decoded = decodeBase64Strict("Updater pubkey", pubkeyBase64);
  } catch (err) {
    throw new Error(
      `Updater pubkey is not valid base64 (${err instanceof Error ? err.message : String(err)}).`,
    );
  }

  // Some setups store the raw Ed25519 public key bytes directly.
  if (decoded.length === 32) {
    return { keyIdHex: null, publicKey: decoded };
  }

  // Some setups store the minisign payload directly: "Ed" + key id (8) + pubkey (32) = 42 bytes.
  if (decoded.length === 42 && decoded[0] === 0x45 && decoded[1] === 0x64) {
    const keyIdBytes = decoded.subarray(2, 10);
    const keyIdHex = Buffer.from(keyIdBytes).reverse().toString("hex").toUpperCase();
    const publicKey = decoded.subarray(10);
    return { keyIdHex, publicKey };
  }

  // Most commonly (tauri signer generate): base64 encodes a minisign public key file, which
  // contains a base64 payload line that decodes to 42 bytes as above.
  const text = decoded.toString("utf8");
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const base64Line = [...lines].reverse().find((l) => /^[A-Za-z0-9+/]+={0,2}$/.test(l));
  if (!base64Line) {
    throw new Error(
      `Updater pubkey did not contain a minisign base64 payload line. Decoded text:\n${text}`,
    );
  }

  const payload = Buffer.from(base64Line, "base64");
  if (payload.length !== 42) {
    throw new Error(`Updater pubkey minisign payload decoded to ${payload.length} bytes (expected 42).`);
  }
  if (payload[0] !== 0x45 || payload[1] !== 0x64) {
    throw new Error(`Updater pubkey minisign payload does not start with \"Ed\" (expected 0x45 0x64).`);
  }

  const keyIdBytes = payload.subarray(2, 10);
  const keyIdHex = Buffer.from(keyIdBytes).reverse().toString("hex").toUpperCase();
  const publicKey = payload.subarray(10);
  if (publicKey.length !== 32) {
    throw new Error(
      `Updater pubkey minisign payload contained ${publicKey.length} public key bytes (expected 32).`,
    );
  }

  return { keyIdHex, publicKey };
}

/**
 * Parse a Tauri updater signature string.
 *
 * `.sig` files and `latest.json.sig` are typically text files containing base64. Depending on the
 * Tauri/minisign version, the base64 may decode to either:
 *  - 64 bytes: raw Ed25519 signature bytes
 *  - 74 bytes: minisign signature structure: "Ed" + key id (8) + signature (64)
 *
 * @param {string} signatureText
 */
function parseTauriSignature(signatureText) {
  const trimmed = signatureText.trim();
  if (!trimmed) {
    throw new Error("Signature file is empty.");
  }

  // Some `.sig` files contain multiple lines; pick base64-ish candidates.
  const candidates = trimmed.includes("\n")
    ? trimmed
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter((l) => /^[A-Za-z0-9+/]+={0,2}$/.test(l))
    : [trimmed];

  for (const candidate of candidates) {
    let bytes;
    try {
      bytes = Buffer.from(candidate, "base64");
    } catch {
      continue;
    }

    if (bytes.length === 64) {
      return { signature: bytes, keyIdHex: null };
    }

    if (bytes.length === 74 && bytes[0] === 0x45 && bytes[1] === 0x64) {
      const keyIdBytes = bytes.subarray(2, 10);
      const keyIdHex = Buffer.from(keyIdBytes).reverse().toString("hex").toUpperCase();
      const signature = bytes.subarray(10);
      if (signature.length === 64) {
        return { signature, keyIdHex };
      }
    }
  }

  throw new Error(
    `Signature did not decode to 64 raw bytes or a 74-byte minisign structure. First 200 chars:\n${trimmed.slice(
      0,
      200,
    )}`,
  );
}

function main() {
  const [latestJsonPath, latestSigPath] = process.argv.slice(2);
  if (!latestJsonPath || !latestSigPath) {
    failBlock(`Usage`, [
      `node scripts/ci/verify-updater-manifest-signature.mjs <path-to-latest.json> <path-to-latest.json.sig>`,
    ]);
    return;
  }

  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Updater manifest signature verification failed`, [
      `Failed to read/parse ${relativeConfigPath}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  const pubkeyValue = config?.plugins?.updater?.pubkey;
  if (typeof pubkeyValue !== "string" || pubkeyValue.trim().length === 0) {
    failBlock(`Missing updater public key`, [
      `Expected ${relativeConfigPath} → plugins.updater.pubkey to be a non-empty string.`,
    ]);
    return;
  }

  if (pubkeyValue.trim() === PLACEHOLDER_PUBKEY || pubkeyValue.trim().includes("REPLACE_WITH")) {
    failBlock(`Invalid updater public key`, [
      `plugins.updater.pubkey is still set to the placeholder value "${PLACEHOLDER_PUBKEY}".`,
      `Replace it with the real updater public key (base64) before tagging a release.`,
    ]);
    return;
  }

  /** @type {{ keyIdHex: string | null; publicKey: Uint8Array }} */
  let parsedPubkey;
  try {
    parsedPubkey = parseTauriUpdaterPubkey(pubkeyValue);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Invalid updater public key`, [
      `Failed to parse ${relativeConfigPath} → plugins.updater.pubkey.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  let latestJsonBytes;
  try {
    latestJsonBytes = readFileSync(latestJsonPath);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Updater manifest signature verification failed`, [
      `Failed to read latest.json from ${JSON.stringify(latestJsonPath)}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  let sigRaw;
  try {
    sigRaw = readFileSync(latestSigPath, "utf8");
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Updater manifest signature verification failed`, [
      `Failed to read latest.json.sig from ${JSON.stringify(latestSigPath)}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  /** @type {{ signature: Uint8Array; keyIdHex: string | null }} */
  let parsedSignature;
  try {
    parsedSignature = parseTauriSignature(sigRaw);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Invalid signature file`, [
      `Failed to parse ${JSON.stringify(latestSigPath)} as a Tauri updater signature.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  let publicKey;
  try {
    if (parsedSignature.keyIdHex && parsedPubkey.keyIdHex && parsedSignature.keyIdHex !== parsedPubkey.keyIdHex) {
      failBlock(`Updater manifest signature mismatch`, [
        `Signature key id mismatch: latest.json.sig uses ${parsedSignature.keyIdHex} but plugins.updater.pubkey is ${parsedPubkey.keyIdHex}.`,
      ]);
      return;
    }

    publicKey = ed25519PublicKeyFromRaw(parsedPubkey.publicKey);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Invalid updater public key`, [
      `Failed to construct an Ed25519 public key from plugins.updater.pubkey.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  let ok = false;
  try {
    // For Ed25519, the digest algorithm is ignored and must be null.
    ok = verify(null, latestJsonBytes, publicKey, parsedSignature.signature);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    failBlock(`Updater manifest signature verification failed`, [
      `crypto.verify() threw an error.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  if (!ok) {
    failBlock(`Updater manifest signature mismatch`, [
      `The signature in ${JSON.stringify(latestSigPath)} does not match ${JSON.stringify(latestJsonPath)}.`,
      `This usually means TAURI_PRIVATE_KEY does not correspond to the committed plugins.updater.pubkey.`,
    ]);
    return;
  }

  console.log("signature OK");
}

main();
