#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseTauriUpdaterPubkey } from "./tauri-minisign.mjs";

const repoRoot = path.resolve(fileURLToPath(new URL("../..", import.meta.url)));
const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const relativeConfigPath = path.relative(repoRoot, configPath);

/**
 * @param {string} message
 */
function err(message) {
  process.exitCode = 1;
  console.error(message);
}

/**
 * @param {string} heading
 * @param {string[]} details
 */
function errBlock(heading, details) {
  err(`\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}`);
}

/**
 * Returns the decoded bytes if the input looks like a base64 string, otherwise `null`.
 * @param {string} value
 */
function decodeBase64(value) {
  const normalized = value.trim().replace(/\s+/g, "");
  if (normalized.length === 0) return null;

  // Support both standard base64 and base64url (GitHub secrets may contain either).
  let base64 = normalized.replace(/-/g, "+").replace(/_/g, "/");

  // Reject anything that isn't plausibly base64. (Node's base64 decoder is permissive, so validate
  // the alphabet and padding ourselves.)
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) return null;

  // Allow unpadded base64 by adding the required '=' chars.
  const mod = base64.length % 4;
  if (mod === 1) return null;
  if (mod !== 0) base64 += "=".repeat(4 - mod);

  return Buffer.from(base64, "base64");
}

const MINISIGN_PUBLIC_KEY_BYTES = 42; // "Ed" + keyId(8) + pubkey(32)
const MINISIGN_SECRET_KEY_MIN_BYTES = 74; // "Ed" + keyId(8) + secret key(64)

/**
 * @param {string} text
 */
function parseMinisignKeyFileBody(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  const payloadLine = lines.find((line) => !line.toLowerCase().startsWith("untrusted comment:"));
  if (!payloadLine) return null;
  const binary = decodeBase64(payloadLine);
  if (!binary) return null;
  if (binary.length < 2 || binary[0] !== 0x45 || binary[1] !== 0x64) return null; // "Ed"
  if (binary.length < 10) return null;
  // minisign prints key IDs in the "untrusted comment" header as a big-endian hex string. The key
  // ID bytes in the binary payload are little-endian, so reverse them for display to match what
  // contributors see in `tauri.conf.json` and `cargo tauri signer generate` output.
  const keyIdHex = Buffer.from(binary.subarray(2, 10))
    .reverse()
    .toString("hex")
    .toUpperCase();
  return { payloadLine, binary, keyIdHex };
}

/**
 * Try to interpret TAURI_PRIVATE_KEY as a minisign key.
 *
 * Tauri's updater uses minisign signatures. `cargo tauri signer generate` outputs base64 strings
 * that decode to a minisign key file (ASCII) whose payload line base64-decodes to a binary key
 * starting with the magic bytes "Ed".
 *
 * @param {string} value
 * @returns {{ kind: 'secret' | 'public', keyIdHex: string, encrypted?: boolean } | null}
 */
function analyzeMinisignKey(value) {
  const trimmed = value.trim();

  const tryParseKeyFile = (content) => {
    const lowered = content.toLowerCase();
    const hasSecretHeader = lowered.includes("minisign secret key");
    const hasPublicHeader = lowered.includes("minisign public key");
    if (!hasSecretHeader && !hasPublicHeader) return null;

    const parsed = parseMinisignKeyFileBody(content);
    if (!parsed) return null;

    if (parsed.binary.length === MINISIGN_PUBLIC_KEY_BYTES) {
      return { kind: "public", keyIdHex: parsed.keyIdHex };
    }

    if (parsed.binary.length < MINISIGN_SECRET_KEY_MIN_BYTES) return null;

    const encrypted = parsed.binary.length !== MINISIGN_SECRET_KEY_MIN_BYTES;
    return { kind: "secret", encrypted, keyIdHex: parsed.keyIdHex };
  };

  // Raw minisign key file (multiline secret).
  {
    const parsed = tryParseKeyFile(trimmed);
    if (parsed) return parsed;
  }

  const decoded = decodeBase64(trimmed);
  if (!decoded) return null;

  // base64-encoded minisign key file
  {
    const decodedText = decoded.toString("utf8");
    const parsed = tryParseKeyFile(decodedText);
    if (parsed) return parsed;
  }

  // base64-encoded minisign binary key (payload line itself).
  if (decoded.length >= 10 && decoded[0] === 0x45 && decoded[1] === 0x64) {
    const keyIdHex = Buffer.from(decoded.subarray(2, 10))
      .reverse()
      .toString("hex")
      .toUpperCase();
    if (decoded.length === MINISIGN_PUBLIC_KEY_BYTES) return { kind: "public", keyIdHex };
    if (decoded.length < MINISIGN_SECRET_KEY_MIN_BYTES) return null;
    const encrypted = decoded.length !== MINISIGN_SECRET_KEY_MIN_BYTES;
    return { kind: "secret", encrypted, keyIdHex };
  }

  return null;
}

/**
 * TAURI_KEY_PASSWORD is only required when the private key is encrypted.
 *
 * We can't perfectly validate every possible Tauri key encoding without invoking the Tauri CLI,
 * but we can reliably detect the common cases:
 *
 * - Raw Ed25519 private keys (32/64 bytes, base64-encoded) are unencrypted.
 * - PKCS#8 DER (base64-encoded) indicates encryption based on the top-level ASN.1 structure.
 * - PEM keys can be detected via their BEGIN header.
 *
 * @param {string} privateKey
 */
function isEncryptedPrivateKey(privateKey) {
  const trimmed = privateKey.trim();

  const minisign = analyzeMinisignKey(trimmed);
  if (minisign?.kind === "secret") return minisign.encrypted === true;

  // PEM (multiline) keys: the header tells us whether it's encrypted.
  if (trimmed.includes("-----BEGIN ENCRYPTED PRIVATE KEY-----")) return true;
  if (trimmed.includes("-----BEGIN PRIVATE KEY-----")) return false;

  const decoded = decodeBase64(trimmed);
  if (!decoded) return undefined;

  // Some setups store the PEM contents base64-encoded in GitHub Secrets to avoid multiline
  // formatting issues. Detect that by scanning the decoded bytes for PEM headers.
  if (decoded.includes(Buffer.from("-----BEGIN ENCRYPTED PRIVATE KEY-----"))) return true;
  if (decoded.includes(Buffer.from("-----BEGIN PRIVATE KEY-----"))) return false;

  // Raw Ed25519 secret key (32 bytes) or secret+public (64 bytes).
  if (decoded.length === 32 || decoded.length === 64) return false;

  // Detect PKCS#8 encryption based on the outer ASN.1 sequence:
  // - PrivateKeyInfo ::= SEQUENCE { INTEGER version, ... } (unencrypted)
  // - EncryptedPrivateKeyInfo ::= SEQUENCE { SEQUENCE algId, OCTET STRING data } (encrypted)
  try {
    /**
     * @param {Buffer} buf
     * @param {number} offset
     */
    function readLength(buf, offset) {
      if (offset >= buf.length) throw new Error("DER length out of bounds");
      const first = buf[offset];
      if ((first & 0x80) === 0) return { length: first, bytes: 1 };
      const count = first & 0x7f;
      if (count === 0 || count > 6) throw new Error("Unsupported DER length encoding");
      if (offset + 1 + count > buf.length) throw new Error("DER length out of bounds");
      let length = 0;
      for (let i = 0; i < count; i++) length = (length << 8) | buf[offset + 1 + i];
      return { length, bytes: 1 + count };
    }

    /**
     * @param {Buffer} buf
     * @param {number} offset
     */
    function readTlv(buf, offset) {
      if (offset >= buf.length) throw new Error("DER tag out of bounds");
      const tag = buf[offset];
      const { length, bytes } = readLength(buf, offset + 1);
      const valueStart = offset + 1 + bytes;
      const valueEnd = valueStart + length;
      if (valueEnd > buf.length) throw new Error("DER value out of bounds");
      return { tag, valueStart, valueEnd, next: valueEnd };
    }

    const outer = readTlv(decoded, 0);
    if (outer.tag !== 0x30) return undefined;

    const first = readTlv(decoded, outer.valueStart);
    if (first.valueEnd > outer.valueEnd) return undefined;

    const second = readTlv(decoded, first.next);
    if (second.valueEnd > outer.valueEnd) return undefined;

    if (first.tag === 0x02) return false; // INTEGER version => unencrypted PKCS#8.
    if (first.tag === 0x30 && second.tag === 0x04) return true; // algId + OCTET STRING => encrypted.
  } catch {
    // Fall through to unknown.
  }

  return undefined;
}

/**
 * Basic format validation for the updater private key.
 *
 * The tauri-action signer supports minisign keys produced by `cargo tauri signer generate`
 * (the common case for Tauri updater signing), and may also support other encodings depending on
 * the Tauri toolchain:
 * - minisign secret key (base64 string that decodes to a minisign key file)
 * - PEM (`-----BEGIN (ENCRYPTED )?PRIVATE KEY-----`)
 * - base64-encoded raw Ed25519 keys (32/64 bytes)
 * - base64-encoded PKCS#8 DER (ASN.1 SEQUENCE, starts with 0x30)
 * - base64-encoded PEM blobs (decoded bytes contain the PEM header)
 *
 * @param {string} privateKey
 */
function validatePrivateKeyFormat(privateKey) {
  const trimmed = privateKey.trim();
  if (trimmed.length === 0) return;

  const minisign = analyzeMinisignKey(trimmed);
  if (minisign) {
    if (minisign.kind === "public") {
      errBlock(`Invalid TAURI_PRIVATE_KEY`, [
        `TAURI_PRIVATE_KEY looks like a minisign *public* key, not a private key.`,
        `Did you accidentally paste apps/desktop/src-tauri/tauri.conf.json → plugins.updater.pubkey ?`,
        `Set TAURI_PRIVATE_KEY to the private key printed by:`,
        `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
        `Then store it as the GitHub Actions secret TAURI_PRIVATE_KEY.`,
      ]);
      err(`\nUpdater signing secrets preflight failed. Fix TAURI_PRIVATE_KEY before tagging a release.\n`);
      throw new Error("invalid TAURI_PRIVATE_KEY");
    }
    return;
  }

  if (
    trimmed.includes("-----BEGIN PRIVATE KEY-----") ||
    trimmed.includes("-----BEGIN ENCRYPTED PRIVATE KEY-----")
  ) {
    return;
  }

  const decoded = decodeBase64(trimmed);
  if (!decoded) {
    errBlock(`Invalid TAURI_PRIVATE_KEY`, [
      `TAURI_PRIVATE_KEY is set but does not look like a valid Tauri updater private key.`,
      `Expected the value printed by \`cargo tauri signer generate\` (minisign secret key; base64 string).`,
      `Make sure you paste the private key printed by:`,
      `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
      `Then store it as the GitHub Actions secret TAURI_PRIVATE_KEY (Settings → Secrets and variables → Actions).`,
    ]);
    err(`\nUpdater signing secrets preflight failed. Fix TAURI_PRIVATE_KEY before tagging a release.\n`);
    throw new Error("invalid TAURI_PRIVATE_KEY");
  }

  // base64-encoded PEM
  if (decoded.includes(Buffer.from("-----BEGIN PRIVATE KEY-----"))) return;
  if (decoded.includes(Buffer.from("-----BEGIN ENCRYPTED PRIVATE KEY-----"))) return;

  // raw key material
  if (decoded.length === 32 || decoded.length === 64) return;

  // minisign binary key (Ed25519 + key id prefix)
  if (decoded.length >= 2 && decoded[0] === 0x45 && decoded[1] === 0x64) {
    if (decoded.length === MINISIGN_PUBLIC_KEY_BYTES) {
      errBlock(`Invalid TAURI_PRIVATE_KEY`, [
        `TAURI_PRIVATE_KEY looks like a minisign *public* key (42 bytes), not a private key.`,
        `Set TAURI_PRIVATE_KEY to the private key printed by:`,
        `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
      ]);
      err(`\nUpdater signing secrets preflight failed. Fix TAURI_PRIVATE_KEY before tagging a release.\n`);
      throw new Error("invalid TAURI_PRIVATE_KEY");
    }
    return;
  }

  // PKCS#8 DER should begin with an ASN.1 SEQUENCE (0x30).
  if (decoded.length > 0 && decoded[0] === 0x30) return;

  errBlock(`Invalid TAURI_PRIVATE_KEY`, [
    `TAURI_PRIVATE_KEY is set but does not look like a supported Tauri signing key format.`,
    `Expected a minisign secret key (from \`cargo tauri signer generate\`), PEM key, base64-encoded raw Ed25519 key, or base64-encoded PKCS#8 DER.`,
    `Regenerate keys with:`,
    `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
    `Then update the GitHub Actions secret TAURI_PRIVATE_KEY.`,
  ]);
  err(`\nUpdater signing secrets preflight failed. Fix TAURI_PRIVATE_KEY before tagging a release.\n`);
  throw new Error("invalid TAURI_PRIVATE_KEY");
}

function main() {
  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock(`Updater signing secrets preflight failed`, [
      `Failed to read/parse ${relativeConfigPath}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  const updater = config?.plugins?.updater;
  const active = updater?.active === true;
  if (!active) {
    console.log(
      `Updater signing secrets preflight: updater is not active (plugins.updater.active !== true); skipping secret validation.`
    );
    return;
  }

  const privateKey = process.env.TAURI_PRIVATE_KEY ?? "";
  const password = process.env.TAURI_KEY_PASSWORD ?? "";

  const missing = [];
  if (privateKey.trim().length === 0) {
    missing.push("TAURI_PRIVATE_KEY");
  }

  // If the key is present, ensure it at least looks like a valid Tauri signing key
  // so we fail early with a helpful error instead of failing inside tauri-action.
  if (privateKey.trim().length > 0) {
    try {
      validatePrivateKeyFormat(privateKey);
    } catch {
      return;
    }
  }

  // This repo's release workflow re-signs the combined updater manifest (`latest.json`) in Node
  // (see scripts/ci/publish-updater-manifest.mjs). That signing path can load:
  // - unencrypted minisign secret keys (74-byte payload), and
  // - PKCS#8 keys (encrypted or not, via TAURI_KEY_PASSWORD)
  //
  // Minisign secret keys encrypted with a passphrase are not supported there, so fail fast here
  // with an actionable error to avoid wasting build time.
  const minisign = privateKey.trim().length > 0 ? analyzeMinisignKey(privateKey) : null;
  if (minisign?.kind === "secret" && minisign.encrypted === true) {
    errBlock(`Encrypted minisign TAURI_PRIVATE_KEY is not supported`, [
      `TAURI_PRIVATE_KEY looks like an encrypted minisign secret key (generated with a password).`,
      `This repository's release workflow currently requires an unencrypted minisign secret key so it can re-sign the combined updater manifest (latest.json).`,
      `Fix: generate a new keypair WITHOUT a password:`,
      `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
      `  # When prompted for a password, leave it blank.`,
      `Then update BOTH:`,
      `  - apps/desktop/src-tauri/tauri.conf.json → plugins.updater.pubkey (public key; safe to commit)`,
      `  - GitHub Actions secret TAURI_PRIVATE_KEY (private key)`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);
    err(`\nUpdater signing secrets preflight failed. Fix TAURI_PRIVATE_KEY before tagging a release.\n`);
    return;
  }

  const passwordTrimmed = password.trim();
  const passwordMissing = passwordTrimmed.length === 0;

  const encrypted = privateKey.trim().length > 0 ? isEncryptedPrivateKey(privateKey) : undefined;
  if (passwordMissing && encrypted === true) missing.push("TAURI_KEY_PASSWORD");

  // If we can't determine whether the key is encrypted, we err on the side of catching
  // misconfiguration early by requiring a password. This avoids failures later inside
  // `tauri-apps/tauri-action` when an encrypted key is provided without a passphrase.
  if (passwordMissing && encrypted === undefined && privateKey.trim().length > 0) {
    missing.push("TAURI_KEY_PASSWORD");
  }

  if (missing.length > 0) {
    errBlock(`Missing Tauri updater signing secrets`, [
      `This release workflow is building a tagged desktop release and needs to sign updater artifacts.`,
      `Expected secrets: TAURI_PRIVATE_KEY, TAURI_KEY_PASSWORD (required if your private key is encrypted).`,
      `Missing/empty GitHub Actions repository secrets (Settings → Secrets and variables → Actions):`,
      ...missing.map((name) => name),
      `Generate a new keypair (prints public + private key):`,
      `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
      `  # Agents: (cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate)`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);

    err(`\nUpdater signing secrets preflight failed. Fix the missing secrets above before tagging a release.\n`);
    return;
  }

  // When both keys are minisign keys, verify the configured updater public key matches the
  // private key used for signing. A mismatch would produce signatures that the app cannot verify,
  // leading to broken auto-update behavior despite a successful build.
  const pubkeyValue = updater?.pubkey;
  if (typeof pubkeyValue === "string" && pubkeyValue.trim().length > 0) {
    /** @type {{ kind: "public"; keyIdHex: string } | null} */
    let pubkey = null;
    try {
      const parsed = parseTauriUpdaterPubkey(pubkeyValue);
      if (parsed.keyId) {
        pubkey = { kind: "public", keyIdHex: parsed.keyId };
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      errBlock(`Invalid updater public key`, [
        `Failed to parse ${relativeConfigPath} → plugins.updater.pubkey as a Tauri/minisign public key.`,
        `Error: ${msg}`,
      ]);
      err(`\nUpdater signing secrets preflight failed. Fix the updater public key above before tagging a release.\n`);
      return;
    }
    const secret = analyzeMinisignKey(privateKey);
    if (pubkey?.kind === "public" && secret?.kind === "secret") {
      if (pubkey.keyIdHex !== secret.keyIdHex) {
        errBlock(`Tauri updater key mismatch`, [
          `The updater public key embedded in the app does not match the private key used for signing.`,
          `apps/desktop/src-tauri/tauri.conf.json → plugins.updater.pubkey key id: ${pubkey.keyIdHex}`,
          `GitHub Actions secret TAURI_PRIVATE_KEY key id: ${secret.keyIdHex}`,
          `Regenerate a matching keypair with:`,
          `  (cd apps/desktop/src-tauri && cargo tauri signer generate)`,
          `Then update BOTH:`,
          `  - tauri.conf.json plugins.updater.pubkey (public key; safe to commit)`,
          `  - GitHub Actions secret TAURI_PRIVATE_KEY (private key)`,
        ]);
        err(`\nUpdater signing secrets preflight failed. Fix the key mismatch above before tagging a release.\n`);
        return;
      }
    }
  }

  if (passwordMissing) {
    console.log(
      `Updater signing secrets preflight passed (TAURI_KEY_PASSWORD is empty; detected unencrypted TAURI_PRIVATE_KEY).`
    );
    return;
  }

  console.log(`Updater signing secrets preflight passed.`);
}

main();
