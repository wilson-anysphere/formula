#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

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
      `  (cd apps/desktop/src-tauri && tauri signer generate)`,
      `  # Agents: (cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate)`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);

    err(`\nUpdater signing secrets preflight failed. Fix the missing secrets above before tagging a release.\n`);
    return;
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
