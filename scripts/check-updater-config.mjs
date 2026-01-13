#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const relativeConfigPath = path.relative(repoRoot, configPath);

const PLACEHOLDER_PUBKEY_MARKER = "REPLACE_WITH";
const PLACEHOLDER_ENDPOINTS = new Set([
  // Documented as a placeholder in docs/release.md.
  "https://releases.formula.app/{{target}}/{{current_version}}",
]);

const MINISIGN_PUBLIC_KEY_BYTES = 42; // "Ed" + keyId(8) + pubkey(32)

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

  // Support both standard base64 and base64url.
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
 * Parse a minisign key file (plaintext) and return the decoded binary key payload.
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
  return binary;
}

/**
 * Returns true if the configured `plugins.updater.pubkey` looks like a valid minisign public key.
 * Accepts:
 * - base64-encoded minisign public key file (what `cargo tauri signer generate` prints)
 * - raw minisign public key file contents (rare, but technically representable in JSON)
 * - base64-encoded minisign binary public key (payload line only)
 * @param {string} value
 */
function looksLikeMinisignPublicKey(value) {
  const trimmed = value.trim();
  if (trimmed.length === 0) return false;

  // Raw key file content.
  if (trimmed.toLowerCase().includes("minisign public key")) {
    const binary = parseMinisignKeyFileBody(trimmed);
    return Boolean(binary && binary.length === MINISIGN_PUBLIC_KEY_BYTES);
  }

  const decoded = decodeBase64(trimmed);
  if (!decoded) return false;

  // base64-encoded key file content
  {
    const text = decoded.toString("utf8");
    if (text.toLowerCase().includes("minisign public key")) {
      const binary = parseMinisignKeyFileBody(text);
      if (binary && binary.length === MINISIGN_PUBLIC_KEY_BYTES) return true;
    }
  }

  // base64-encoded binary key (payload line)
  return decoded.length === MINISIGN_PUBLIC_KEY_BYTES && decoded[0] === 0x45 && decoded[1] === 0x64;
}

function main() {
  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock(`Updater config preflight failed`, [
      `Failed to read/parse ${relativeConfigPath}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  const updater = config?.plugins?.updater;
  const active = updater?.active === true;

  if (!active) {
    console.log(
      `Updater config preflight: updater is not active (plugins.updater.active !== true); skipping validation.`
    );
    return;
  }

  const pubkey = updater?.pubkey;
  if (typeof pubkey !== "string" || pubkey.trim().length === 0) {
    errBlock(`Invalid updater config: plugins.updater.pubkey`, [
      `Expected a non-empty string because plugins.updater.active=true.`,
      `Set ${relativeConfigPath} → plugins.updater.pubkey to the public key printed by:`,
      `  cd apps/desktop/src-tauri && cargo tauri signer generate`,
      `  # Agents: cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);
  } else if (pubkey.includes(PLACEHOLDER_PUBKEY_MARKER)) {
    errBlock(`Invalid updater config: plugins.updater.pubkey`, [
      `Looks like a placeholder value (contains "${PLACEHOLDER_PUBKEY_MARKER}").`,
      `Replace it with the real updater public key (safe to commit).`,
      `The matching private key must be present in GitHub Actions as the TAURI_PRIVATE_KEY secret.`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);
  } else if (!looksLikeMinisignPublicKey(pubkey)) {
    errBlock(`Invalid updater config: plugins.updater.pubkey`, [
      `plugins.updater.pubkey is set but does not look like a minisign public key.`,
      `Expected the value printed by \`cargo tauri signer generate\` (base64-encoded minisign key file).`,
      `Update ${relativeConfigPath} → plugins.updater.pubkey with a real updater public key before tagging a release.`,
      `See docs/release.md ("Tauri updater keys").`,
    ]);
  }

  const endpoints = updater?.endpoints;
  if (!Array.isArray(endpoints) || endpoints.length === 0) {
    errBlock(`Invalid updater config: plugins.updater.endpoints`, [
      `Expected a non-empty array because plugins.updater.active=true.`,
      `Set ${relativeConfigPath} → plugins.updater.endpoints to one or more update JSON URLs.`,
      `Example: ["https://updates.example.com/{{target}}/{{current_version}}"]`,
      `See docs/release.md ("Hosting updater endpoints").`,
    ]);
    return;
  }

  const invalidEndpoints = endpoints
    .map((value, i) => ({ value, i }))
    .filter(({ value }) => typeof value !== "string" || value.trim().length === 0);
  if (invalidEndpoints.length > 0) {
    errBlock(`Invalid updater config: plugins.updater.endpoints`, [
      `All endpoints must be non-empty strings.`,
      ...invalidEndpoints.map(
        ({ i, value }) =>
          `endpoints[${i}] is ${typeof value === "string" ? JSON.stringify(value) : String(value)}`
      ),
    ]);
  }

  const placeholderEndpoints = endpoints
    .map((value, i) => ({ value, i }))
    .filter(({ value }) => typeof value === "string")
    .filter(({ value }) => {
      const trimmed = value.trim();
      return (
        PLACEHOLDER_ENDPOINTS.has(trimmed) ||
        trimmed.includes("REPLACE_WITH") ||
        trimmed.includes("example.com") ||
        trimmed.includes("localhost")
      );
    });
  if (placeholderEndpoints.length > 0) {
    errBlock(`Invalid updater config: plugins.updater.endpoints`, [
      `One or more endpoints look like placeholder values.`,
      ...placeholderEndpoints.map(
        ({ i, value }) => `endpoints[${i}] = ${JSON.stringify(value.trim())}`
      ),
      `Replace them with your real update JSON URL(s) before tagging a release.`,
      `See docs/release.md ("Hosting updater endpoints").`,
    ]);
  }

  if (process.exitCode) {
    err(`\nUpdater config preflight failed. Fix the errors above before tagging a release.\n`);
    return;
  }

  console.log(`Updater config preflight passed (${relativeConfigPath}).`);
}

main();
