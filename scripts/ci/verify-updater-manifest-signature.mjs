#!/usr/bin/env node
/**
 * Verify that `latest.json.sig` matches `latest.json` under the updater public key embedded in
 * apps/desktop/src-tauri/tauri.conf.json (plugins.updater.pubkey).
 *
 * Tauri's updater uses Ed25519 signatures. The updater public key in `tauri.conf.json` is a base64
 * string produced by `cargo tauri signer generate`. Depending on the Tauri/minisign version, the base64
 * may represent:
 *   - a minisign public key file (base64 of a text file containing a base64 payload line), or
 *   - the raw minisign payload ("Ed" + key id + raw Ed25519 public key), or
 *   - the raw 32-byte Ed25519 public key (less common).
 *
 * Usage:
 *   node scripts/ci/verify-updater-manifest-signature.mjs path/to/latest.json path/to/latest.json.sig
 */
import { appendFileSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { verify } from "node:crypto";
import process from "node:process";
import {
  ed25519PublicKeyFromRaw,
  parseTauriUpdaterPubkey,
  parseTauriUpdaterSignature,
} from "./tauri-minisign.mjs";

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

  /** @type {{ keyId: string | null; publicKeyBytes: Uint8Array }} */
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

  /** @type {{ signatureBytes: Uint8Array; keyId: string | null }} */
  let parsedSignature;
  try {
    parsedSignature = parseTauriUpdaterSignature(sigRaw, path.basename(latestSigPath));
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
    if (parsedSignature.keyId && parsedPubkey.keyId && parsedSignature.keyId !== parsedPubkey.keyId) {
      failBlock(`Updater manifest signature mismatch`, [
        `Signature key id mismatch: latest.json.sig uses ${parsedSignature.keyId} but plugins.updater.pubkey is ${parsedPubkey.keyId}.`,
      ]);
      return;
    }

    publicKey = ed25519PublicKeyFromRaw(parsedPubkey.publicKeyBytes);
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
    ok = verify(null, latestJsonBytes, publicKey, parsedSignature.signatureBytes);
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

  // If running in GitHub Actions, append a short note to the step summary so the
  // signature check result is visible alongside the manifest target table written
  // by the upstream validator.
  const stepSummaryPath = process.env.GITHUB_STEP_SUMMARY;
  if (stepSummaryPath) {
    try {
      appendFileSync(stepSummaryPath, `- Manifest signature: OK\n`, "utf8");
    } catch {
      // Non-fatal: the signature verification already passed; don't fail the release
      // workflow just because the step summary could not be updated.
    }
  }
}

main();
