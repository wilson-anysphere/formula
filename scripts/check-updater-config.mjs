#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const repoRoot = path.resolve(path.dirname(SCRIPT_PATH), "..");

const DEFAULT_TAURI_CONFIG_RELATIVE_PATH = "apps/desktop/src-tauri/tauri.conf.json";
const DEFAULT_TAURI_CONFIG_PATH = path.join(repoRoot, DEFAULT_TAURI_CONFIG_RELATIVE_PATH);

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

/**
 * @typedef {{ heading: string, details: string[] }} ErrorBlock
 */

/**
 * Validate a parsed `tauri.conf.json` updater config.
 *
 * Exported so node:test suites can validate the script logic without spawning the CLI.
 *
 * @param {any} config
 * @param {{ configRelativePath?: string }} [options]
 * @returns {{ ok: boolean, skipped: boolean, errorBlocks: ErrorBlock[] }}
 */
export function checkUpdaterConfig(config, options = {}) {
  const configRelativePath = options.configRelativePath ?? DEFAULT_TAURI_CONFIG_RELATIVE_PATH;
  /** @type {ErrorBlock[]} */
  const errorBlocks = [];

  const updater = config?.plugins?.updater;
  const active = updater?.active === true;

  if (!active) {
    return { ok: true, skipped: true, errorBlocks };
  }

  const pubkey = updater?.pubkey;
  if (typeof pubkey !== "string" || pubkey.trim().length === 0) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.pubkey",
      details: [
        "Expected a non-empty string because plugins.updater.active=true.",
        `Set ${configRelativePath} → plugins.updater.pubkey to the public key printed by:`,
        "  cd apps/desktop/src-tauri && cargo tauri signer generate",
        "  # Agents: cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate",
        'See docs/release.md ("Tauri updater keys").',
      ],
    });
  } else if (pubkey.includes(PLACEHOLDER_PUBKEY_MARKER)) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.pubkey",
      details: [
        `Looks like a placeholder value (contains "${PLACEHOLDER_PUBKEY_MARKER}").`,
        "Replace it with the real updater public key (safe to commit).",
        "The matching private key must be present in GitHub Actions as the TAURI_PRIVATE_KEY secret.",
        'See docs/release.md ("Tauri updater keys").',
      ],
    });
  } else if (!looksLikeMinisignPublicKey(pubkey)) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.pubkey",
      details: [
        "plugins.updater.pubkey is set but does not look like a minisign public key.",
        "Expected the value printed by `cargo tauri signer generate` (base64-encoded minisign key file).",
        `Update ${configRelativePath} → plugins.updater.pubkey with a real updater public key before tagging a release.`,
        'See docs/release.md ("Tauri updater keys").',
      ],
    });
  }

  const endpoints = updater?.endpoints;
  if (!Array.isArray(endpoints) || endpoints.length === 0) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.endpoints",
      details: [
        "Expected a non-empty array because plugins.updater.active=true.",
        `Set ${configRelativePath} → plugins.updater.endpoints to one or more update JSON URLs.`,
        `Example: ["https://updates.example.com/{{target}}/{{current_version}}"]`,
        'See docs/release.md ("Hosting updater endpoints").',
      ],
    });
    return { ok: false, skipped: false, errorBlocks };
  }

  const invalidEndpoints = endpoints
    .map((value, i) => ({ value, i }))
    .filter(({ value }) => typeof value !== "string" || value.trim().length === 0);
  if (invalidEndpoints.length > 0) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.endpoints",
      details: [
        "All endpoints must be non-empty strings.",
        ...invalidEndpoints.map(
          ({ i, value }) =>
            `endpoints[${i}] is ${typeof value === "string" ? JSON.stringify(value) : String(value)}`,
        ),
      ],
    });
  }

  const placeholderEndpoints = endpoints
    .map((value, i) => ({ value, i }))
    .filter(({ value }) => typeof value === "string")
    .filter(({ value }) => {
      const trimmed = value.trim();
      return (
        PLACEHOLDER_ENDPOINTS.has(trimmed) ||
        trimmed.includes(PLACEHOLDER_PUBKEY_MARKER) ||
        trimmed.includes("example.com") ||
        trimmed.includes("localhost")
      );
    });
  if (placeholderEndpoints.length > 0) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.endpoints",
      details: [
        "One or more endpoints look like placeholder values.",
        ...placeholderEndpoints.map(({ i, value }) => `endpoints[${i}] = ${JSON.stringify(value.trim())}`),
        "Replace them with your real update JSON URL(s) before tagging a release.",
        'See docs/release.md ("Hosting updater endpoints").',
      ],
    });
  }

  /** @type {Array<{ i: number, value: string, reason: string }>} */
  const endpointUrlErrors = [];
  for (let i = 0; i < endpoints.length; i += 1) {
    const raw = endpoints[i];
    if (typeof raw !== "string") continue;
    const value = raw.trim();
    if (value.length === 0) continue;

    let url;
    try {
      url = new URL(value);
    } catch {
      endpointUrlErrors.push({
        i,
        value,
        reason:
          "must be a valid absolute URL starting with https:// (for example: https://github.com/<org>/<repo>/releases/latest/download/latest.json)",
      });
      continue;
    }

    if (url.protocol !== "https:") {
      endpointUrlErrors.push({
        i,
        value,
        reason: `must use https:// (plaintext ${url.protocol}// updater endpoints are not allowed)`,
      });
      continue;
    }

    if (!url.hostname) {
      endpointUrlErrors.push({
        i,
        value,
        reason: "must include a hostname (absolute https:// URL)",
      });
    }
  }

  if (endpointUrlErrors.length > 0) {
    errorBlocks.push({
      heading: "Invalid updater config: plugins.updater.endpoints",
      details: [
        "Each endpoint must be a valid absolute URL and must use https:// (plaintext http:// is not allowed).",
        ...endpointUrlErrors.map(
          ({ i, value, reason }) => `endpoints[${i}] = ${JSON.stringify(value)} (${reason})`,
        ),
        `Fix: update ${configRelativePath} → plugins.updater.endpoints.`,
      ],
    });
  }

  return { ok: errorBlocks.length === 0, skipped: false, errorBlocks };
}

function main() {
  // Test helper: allow overriding the config path so node:test suites can validate error cases
  // without mutating the real repo configuration.
  const configPath = process.env.FORMULA_TAURI_CONFIG_PATH
    ? path.resolve(repoRoot, process.env.FORMULA_TAURI_CONFIG_PATH)
    : DEFAULT_TAURI_CONFIG_PATH;
  const configRelativePath = path.relative(repoRoot, configPath) || DEFAULT_TAURI_CONFIG_RELATIVE_PATH;

  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Updater config preflight failed", [
      `Failed to read/parse ${configRelativePath}.`,
      `Error: ${msg}`,
    ]);
    return;
  }

  const result = checkUpdaterConfig(config, { configRelativePath });
  if (result.skipped) {
    console.log(
      "Updater config preflight: updater is not active (plugins.updater.active !== true); skipping validation.",
    );
    return;
  }

  if (!result.ok) {
    for (const block of result.errorBlocks) {
      errBlock(block.heading, block.details);
    }
    err("\nUpdater config preflight failed. Fix the errors above before tagging a release.\n");
    return;
  }

  console.log(`Updater config preflight passed (${configRelativePath}).`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(SCRIPT_PATH)) {
  main();
}
