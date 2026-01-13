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
  if (normalized.length === 0 || normalized.length % 4 !== 0) return null;
  if (!/^[A-Za-z0-9+/]+=*$/.test(normalized)) return null;
  try {
    return Buffer.from(normalized, "base64");
  } catch {
    return null;
  }
}

/**
 * TAURI_KEY_PASSWORD is only required when the private key is encrypted.
 *
 * We cannot fully validate the key format without invoking the Tauri CLI, so we use a conservative
 * heuristic:
 * - If TAURI_KEY_PASSWORD is empty AND the private key decodes to a small (32/64-byte) raw key, we
 *   assume the key is unencrypted and accept an empty password.
 * - Otherwise, require TAURI_KEY_PASSWORD so tagged releases fail early instead of inside
 *   tauri-apps/tauri-action.
 *
 * @param {string} privateKey
 */
function isClearlyUnencryptedPrivateKey(privateKey) {
  const decoded = decodeBase64(privateKey);
  if (!decoded) return false;
  return decoded.length === 32 || decoded.length === 64;
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
  if (passwordMissing && privateKey.trim().length > 0 && !isClearlyUnencryptedPrivateKey(privateKey)) {
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
      `Updater signing secrets preflight passed (TAURI_KEY_PASSWORD is empty; assuming TAURI_PRIVATE_KEY is unencrypted).`
    );
    return;
  }

  console.log(`Updater signing secrets preflight passed.`);
}

main();
