#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
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
      `Updater signing secrets preflight: updater is not active (plugins.updater.active !== true); skipping validation.`
    );
    return;
  }

  const privateKey = process.env.TAURI_PRIVATE_KEY;
  const hasPrivateKey = typeof privateKey === "string" && privateKey.trim().length > 0;
  if (!hasPrivateKey) {
    errBlock(`Missing TAURI_PRIVATE_KEY (Tauri updater signing)`, [
      `plugins.updater.active=true in ${relativeConfigPath}, so release artifacts must be signed.`,
      `Without TAURI_PRIVATE_KEY, the release workflow will upload unsigned updater metadata (missing latest.json.sig), and auto-update will not work.`,
      `TAURI_PRIVATE_KEY is not set (or is empty).`,
      ``,
      `Fix: add the required GitHub Actions repository secrets (Settings → Secrets and variables → Actions):`,
      `TAURI_PRIVATE_KEY (required)`,
      `TAURI_KEY_PASSWORD (if your private key is encrypted / you set a password)`,
      `See docs/release.md ("Tauri updater keys").`,
      ``,
      `If you intentionally do not want auto-update, set ${relativeConfigPath} → plugins.updater.active=false.`,
    ]);
  }

  if (hasPrivateKey) {
    const keyPassword = process.env.TAURI_KEY_PASSWORD;
    if (typeof keyPassword !== "string" || keyPassword.length === 0) {
      console.log(
        [
          ``,
          `Warning: TAURI_KEY_PASSWORD is not set.`,
          `- If your TAURI_PRIVATE_KEY was generated with a password, tauri-action will fail later when signing.`,
          `- If your key is unencrypted, this is OK.`,
          `See docs/release.md ("Tauri updater keys").`,
          ``,
        ].join("\n")
      );
    }
  }

  if (process.exitCode) {
    err(
      `\nUpdater signing secrets preflight failed. Configure the updater signing secrets above before tagging a release.\n`
    );
    return;
  }

  console.log(`Updater signing secrets preflight passed (TAURI_PRIVATE_KEY is set).`);
}

main();
