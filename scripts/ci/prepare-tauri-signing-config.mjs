#!/usr/bin/env node
/**
 * GitHub Actions helper for fork-friendly desktop releases.
 *
 * The Tauri config in `apps/desktop/src-tauri/tauri.conf.json` enables macOS signing by default via
 * `bundle.macOS.signingIdentity`. On forks/dry-runs, the required signing secrets are often not
 * configured, which causes the macOS build to fail during bundling/codesign.
 *
 * This script:
 *   1) Detects whether macOS/Windows signing (and macOS notarization) secrets are present.
 *   2) If secrets are missing, patches `tauri.conf.json` to disable the corresponding signing
 *      configuration for the current CI run (so unsigned bundles still build).
 *   3) Exports signing/notarization env vars to subsequent steps via `$GITHUB_ENV` only when all
 *      required secrets are present. This avoids `tauri-action` attempting partial signing flows.
 *
 * It is intended to run in CI before `tauri-apps/tauri-action`.
 */

import { appendFileSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("../..", import.meta.url)));
const configPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const relativeConfigPath = path.relative(repoRoot, configPath);

/**
 * @param {string} name
 */
function envHasValue(name) {
  const value = process.env[name];
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Append an env var for subsequent GitHub Actions steps, handling multiline values safely.
 *
 * @param {string} name
 * @param {string} value
 */
function exportGithubEnv(name, value) {
  const envFile = process.env.GITHUB_ENV;
  if (!envFile) return;

  const normalized = value.replace(/\r\n/g, "\n");
  const delimiter = `__FORMULA_ENV_${name}_${Date.now()}_${Math.random().toString(16).slice(2)}__`;
  appendFileSync(envFile, `${name}<<${delimiter}\n${normalized}\n${delimiter}\n`, "utf8");
}

/**
 * @param {string} message
 */
function log(message) {
  console.log(`[prepare-tauri-signing-config] ${message}`);
}

function main() {
  const hasAppleCert = envHasValue("APPLE_CERTIFICATE");
  const hasAppleCertPassword = envHasValue("APPLE_CERTIFICATE_PASSWORD");
  const hasMacSigningSecrets = hasAppleCert && hasAppleCertPassword;

  const hasAppleNotarizationSecrets =
    envHasValue("APPLE_ID") && envHasValue("APPLE_PASSWORD") && envHasValue("APPLE_TEAM_ID");

  const hasWindowsCert = envHasValue("WINDOWS_CERTIFICATE");
  const hasWindowsCertPassword = envHasValue("WINDOWS_CERTIFICATE_PASSWORD");
  const hasWindowsSigningSecrets = hasWindowsCert && hasWindowsCertPassword;

  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(
      `[prepare-tauri-signing-config] Failed to read/parse ${relativeConfigPath}: ${msg}`
    );
    process.exitCode = 1;
    return;
  }

  let changed = false;

  if (!hasMacSigningSecrets) {
    const currentIdentity = config?.bundle?.macOS?.signingIdentity;
    if (currentIdentity !== null && currentIdentity !== undefined) {
      config.bundle ??= {};
      config.bundle.macOS ??= {};
      config.bundle.macOS.signingIdentity = null;
      changed = true;
      log(
        `macOS code signing secrets not detected; setting bundle.macOS.signingIdentity=null for this run.`
      );
    } else {
      log(`macOS code signing secrets not detected; config already has signing disabled.`);
    }
  } else {
    log(`macOS code signing secrets detected; leaving bundle.macOS.signingIdentity unchanged.`);
  }

  if (!hasWindowsSigningSecrets) {
    const currentThumbprint = config?.bundle?.windows?.certificateThumbprint;
    if (currentThumbprint !== null && currentThumbprint !== undefined) {
      config.bundle ??= {};
      config.bundle.windows ??= {};
      config.bundle.windows.certificateThumbprint = null;
      changed = true;
      log(
        `Windows code signing secrets not detected; setting bundle.windows.certificateThumbprint=null for this run.`
      );
    } else {
      log(`Windows code signing secrets not detected; config already has signing disabled.`);
    }
  } else {
    log(`Windows code signing secrets detected; leaving bundle.windows.certificateThumbprint unchanged.`);
  }

  if (changed) {
    writeFileSync(configPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");
  }

  // Export env vars for subsequent steps only when fully configured.
  if (hasMacSigningSecrets) {
    exportGithubEnv("APPLE_CERTIFICATE", process.env.APPLE_CERTIFICATE ?? "");
    exportGithubEnv("APPLE_CERTIFICATE_PASSWORD", process.env.APPLE_CERTIFICATE_PASSWORD ?? "");
    if (envHasValue("APPLE_SIGNING_IDENTITY")) {
      exportGithubEnv("APPLE_SIGNING_IDENTITY", process.env.APPLE_SIGNING_IDENTITY ?? "");
    }
  }

  // Notarization should only run when signing is enabled and all notarization creds are present.
  if (hasMacSigningSecrets && hasAppleNotarizationSecrets) {
    exportGithubEnv("APPLE_ID", process.env.APPLE_ID ?? "");
    exportGithubEnv("APPLE_PASSWORD", process.env.APPLE_PASSWORD ?? "");
    exportGithubEnv("APPLE_TEAM_ID", process.env.APPLE_TEAM_ID ?? "");
    log(`macOS notarization credentials detected; notarization will be enabled.`);
  } else {
    log(`macOS notarization credentials not fully configured; notarization will be skipped.`);
  }

  if (hasWindowsSigningSecrets) {
    exportGithubEnv("WINDOWS_CERTIFICATE", process.env.WINDOWS_CERTIFICATE ?? "");
    exportGithubEnv("WINDOWS_CERTIFICATE_PASSWORD", process.env.WINDOWS_CERTIFICATE_PASSWORD ?? "");
  }
}

main();

