#!/usr/bin/env node
/**
 * GitHub Actions helper for fork-friendly desktop releases.
 *
 * The desktop app's committed Tauri config intentionally avoids hardcoding a macOS signing identity
 * so local builds work without Developer ID certificates installed. In CI we want:
 *   - unsigned builds to succeed cleanly on forks/dry-runs (no secrets)
 *   - signed/notarized builds to use the explicit `APPLE_SIGNING_IDENTITY` provided by secrets
 *     when available (avoid ambiguous identity selection when multiple certs exist). If the
 *     identity is not provided, we fall back to the generic "Developer ID Application" selector.
 *
 * This script:
 *   1) Detects whether macOS/Windows signing (and macOS notarization) secrets are present.
 *   2) If secrets are missing, patches `tauri.conf.json` to disable the corresponding signing
 *      configuration for the current CI run (so unsigned bundles still build).
 *   3) Exports signing/notarization env vars to subsequent steps via `$GITHUB_ENV` only when all
 *      required secrets are present. This avoids `tauri-action` attempting partial signing flows.
 *
 * If maintainers set the GitHub Actions variable `FORMULA_REQUIRE_CODESIGN=1`, this script switches
 * to enforcement mode and **fails fast** when platform signing secrets are missing (instead of
 * patching `tauri.conf.json` to disable signing).
 *
 * It is intended to run in CI before `tauri-apps/tauri-action`.
 */

import { appendFileSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("../..", import.meta.url)));
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
// Test hook: allow overriding the Tauri config path so unit tests can operate on a temp copy
// instead of mutating the real repo config.
const configPath = process.env.FORMULA_TAURI_CONF_PATH
  ? path.resolve(repoRoot, process.env.FORMULA_TAURI_CONF_PATH)
  : defaultConfigPath;
const relativeConfigPath = path.relative(repoRoot, configPath);

/**
 * @param {string} name
 */
function envHasValue(name) {
  const value = process.env[name];
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * @param {unknown} value
 */
function isTruthy(value) {
  if (value === undefined || value === null) return false;
  const normalized = String(value).trim().toLowerCase();
  return normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on";
}

function getRunnerOs() {
  const envOs = process.env.RUNNER_OS;
  if (typeof envOs === "string" && envOs.trim().length > 0) return envOs.trim();

  // Fallback for local runs.
  switch (process.platform) {
    case "darwin":
      return "macOS";
    case "win32":
      return "Windows";
    default:
      return "Linux";
  }
}

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
  const requireCodesign = isTruthy(process.env.FORMULA_REQUIRE_CODESIGN);
  const runnerOs = getRunnerOs();
  const isMacRunner = runnerOs === "macOS";
  const isWindowsRunner = runnerOs === "Windows";

  if (requireCodesign) {
    const missing = [];

    if (runnerOs === "macOS") {
      const required = [
        "APPLE_CERTIFICATE",
        "APPLE_CERTIFICATE_PASSWORD",
        "APPLE_SIGNING_IDENTITY",
        "APPLE_ID",
        "APPLE_PASSWORD",
        "APPLE_TEAM_ID",
      ];
      for (const name of required) {
        if (!envHasValue(name)) missing.push(name);
      }
    } else if (runnerOs === "Windows") {
      const required = ["WINDOWS_CERTIFICATE", "WINDOWS_CERTIFICATE_PASSWORD"];
      for (const name of required) {
        if (!envHasValue(name)) missing.push(name);
      }
    }

    if (missing.length > 0) {
      errBlock(`Code signing is required (${runnerOs}) but secrets are missing`, [
        `FORMULA_REQUIRE_CODESIGN is enabled, so unsigned artifacts are not allowed.`,
        `Missing/empty GitHub Actions repository secrets (Settings → Secrets and variables → Actions):`,
        ...missing.map((name) => name),
        `To allow unsigned builds again, unset the GitHub Actions variable FORMULA_REQUIRE_CODESIGN.`,
        `See docs/release.md ("Code signing").`,
      ]);
      return;
    }
  }

  const hasAppleCert = envHasValue("APPLE_CERTIFICATE");
  const hasAppleCertPassword = envHasValue("APPLE_CERTIFICATE_PASSWORD");
  // The signing identity is optional; when not present we fall back to a safe default identity.
  const hasMacSigningSecrets = isMacRunner && hasAppleCert && hasAppleCertPassword;

  const hasAppleNotarizationSecrets =
    isMacRunner &&
    envHasValue("APPLE_ID") &&
    envHasValue("APPLE_PASSWORD") &&
    envHasValue("APPLE_TEAM_ID");

  const hasWindowsCert = envHasValue("WINDOWS_CERTIFICATE");
  const hasWindowsCertPassword = envHasValue("WINDOWS_CERTIFICATE_PASSWORD");
  const hasWindowsSigningSecrets = isWindowsRunner && hasWindowsCert && hasWindowsCertPassword;

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

  // Optional override: allow CI/workflows to switch timestamp servers for a single run without
  // committing a config change (useful if a timestamp server is down or blocked by a network proxy).
  //
  // This is intentionally only applied on Windows runners, since it only affects Authenticode signing.
  if (isWindowsRunner) {
    const overrideTimestampUrl = process.env.FORMULA_WINDOWS_TIMESTAMP_URL?.trim() ?? "";
    if (overrideTimestampUrl.length > 0) {
      let parsed;
      try {
        parsed = new URL(overrideTimestampUrl);
      } catch {
        errBlock("Invalid Windows timestamp URL override", [
          `FORMULA_WINDOWS_TIMESTAMP_URL must be a valid absolute URL.`,
          `Got: ${JSON.stringify(overrideTimestampUrl)}`,
        ]);
        return;
      }

      if (parsed.protocol !== "https:") {
        errBlock("Invalid Windows timestamp URL override", [
          `FORMULA_WINDOWS_TIMESTAMP_URL must use https:// (no plaintext HTTP timestamping).`,
          `Got: ${JSON.stringify(overrideTimestampUrl)}`,
        ]);
        return;
      }

      const currentTimestampUrl = config?.bundle?.windows?.timestampUrl;
      if (currentTimestampUrl !== overrideTimestampUrl) {
        config.bundle ??= {};
        config.bundle.windows ??= {};
        config.bundle.windows.timestampUrl = overrideTimestampUrl;
        changed = true;
        log(`Overriding bundle.windows.timestampUrl for this run: ${overrideTimestampUrl}`);
      } else {
        log(`FORMULA_WINDOWS_TIMESTAMP_URL matches bundle.windows.timestampUrl; leaving timestampUrl unchanged.`);
      }
    }
  }

  if (isMacRunner) {
    if (!hasMacSigningSecrets) {
      const currentIdentity = config?.bundle?.macOS?.signingIdentity;
      if (currentIdentity !== null) {
        config.bundle ??= {};
        config.bundle.macOS ??= {};
        config.bundle.macOS.signingIdentity = null;
        changed = true;
        log(
          `macOS code signing secrets not fully configured; setting bundle.macOS.signingIdentity=null for this run.`
        );
      } else {
        log(`macOS code signing secrets not detected; config already has signing disabled.`);
      }
    } else {
      const explicitIdentity = process.env.APPLE_SIGNING_IDENTITY?.trim() ?? "";
      const currentIdentity = config?.bundle?.macOS?.signingIdentity;
      const desiredIdentity =
        explicitIdentity.length > 0
          ? explicitIdentity
          : typeof currentIdentity === "string" && currentIdentity.trim().length > 0
            ? currentIdentity
            : "Developer ID Application";

      if (currentIdentity !== desiredIdentity) {
        config.bundle ??= {};
        config.bundle.macOS ??= {};
        config.bundle.macOS.signingIdentity = desiredIdentity;
        changed = true;
        log(
          explicitIdentity.length > 0
            ? `macOS code signing secrets detected; setting bundle.macOS.signingIdentity to explicit APPLE_SIGNING_IDENTITY for this run.`
            : `macOS code signing secrets detected; setting bundle.macOS.signingIdentity to ${JSON.stringify(desiredIdentity)} for this run.`
        );
      } else {
        log(`macOS code signing secrets detected; leaving bundle.macOS.signingIdentity unchanged.`);
      }
    }
  }

  if (isWindowsRunner) {
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

  // Notarization should only run on macOS runners, and only when signing is enabled and all
  // notarization creds are present.
  if (isMacRunner) {
    if (hasMacSigningSecrets && hasAppleNotarizationSecrets) {
      exportGithubEnv("APPLE_ID", process.env.APPLE_ID ?? "");
      exportGithubEnv("APPLE_PASSWORD", process.env.APPLE_PASSWORD ?? "");
      exportGithubEnv("APPLE_TEAM_ID", process.env.APPLE_TEAM_ID ?? "");
      log(`macOS notarization credentials detected; notarization will be enabled.`);
    } else {
      log(`macOS notarization credentials not fully configured; notarization will be skipped.`);
    }
  }

  if (hasWindowsSigningSecrets) {
    exportGithubEnv("WINDOWS_CERTIFICATE", process.env.WINDOWS_CERTIFICATE ?? "");
    exportGithubEnv("WINDOWS_CERTIFICATE_PASSWORD", process.env.WINDOWS_CERTIFICATE_PASSWORD ?? "");
  }
}

main();
