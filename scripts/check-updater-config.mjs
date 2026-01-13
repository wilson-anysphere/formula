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
