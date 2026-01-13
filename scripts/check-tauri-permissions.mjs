#!/usr/bin/env node
/**
 * Validate that Tauri capability files only reference permission identifiers that
 * exist in the installed Tauri toolchain.
 *
 * Why this exists:
 * - `apps/desktop/src-tauri/capabilities/*.json` can drift when Tauri/plugins are upgraded.
 * - The canonical list of valid permission identifiers comes from the installed toolchain:
 *   `cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls`
 *
 * This script runs that command, parses the permission identifiers it reports, then validates
 * every permission entry in the capability JSON files:
 * - string-form permissions: validate the string value
 * - object-form permissions: validate `permission.identifier`
 */

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const capabilitiesDir = path.join(repoRoot, "apps", "desktop", "src-tauri", "capabilities");

function stripAnsi(text) {
  // Covers common ANSI SGR + cursor control sequences.
  // eslint-disable-next-line no-control-regex
  return text.replace(/\x1b\[[0-9;]*[A-Za-z]/g, "");
}

function runTauriPermissionLs() {
  const cmd = "bash";
  const args = [
    "-lc",
    // Use the repo cargo wrapper so agent-specific safe defaults apply.
    "cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls",
  ];
  const result = spawnSync(cmd, args, {
    cwd: repoRoot,
    encoding: "utf8",
    maxBuffer: 20 * 1024 * 1024,
    env: {
      ...process.env,
      // Keep output stable/parseable.
      NO_COLOR: "1",
      CARGO_TERM_COLOR: "never",
    },
  });

  if (result.status !== 0) {
    const stdout = (result.stdout ?? "").trim();
    const stderr = (result.stderr ?? "").trim();
    if (stdout) process.stderr.write(`${stdout}\n`);
    if (stderr) process.stderr.write(`${stderr}\n`);
    const hint = [
      "Failed to list Tauri permissions via `cargo tauri permission ls`.",
      "",
      "Common causes:",
      "- `cargo-tauri` is not installed (install with: bash scripts/cargo_agent.sh install cargo-tauri --locked)",
      "- Linux WebView deps are missing (gtk/webkit2gtk; see `.github/workflows/ci.yml` desktop-tauri-check job)",
      "",
      "Manual debug command:",
      "  cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls",
    ].join("\n");
    throw new Error(hint);
  }

  return stripAnsi(String(result.stdout ?? ""));
}

function parsePermissionIdentifiers(permissionLsOutput) {
  // The CLI output format has changed across Tauri versions (bullets, tables, etc).
  // We parse conservatively by extracting tokens that *look like* permission identifiers:
  // - core/plugin permissions: `<segment>:<segment>[:<segment>]...`
  // - application permissions: `allow-invoke` (hyphenated, no `:`)
  const identifiers = new Set();

  const colonIdentifier = /\b[a-z0-9][a-z0-9_-]*(?::[a-z0-9_-]+)+\b/g;
  // Exclude sub-matches inside colon identifiers (e.g. `allow-listen` inside `core:event:allow-listen`).
  const hyphenIdentifier = /(?<!:)\b[a-z0-9][a-z0-9_-]*-[a-z0-9_-]+(?:-[a-z0-9_-]+)*\b(?!:)/g;

  for (const match of permissionLsOutput.matchAll(colonIdentifier)) {
    identifiers.add(match[0]);
  }
  for (const match of permissionLsOutput.matchAll(hyphenIdentifier)) {
    identifiers.add(match[0]);
  }

  // Sanity check: if we couldn't parse anything, fail loudly rather than silently passing.
  if (identifiers.size === 0) {
    const sample = permissionLsOutput.trim().slice(0, 2000);
    throw new Error(
      [
        "Unable to parse any permission identifiers from `cargo tauri permission ls` output.",
        "",
        "Output sample:",
        sample || "(empty)",
      ].join("\n"),
    );
  }

  return identifiers;
}

function collectCapabilityPermissionRefs(capabilityJson, relPath) {
  const refs = [];

  const permissions = capabilityJson?.permissions;
  if (!Array.isArray(permissions)) {
    refs.push({
      file: relPath,
      pointer: "permissions",
      identifier: null,
      error: "expected `permissions` to be an array",
    });
    return refs;
  }

  for (let i = 0; i < permissions.length; i++) {
    const entry = permissions[i];
    if (typeof entry === "string") {
      refs.push({
        file: relPath,
        pointer: `permissions[${i}]`,
        identifier: entry,
        error: null,
      });
      continue;
    }

    if (entry && typeof entry === "object") {
      const identifier = entry.identifier;
      if (typeof identifier === "string") {
        refs.push({
          file: relPath,
          pointer: `permissions[${i}].identifier`,
          identifier,
          error: null,
        });
      } else {
        refs.push({
          file: relPath,
          pointer: `permissions[${i}].identifier`,
          identifier: null,
          error: "expected permission object to have a string `identifier` field",
        });
      }
      continue;
    }

    refs.push({
      file: relPath,
      pointer: `permissions[${i}]`,
      identifier: null,
      error: `expected permission entry to be a string or object (got ${entry === null ? "null" : typeof entry})`,
    });
  }

  return refs;
}

function main() {
  if (!fs.existsSync(capabilitiesDir)) {
    throw new Error(`Capabilities directory not found: ${capabilitiesDir}`);
  }

  const permissionLsOutput = runTauriPermissionLs();
  const validIdentifiers = parsePermissionIdentifiers(permissionLsOutput);

  const capabilityFiles = fs
    .readdirSync(capabilitiesDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
    .map((entry) => path.join(capabilitiesDir, entry.name))
    .sort();

  const problems = [];

  for (const filePath of capabilityFiles) {
    const relPath = path.relative(repoRoot, filePath).replace(/\\/g, "/");
    const jsonText = fs.readFileSync(filePath, "utf8");
    let parsed;
    try {
      parsed = JSON.parse(jsonText);
    } catch (err) {
      problems.push({
        file: relPath,
        pointer: "",
        identifier: null,
        error: `invalid JSON: ${(err && err.message) || String(err)}`,
      });
      continue;
    }

    const refs = collectCapabilityPermissionRefs(parsed, relPath);
    for (const ref of refs) {
      if (ref.error) {
        problems.push(ref);
        continue;
      }
      if (!validIdentifiers.has(ref.identifier)) {
        problems.push({
          ...ref,
          error: "unknown permission identifier (not reported by `cargo tauri permission ls`)",
        });
      }
    }
  }

  if (problems.length > 0) {
    const unknown = problems.filter((p) => p.identifier);
    const structural = problems.filter((p) => !p.identifier);

    const lines = [];
    lines.push("Tauri capability permission validation failed.");
    lines.push("");

    if (unknown.length > 0) {
      lines.push("Unknown permission identifiers:");
      for (const p of unknown) {
        lines.push(`- ${p.file} ${p.pointer}: "${p.identifier}" (${p.error})`);
      }
      lines.push("");
    }

    if (structural.length > 0) {
      lines.push("Malformed capability permission entries:");
      for (const p of structural) {
        lines.push(`- ${p.file} ${p.pointer}: ${p.error}`);
      }
      lines.push("");
    }

    lines.push(`Toolchain reported ${validIdentifiers.size} permission identifiers.`);
    lines.push("To list them manually:");
    lines.push("  cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls");
    lines.push("");

    process.stderr.write(`${lines.join("\n")}\n`);
    process.exit(1);
  }

  process.stdout.write(
    `OK: all capability permission identifiers exist in the installed Tauri toolchain (${validIdentifiers.size} identifiers).\n`,
  );
}

try {
  main();
} catch (err) {
  const message = err && typeof err === "object" && "message" in err ? err.message : String(err);
  process.stderr.write(`${message}\n`);
  process.exit(1);
}

