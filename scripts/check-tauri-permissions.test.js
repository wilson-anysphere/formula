import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-tauri-permissions.mjs");
const capabilityPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "capabilities", "main.json");

function readCapabilityPermissionIdentifiers() {
  const cap = JSON.parse(fs.readFileSync(capabilityPath, "utf8"));
  assert.ok(Array.isArray(cap?.permissions), "expected capabilities/main.json to have a permissions array");

  /** @type {string[]} */
  const identifiers = [];
  for (const entry of cap.permissions) {
    if (typeof entry === "string") {
      identifiers.push(entry);
    } else if (entry && typeof entry === "object" && typeof entry.identifier === "string") {
      identifiers.push(entry.identifier);
    }
  }
  return identifiers;
}

function runCheckWithCachedPermissionLs(outputText) {
  const tmpDir = mkdtempSync(path.join(os.tmpdir(), "formula-tauri-perms-"));
  const cachePath = path.join(tmpDir, "permission-ls.txt");
  writeFileSync(cachePath, outputText, "utf8");

  try {
    const result = spawnSync(process.execPath, [scriptPath], {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        FORMULA_TAURI_PERMISSION_LS_CACHE_PATH: cachePath,
      },
    });
    return result;
  } finally {
    rmSync(tmpDir, { recursive: true, force: true });
  }
}

test("check-tauri-permissions: passes when toolchain output includes all referenced toolchain identifiers", () => {
  const referenced = readCapabilityPermissionIdentifiers();
  const toolchainIds = referenced.filter((id) => id.includes(":"));
  assert.ok(toolchainIds.length > 0, "expected at least one toolchain permission identifier in capabilities/main.json");

  const output = `${toolchainIds.join("\n")}\n`;
  const result = runCheckWithCachedPermissionLs(output);
  assert.equal(result.status, 0, result.stderr || result.stdout);
});

test("check-tauri-permissions: fails when a capability references an unknown toolchain identifier", () => {
  const referenced = readCapabilityPermissionIdentifiers();
  const toolchainIds = referenced.filter((id) => id.includes(":"));
  assert.ok(toolchainIds.length > 0, "expected at least one toolchain permission identifier in capabilities/main.json");

  const missing = toolchainIds[0];
  const output = `${toolchainIds.filter((id) => id !== missing).join("\n")}\n`;
  const result = runCheckWithCachedPermissionLs(output);
  assert.notEqual(result.status, 0, "expected script to fail when a referenced permission is missing from toolchain output");

  const stderr = result.stderr || "";
  assert.match(stderr, /Unknown permission identifiers:/);
  assert.match(stderr, new RegExp(missing.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  assert.match(stderr, /apps\/desktop\/src-tauri\/capabilities\/main\.json/);
});

test("check-tauri-permissions: does not treat hyphenated substrings as valid identifiers", () => {
  const referenced = readCapabilityPermissionIdentifiers();
  const toolchainIds = referenced.filter((id) => id.includes(":"));
  assert.ok(toolchainIds.length > 0, "expected at least one toolchain permission identifier in capabilities/main.json");

  // Pick a permission and include only its trailing segment (e.g. `allow-open`) in the
  // mocked toolchain output. The checker should still treat the full identifier as missing.
  const full = toolchainIds.find(Boolean);
  assert.ok(full, "expected a toolchain identifier");
  const lastSegment = full.split(":").pop();
  assert.ok(lastSegment && lastSegment !== full);

  const output = `${toolchainIds
    .filter((id) => id !== full)
    .concat([lastSegment])
    .join("\n")}\n`;
  const result = runCheckWithCachedPermissionLs(output);
  assert.notEqual(result.status, 0, "expected script to fail when only a substring is present");

  const stderr = result.stderr || "";
  assert.match(stderr, new RegExp(full.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
});

