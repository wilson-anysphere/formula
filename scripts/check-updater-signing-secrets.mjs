#!/usr/bin/env node
/**
 * Backwards-compatible wrapper: historically the release workflow invoked this script from `scripts/`.
 *
 * The real validation logic now lives in `scripts/ci/check-tauri-updater-secrets.mjs` (it understands
 * the minisign key formats produced by `cargo tauri signer generate`, detects encryption, and checks
 * key-id mismatches).
 */
import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-tauri-updater-secrets.mjs");

const proc = spawnSync(process.execPath, [scriptPath], {
  cwd: repoRoot,
  env: process.env,
  stdio: "inherit",
  encoding: "utf8",
});
if (proc.error) {
  console.error(proc.error);
  process.exit(1);
}
process.exit(proc.status ?? 1);
