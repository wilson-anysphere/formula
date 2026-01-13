#!/usr/bin/env node
/**
 * Thin wrapper around `scripts/ci/validate-updater-manifest.mjs`.
 *
 * The release workflow uses this entrypoint to validate the `latest.json` updater
 * manifest on a draft GitHub Release contains all expected OS/arch targets.
 *
 * Keeping the wrapper in `scripts/` (vs `scripts/ci/`) makes it easier to discover
 * and matches the file path referenced by platform/release docs.
 */
import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const ciValidatorPath = path.join(scriptDir, "ci", "validate-updater-manifest.mjs");

const res = spawnSync(process.execPath, [ciValidatorPath, ...process.argv.slice(2)], {
  stdio: "inherit",
});

if (res.error) {
  console.error(
    `Failed to run updater manifest validator (${ciValidatorPath}): ${res.error.message}`,
  );
  process.exit(1);
}

process.exit(res.status ?? 1);

