import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-macos-entitlements.mjs");

test("macOS entitlements guardrail passes against the repo config", () => {
  const proc = spawnSync(process.execPath, [scriptPath, "--root", repoRoot], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  assert.equal(proc.status, 0, proc.stderr || proc.stdout);
});

