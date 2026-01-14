import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const releaseWorkflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

async function readReleaseWorkflow() {
  return await readFile(releaseWorkflowPath, "utf8");
}

test("release workflow includes a Linux ARM64 (aarch64) build runner in the matrix", async () => {
  const text = await readReleaseWorkflow();

  const armRunnerRe = /platform:\s*ubuntu-24\.04-arm(?:64)?\b/;
  assert.match(
    text,
    armRunnerRe,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to include an Ubuntu ARM64 runner (ubuntu-24.04-arm64).`,
  );

  const idx = text.search(armRunnerRe);
  assert.ok(idx >= 0);
  const window = text.slice(idx, idx + 600);
  assert.match(window, /cache_target:\s*aarch64-unknown-linux-gnu\b/);
  assert.match(window, /tauri_args:\s*\"--bundles\s+appimage,deb(?:,rpm)?\"/);
});

test("release workflow validates .deb packages on all Linux runners (x86_64 + arm64)", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);

  const stepNeedle = "Verify Linux .deb package (deps + ldd + desktop integration)";
  const idx = lines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to contain a step named: ${stepNeedle}`,
  );

  // Scan the next few lines to find the step's `if:` guard.
  const snippet = lines.slice(idx, idx + 12).join("\n");
  assert.match(
    snippet,
    /if:\s*runner\.os\s*==\s*['"]Linux['"]/,
    `Expected the Linux .deb validation step to be gated by runner.os == 'Linux' (so it runs for both x86_64 and arm64).`,
  );
});

