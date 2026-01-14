import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const releaseWorkflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

async function readReleaseWorkflow() {
  return stripHashComments(await readFile(releaseWorkflowPath, "utf8"));
}

test("release workflow validates built Windows bundles (MSI + NSIS) after build", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const stepNeedle = "Validate Windows installer bundles";
  const idx = searchLines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to contain a step named: ${stepNeedle}`,
  );

  // Scan a small window; keep this resilient to harmless formatting tweaks while still enforcing
  // that the validator runs with the expected script + args.
  const snippet = lines.slice(idx, idx + 40).join("\n");
  assert.match(snippet, /if:\s*runner\.os\s*==\s*['"]Windows['"]/);
  assert.match(snippet, /validate-windows-bundles\.ps1/);
  assert.match(snippet, /-RequireExe\b/);
  assert.match(snippet, /-RequireMsi\b/);

  // Prefer invoking via `pwsh -ExecutionPolicy Bypass` so CI isn't sensitive to runner policy.
  assert.match(
    snippet,
    /pwsh\s+-NoProfile\s+-ExecutionPolicy\s+Bypass\s+-File\s+(?:\.\/)?scripts\/validate-windows-bundles\.ps1/,
  );
});

test("release workflow validates desktop compliance artifact bundling config (LICENSE/NOTICE)", async () => {
  const text = await readReleaseWorkflow();
  assert.match(
    text,
    /node\s+scripts\/ci\/check-desktop-compliance-artifacts\.mjs\b/,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to run scripts/ci/check-desktop-compliance-artifacts.mjs in preflight.`,
  );
});
