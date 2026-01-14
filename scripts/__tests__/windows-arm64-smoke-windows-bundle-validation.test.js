import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "windows-arm64-smoke.yml");

async function readWorkflow() {
  return stripHashComments(await readFile(workflowPath, "utf8"));
}

test("windows-arm64-smoke workflow validates built bundles via validate-windows-bundles.ps1", async () => {
  const text = await readWorkflow();
  assert.match(text, /validate-windows-bundles\.ps1/);
  assert.match(text, /-RequireExe\b/);
  assert.match(text, /-RequireMsi\b/);
  assert.match(
    text,
    /pwsh\s+-NoProfile\s+-ExecutionPolicy\s+Bypass\s+-File\s+(?:\.\/)?scripts\/validate-windows-bundles\.ps1/,
  );
});

test("windows-arm64-smoke workflow validates desktop compliance artifact bundling config (LICENSE/NOTICE)", async () => {
  const text = await readWorkflow();
  assert.match(text, /node\s+scripts\/ci\/check-desktop-compliance-artifacts\.mjs\b/);
});

test("windows-arm64-smoke workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const buildNeedle = "Build Windows ARM64 bundles (MSI + NSIS)";
  const buildIdx = searchLines.findIndex((line) => line.includes(buildNeedle));
  assert.ok(
    buildIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${buildNeedle}`,
  );

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const stripIdx = searchLines.findIndex((line) => line.includes(stripNeedle));
  assert.ok(
    stripIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${stripNeedle}`,
  );
  assert.ok(stripIdx > buildIdx, `Expected strip verification to run after the build step.`);

  const snippet = lines.slice(stripIdx, stripIdx + 10).join("\n");
  assert.match(snippet, /run:\s*python(?:3)?\s+scripts\/verify_desktop_binary_stripped\.py\b/);
});
