import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "desktop-bundle-dry-run.yml");

async function readWorkflow() {
  return await readFile(workflowPath, "utf8");
}

test("desktop-bundle-dry-run workflow validates built Windows bundles (MSI + NSIS)", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const stepNeedle = "Validate Windows installer bundles";
  const idx = lines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(idx >= 0, `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${stepNeedle}`);

  const snippet = lines.slice(idx, idx + 40).join("\n");
  assert.match(snippet, /if:\s*runner\.os\s*==\s*['"]Windows['"]/);
  assert.match(snippet, /validate-windows-bundles\.ps1/);
  assert.match(snippet, /-RequireExe\b/);
  assert.match(snippet, /-RequireMsi\b/);
  assert.match(
    snippet,
    /pwsh\s+-NoProfile\s+-ExecutionPolicy\s+Bypass\s+-File\s+(?:\.\/)?scripts\/validate-windows-bundles\.ps1/,
  );
});

test("desktop-bundle-dry-run workflow restores tauri.conf.json before asserting a clean git diff", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const restoreNeedle = "Restore CI-only Tauri config patches";
  const restoreIdx = lines.findIndex((line) => line.includes(restoreNeedle));
  assert.ok(
    restoreIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${restoreNeedle}`,
  );

  const restoreSnippet = lines.slice(restoreIdx, restoreIdx + 20).join("\n");
  assert.match(
    restoreSnippet,
    /git restore --source=HEAD -- apps\/desktop\/src-tauri\/tauri\.conf\.json/,
    `Expected restore step to reset apps/desktop/src-tauri/tauri.conf.json so the git diff guard does not fail after CI-only patches.`,
  );

  const diffNeedle = "Fail if the build modified tracked files";
  const diffIdx = lines.findIndex((line) => line.includes(diffNeedle));
  assert.ok(
    diffIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${diffNeedle}`,
  );
  assert.ok(
    restoreIdx < diffIdx,
    `Expected ${restoreNeedle} to appear before ${diffNeedle} so CI-only config patches don't fail the reproducibility guard.`,
  );
});
