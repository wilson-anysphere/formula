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

/**
 * Extracts a YAML list item's block by scanning forward until either:
 * - the next list item at the same indentation, or
 * - an outdent (indentation decreases), which indicates the list ended.
 *
 * This keeps tests resilient to harmless workflow formatting churn without
 * needing a YAML parser.
 *
 * @param {string[]} lines
 * @param {number} startIdx
 */
function yamlListItemBlock(lines, startIdx) {
  const startLine = lines[startIdx] ?? "";
  const indent = startLine.match(/^\s*/)?.[0]?.length ?? 0;
  const nextItemRe = new RegExp(`^\\s{${indent}}-\\s+`);

  let endIdx = startIdx + 1;
  for (; endIdx < lines.length; endIdx += 1) {
    const line = lines[endIdx] ?? "";
    if (line.trim() === "") continue;
    const lineIndent = line.match(/^\s*/)?.[0]?.length ?? 0;
    if (lineIndent < indent) break;
    if (nextItemRe.test(line)) break;
  }
  return lines.slice(startIdx, endIdx).join("\n");
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

test("desktop-bundle-dry-run workflow validates desktop compliance artifact bundling config (LICENSE/NOTICE)", async () => {
  const text = await readWorkflow();
  assert.match(
    text,
    /node\s+scripts\/ci\/check-desktop-compliance-artifacts\.mjs\b/,
    `Expected ${path.relative(repoRoot, workflowPath)} to run scripts/ci/check-desktop-compliance-artifacts.mjs in preflight.`,
  );
});

test("desktop-bundle-dry-run workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const buildNeedle = "Build desktop bundles (dry run)";
  const buildIdx = lines.findIndex((line) => line.includes(buildNeedle));
  assert.ok(
    buildIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${buildNeedle}`,
  );

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const idx = lines.findIndex((line) => line.includes(stripNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${stripNeedle}`,
  );
  assert.ok(idx > buildIdx, `Expected strip verification to run after the Tauri build step.`);

  const snippet = yamlListItemBlock(lines, idx);
  assert.match(
    snippet,
    /run:\s*python(?:3)?\s+scripts\/verify_desktop_binary_stripped\.py\b/,
    `Expected strip verification step to invoke scripts/verify_desktop_binary_stripped.py.\nSaw snippet:\n${snippet}`,
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
