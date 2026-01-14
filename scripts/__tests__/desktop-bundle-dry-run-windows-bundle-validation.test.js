import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { extractYamlRunSteps, stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "desktop-bundle-dry-run.yml");

async function readWorkflow() {
  return stripHashComments(await readFile(workflowPath, "utf8"));
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
  let inBlock = false;
  let blockIndent = 0;
  const blockRe = /:[\t ]*[>|][0-9+-]*[\t ]*$/;
  for (; endIdx < lines.length; endIdx += 1) {
    const line = lines[endIdx] ?? "";
    const trimmed = line.trim();
    const lineIndent = line.match(/^\s*/)?.[0]?.length ?? 0;

    if (inBlock) {
      if (trimmed === "") continue;
      if (lineIndent > blockIndent) continue;
      inBlock = false;
    }

    if (trimmed === "") continue;
    if (lineIndent < indent) break;
    if (nextItemRe.test(line)) break;

    if (blockRe.test(line.trimEnd())) {
      inBlock = true;
      blockIndent = lineIndent;
    }
  }
  return lines.slice(startIdx, endIdx).join("\n");
}

test("desktop-bundle-dry-run workflow validates built Windows bundles (MSI + NSIS)", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const stepNeedle = "Validate Windows installer bundles";
  const idx = searchLines.findIndex((line) => line.includes(stepNeedle));
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
  const runSteps = extractYamlRunSteps(text);
  assert.ok(
    runSteps.some((step) => step.script.includes("node scripts/ci/check-desktop-compliance-artifacts.mjs")),
    `Expected ${path.relative(repoRoot, workflowPath)} to run scripts/ci/check-desktop-compliance-artifacts.mjs in a run step.`,
  );
});

test("desktop-bundle-dry-run workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const buildNeedle = "Build desktop bundles (dry run)";
  const buildIdx = searchLines.findIndex((line) => line.includes(buildNeedle));
  assert.ok(
    buildIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${buildNeedle}`,
  );

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const idx = searchLines.findIndex((line) => line.includes(stripNeedle));
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
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const restoreNeedle = "Restore CI-only Tauri config patches";
  const restoreIdx = searchLines.findIndex((line) => line.includes(restoreNeedle));
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
  const diffIdx = searchLines.findIndex((line) => line.includes(diffNeedle));
  assert.ok(
    diffIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${diffNeedle}`,
  );
  assert.ok(
    restoreIdx < diffIdx,
    `Expected ${restoreNeedle} to appear before ${diffNeedle} so CI-only config patches don't fail the reproducibility guard.`,
  );
});

test("desktop-bundle-dry-run workflow uploads bundles without recursive target/** globs", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const needle = "Upload desktop bundles";
  const idx = searchLines.findIndex((line) => line.includes(needle));
  assert.ok(idx >= 0, `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${needle}`);

  const snippet = yamlListItemBlock(lines, idx);

  // Avoid patterns like `target/**/release/bundle/**/*.dmg`: recursive globs over Cargo target dirs
  // can be surprisingly slow in CI. Prefer predictable release/bundle locations.
  assert.doesNotMatch(
    snippet,
    /target\/\*\*\/release\/bundle\/\*\*\/\*\.dmg/i,
    `Expected upload step to avoid recursive target/**/release/bundle/**/*.dmg globs.\nSaw snippet:\n${snippet}`,
  );

  assert.match(
    snippet,
    /apps\/desktop\/src-tauri\/target\/release\/bundle\/dmg\/\*\.dmg/,
    `Expected upload step to include the non-recursive DMG path pattern.\nSaw snippet:\n${snippet}`,
  );
  assert.match(
    snippet,
    /apps\/desktop\/src-tauri\/target\/\*\/release\/bundle\/dmg\/\*\.dmg/,
    `Expected upload step to include the target-triple DMG path pattern.\nSaw snippet:\n${snippet}`,
  );

  // Updater manifests are written to the bundle root (not inside a platform subdir).
  assert.match(
    snippet,
    /apps\/desktop\/src-tauri\/target\/release\/bundle\/latest\.json/,
    `Expected upload step to include the bundle-root latest.json pattern.\nSaw snippet:\n${snippet}`,
  );
});

test("desktop-bundle-dry-run workflow caches Cargo release artifacts without recursive target/** globs", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const needle = "Cache cargo target (release build artifacts)";
  const idx = searchLines.findIndex((line) => line.includes(needle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to include step: ${needle}`,
  );

  const snippet = yamlListItemBlock(lines, idx);
  assert.doesNotMatch(
    snippet,
    /target\/\*\*\/release\/deps/,
    `Expected cache step to avoid recursive target/**/release/deps globs.\nSaw snippet:\n${snippet}`,
  );
  assert.ok(snippet.includes("target/release/deps"), `Expected cache step to include target/release/deps.`);
  assert.ok(snippet.includes("target/*/release/deps"), `Expected cache step to include target/*/release/deps.`);
});
