import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

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

    // Detect YAML block scalars (e.g. `run: |`, `path: >-`) so `- ...` inside the scalar body
    // doesn't terminate the list item early.
    if (blockRe.test(line.trimEnd())) {
      inBlock = true;
      blockIndent = lineIndent;
    }
  }
  return lines.slice(startIdx, endIdx).join("\n");
}

test("release workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const stripIdx = searchLines.findIndex((line) => line.includes(stripNeedle));
  assert.ok(
    stripIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${stripNeedle}`,
  );

  // The strip check must happen after the desktop bundling step(s) (`tauri-apps/tauri-action`).
  const tauriActionRe = /^\s*uses:\s*tauri-apps\/tauri-action@/;
  const buildIdx = searchLines.findIndex((line) => tauriActionRe.test(line));
  assert.ok(
    buildIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to use tauri-apps/tauri-action for bundling.`,
  );
  assert.ok(stripIdx > buildIdx, `Expected the strip verification step to run after bundling.`);

  const snippet = yamlListItemBlock(lines, stripIdx);
  assert.match(
    snippet,
    /run:\s*python(?:3)?\s+scripts\/verify_desktop_binary_stripped\.py\b/,
    `Expected the strip verification step to invoke scripts/verify_desktop_binary_stripped.py.\nSaw snippet:\n${snippet}`,
  );
});

test("release workflow uploads dry-run desktop bundles via non-recursive release/bundle globs", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const needle = "Upload desktop bundles (dry run)";
  const idx = searchLines.findIndex((line) => line.includes(needle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${needle}`,
  );

  const snippet = yamlListItemBlock(lines, idx);

  // Avoid `**/release/bundle/**` patterns here: Cargo target directories can be large, and recursive
  // globbing during artifact upload can noticeably slow CI.
  assert.doesNotMatch(
    snippet,
    /target\/\*\*\/release\/bundle\/\*\*/,
    `Expected dry-run bundle upload step to avoid target/**/release/bundle/** patterns.\nSaw snippet:\n${snippet}`,
  );

  for (const required of [
    "target/release/bundle/**",
    "target/*/release/bundle/**",
    "apps/desktop/src-tauri/target/release/bundle/**",
    "apps/desktop/src-tauri/target/*/release/bundle/**",
  ]) {
    assert.ok(
      snippet.includes(required),
      `Expected dry-run bundle upload step to include "${required}".\nSaw snippet:\n${snippet}`,
    );
  }
});

test("release workflow caches Cargo release artifacts without recursive target/** globs", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const needle = "Cache cargo target (release build artifacts)";
  const idx = searchLines.findIndex((line) => line.includes(needle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${needle}`,
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
