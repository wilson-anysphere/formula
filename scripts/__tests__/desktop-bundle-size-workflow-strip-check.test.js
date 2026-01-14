import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "desktop-bundle-size.yml");

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

test("desktop-bundle-size workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);
  const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);

  const buildNeedle = "Build desktop bundles (Tauri)";
  const buildIdx = searchLines.findIndex((line) => line.includes(buildNeedle));
  assert.ok(
    buildIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${buildNeedle}`,
  );
  const buildSnippet = yamlListItemBlock(lines, buildIdx);
  // `scripts/cargo_agent.sh` sets CARGO_PROFILE_RELEASE_CODEGEN_UNITS based on its job count unless
  // callers override it. Ensure this workflow pins it to 1 so bundle sizes match the repo's
  // Cargo.toml release profile and remain comparable to tagged releases.
  assert.match(
    buildSnippet,
    /\bCARGO_PROFILE_RELEASE_CODEGEN_UNITS:\s*["']?1["']?\b/,
    `Expected the Tauri build step to set CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1.\nSaw snippet:\n${buildSnippet}`,
  );

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const idx = searchLines.findIndex((line) => line.includes(stripNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${stripNeedle}`,
  );
  assert.ok(
    idx > buildIdx,
    `Expected the strip verification step to run after the Tauri build step.`,
  );

  const snippet = yamlListItemBlock(lines, idx);
  assert.match(
    snippet,
    /run:\s*python(?:3)?\s+scripts\/verify_desktop_binary_stripped\.py\b/,
    `Expected the strip verification step to invoke scripts/verify_desktop_binary_stripped.py.\nSaw snippet:\n${snippet}`,
  );
});

test("desktop-bundle-size workflow validates desktop compliance artifact bundling config (LICENSE/NOTICE)", async () => {
  const text = await readWorkflow();
  assert.match(
    text,
    /node\s+scripts\/ci\/check-desktop-compliance-artifacts\.mjs\b/,
    `Expected ${path.relative(repoRoot, workflowPath)} to run scripts/ci/check-desktop-compliance-artifacts.mjs before building bundles.`,
  );
});
