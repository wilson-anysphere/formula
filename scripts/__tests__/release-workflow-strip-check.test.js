import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

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

test("release workflow verifies the produced desktop binary is stripped", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const stripNeedle = "Verify desktop binary is stripped (no symbols)";
  const stripIdx = lines.findIndex((line) => line.includes(stripNeedle));
  assert.ok(
    stripIdx >= 0,
    `Expected ${path.relative(repoRoot, workflowPath)} to contain a step named: ${stripNeedle}`,
  );

  // The strip check must happen after the desktop bundling step(s) (`tauri-apps/tauri-action`).
  const tauriActionRe = /^\s*uses:\s*tauri-apps\/tauri-action@/;
  const buildIdx = lines.findIndex((line) => tauriActionRe.test(line));
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

