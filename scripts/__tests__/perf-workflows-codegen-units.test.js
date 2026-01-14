import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

const workflows = [
  {
    name: "perf.yml",
    path: path.join(repoRoot, ".github", "workflows", "perf.yml"),
    stepName: "Build desktop binary (release, desktop feature)",
  },
  {
    name: "desktop-memory-perf.yml",
    path: path.join(repoRoot, ".github", "workflows", "desktop-memory-perf.yml"),
    stepName: "Build desktop binary (release, desktop feature)",
  },
];

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

for (const wf of workflows) {
  test(`${wf.name} builds the desktop release binary with codegen-units=1 (matches Cargo.toml release profile)`, async () => {
    const text = stripHashComments(await readFile(wf.path, "utf8"));
    const lines = text.split(/\r?\n/);
    const searchLines = stripYamlBlockScalarBodies(text).split(/\r?\n/);
    const idx = searchLines.findIndex((line) => line.includes(wf.stepName));
    assert.ok(
      idx >= 0,
      `Expected ${path.relative(repoRoot, wf.path)} to contain a step named: ${wf.stepName}`,
    );
    const snippet = yamlListItemBlock(lines, idx);
    assert.match(
      snippet,
      /\bCARGO_PROFILE_RELEASE_CODEGEN_UNITS:\s*["']?1["']?\b/,
      `Expected ${wf.stepName} to set CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1.\nSaw snippet:\n${snippet}`,
    );
  });
}
