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
 * @param {string[]} lines
 * @param {RegExp} needle
 * @param {number} windowSize
 */
function snippetAfter(lines, needle, windowSize = 40) {
  const searchLines = stripYamlBlockScalarBodies(lines.join("\n")).split(/\r?\n/);
  const idx = searchLines.findIndex((line) => needle.test(line));
  assert.ok(idx >= 0, `Expected to find line matching ${needle} in ${path.relative(repoRoot, workflowPath)}`);
  return searchLines.slice(idx, idx + windowSize).join("\n");
}

test("release workflow gates updater secret enforcement on upstream repo or TAURI_PRIVATE_KEY", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const snippet = snippetAfter(lines, /- name:\s+Validate Tauri updater signing secrets/, 25);
  assert.match(snippet, /secrets\.TAURI_PRIVATE_KEY\s*!=\s*''/);
  assert.match(snippet, /github\.repository\s*==\s*'wilson-anysphere\/formula'/);
});

test("release workflow gates publish-updater-manifest job on upstream repo or TAURI_PRIVATE_KEY", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const snippet = snippetAfter(lines, /^  publish-updater-manifest:/, 20);
  assert.match(snippet, /if:\s*needs\.preflight\.outputs\.upload\s*==\s*'true'/);
  assert.match(snippet, /secrets\.TAURI_PRIVATE_KEY\s*!=\s*''/);
});

test("release workflow gates verify-updater-manifest job on upstream repo or TAURI_PRIVATE_KEY", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const snippet = snippetAfter(lines, /^  verify-updater-manifest:/, 25);
  assert.match(snippet, /if:\s*needs\.preflight\.outputs\.upload\s*==\s*'true'/);
  assert.match(snippet, /secrets\.TAURI_PRIVATE_KEY\s*!=\s*''/);
});

test("release workflow gates verify-release-assets job on upstream repo or TAURI_PRIVATE_KEY", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const snippet = snippetAfter(lines, /^  verify-release-assets:/, 25);
  assert.match(snippet, /if:\s*needs\.preflight\.outputs\.upload\s*==\s*'true'/);
  assert.match(snippet, /secrets\.TAURI_PRIVATE_KEY\s*!=\s*''/);
});

test("release workflow gates checksums job on upstream repo or TAURI_PRIVATE_KEY", async () => {
  const text = await readWorkflow();
  const lines = text.split(/\r?\n/);

  const snippet = snippetAfter(lines, /^  checksums:/, 25);
  assert.match(snippet, /if:\s*needs\.preflight\.outputs\.upload\s*==\s*'true'/);
  assert.match(snippet, /secrets\.TAURI_PRIVATE_KEY\s*!=\s*''/);
});
