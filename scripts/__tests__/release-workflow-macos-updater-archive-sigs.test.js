import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { extractYamlRunSteps, stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

async function readWorkflow() {
  return stripHashComments(await readFile(workflowPath, "utf8"));
}

test("release workflow requires signatures for macOS updater tarballs but excludes AppImage tarballs", async () => {
  const text = await readWorkflow();
  const runSteps = extractYamlRunSteps(text);

  const step = runSteps.find((s) => s.script.includes('require_sigs_for_ext "macOS updater archive"'));
  assert.ok(step, `Expected ${path.relative(repoRoot, workflowPath)} to check macOS updater archive signatures in a run step.`);

  const scriptLines = step.script.split(/\r?\n/);
  const idx = scriptLines.findIndex((line) => line.includes('require_sigs_for_ext "macOS updater archive"'));
  assert.ok(idx >= 0);

  const snippet = scriptLines.slice(idx, idx + 3).join("\n");
  assert.ok(
    snippet.includes("'\\\\.(tar\\\\.gz|tgz)$'"),
    `Expected signature check to target tarball suffixes (tar.gz/tgz).\nSaw snippet:\n${snippet}`,
  );
  assert.ok(
    snippet.includes("'\\\\.AppImage\\\\.(tar\\\\.gz|tgz)$'"),
    `Expected signature check to exclude AppImage tarballs.\nSaw snippet:\n${snippet}`,
  );
});
