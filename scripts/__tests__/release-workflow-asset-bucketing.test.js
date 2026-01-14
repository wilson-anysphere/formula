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

test("release workflow buckets Linux AppImage tarballs as linux (includes *.AppImage.tgz)", async () => {
  const text = await readWorkflow();
  const runSteps = extractYamlRunSteps(text);

  // The `bucket_for()` helper groups assets for a step summary. Ensure AppImage tarballs are treated
  // as Linux (not mis-grouped as macOS tarballs).
  const step = runSteps.find((s) => s.script.includes('elif [[ "$base" == *.deb'));
  assert.ok(step, `Expected ${path.relative(repoRoot, workflowPath)} to define bucket_for() linux branch in a run step.`);

  const scriptLines = step.script.split(/\r?\n/);
  const idx = scriptLines.findIndex((line) => line.includes('elif [[ "$base" == *.deb'));
  assert.ok(idx >= 0);

  const snippet = scriptLines.slice(idx, idx + 6).join("\n");
  assert.match(
    snippet,
    /\$base\" == \*\.AppImage\.tgz\b/,
    `Expected linux bucket_for() branch to include *.AppImage.tgz.\nSaw snippet:\n${snippet}`,
  );
});
