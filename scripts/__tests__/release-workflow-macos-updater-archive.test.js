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

test("release workflow accepts macOS updater tarballs beyond strict *.app.tar.gz (allow *.tar.gz/*.tgz; exclude AppImage tarballs)", async () => {
  const text = await readWorkflow();
  const runSteps = extractYamlRunSteps(text);

  const step = runSteps.find((s) => s.script.includes("mac_archive_count="));
  assert.ok(step, `Expected ${path.relative(repoRoot, workflowPath)} to define mac_archive_count in a run step.`);

  const scriptLines = step.script.split(/\r?\n/);
  const idx = scriptLines.findIndex((line) => line.includes("mac_archive_count="));
  assert.ok(idx >= 0);

  const snippet = scriptLines.slice(idx, idx + 10).join("\n");
  const normalized = snippet.replace(/\s+/g, "");

  // Guard should accept any tarball suffix, not just `*.app.tar.gz`.
  const hasGenericTarballMatcher =
    normalized.includes('select(test("\\\\.(tar\\\\.gz|tgz)$";"i"))') ||
    (normalized.includes('select(test("\\\\.tar\\\\.gz$";"i"))') &&
      normalized.includes('select(test("\\\\.tgz$";"i"))'));
  assert.ok(
    hasGenericTarballMatcher,
    `Expected mac_archive_count jq filter to accept *.tar.gz/*.tgz (not just *.app.tar.gz).\nSaw snippet:\n${snippet}`,
  );

  // Guard should exclude Linux AppImage tarballs so they don't satisfy the macOS updater requirement.
  const excludesAppImageTarballs =
    normalized.includes('select(test("\\\\.AppImage\\\\.(tar\\\\.gz|tgz)$";"i")|not)') ||
    (normalized.includes('select(test("\\\\.AppImage\\\\.tar\\\\.gz$";"i")|not)') &&
      normalized.includes('select(test("\\\\.AppImage\\\\.tgz$";"i")|not)'));
  assert.ok(
    excludesAppImageTarballs,
    `Expected mac_archive_count jq filter to exclude *.AppImage.(tar.gz|tgz).\nSaw snippet:\n${snippet}`,
  );

  // Regression guard: do not revert to matching only `\\.app\\.tar\\.gz$`.
  assert.equal(
    normalized.includes('select(test("\\\\.app\\\\.tar\\\\.gz$";"i"))'),
    false,
    `Expected mac_archive_count jq filter not to require *.app.tar.gz exclusively.\nSaw snippet:\n${snippet}`,
  );
});
