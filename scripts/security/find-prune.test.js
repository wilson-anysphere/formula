import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

function read(relPath) {
  return readFileSync(path.join(repoRoot, relPath), "utf8");
}

test("security scripts prune target/node_modules in find-based scans (perf guardrail)", () => {
  const configHardening = stripHashComments(read("scripts/security/config_hardening.sh"));
  const securityCi = stripHashComments(read("scripts/security/ci.sh"));

  // Ensure we don't regress back to `-not -path "*/target/*"` filters (which still traverse the
  // tree). Comments are allowed; we're looking for the previous invocation shape.
  assert.doesNotMatch(configHardening, /find \. -type f -name "tauri\.conf\.json"/);
  assert.doesNotMatch(securityCi, /find \. -type f -name \"\\$name\"/);

  // Require `-prune` usage so the scan skips huge build trees in CI.
  assert.ok(
    configHardening.includes("-prune -o") && configHardening.includes('-type f -name "tauri.conf.json"'),
    "Expected config_hardening.sh to use a pruned find for tauri.conf.json discovery",
  );
  assert.ok(
    securityCi.includes("-prune -o") && securityCi.includes('-type f -name "$name"'),
    "Expected security/ci.sh to use a pruned find for lockfile/manifest discovery",
  );

  // Ensure the key build dirs are actually listed in the prune set.
  for (const contents of [configHardening, securityCi]) {
    assert.match(contents, /-name 'node_modules'/);
    assert.match(contents, /-name 'target'/);
    assert.match(contents, /-name '\.git'/);
    assert.match(contents, /-name 'dist'/);
    assert.match(contents, /-name 'build'/);
    assert.match(contents, /-name 'coverage'/);
    assert.match(contents, /-name '\.pnpm-store'/);
    assert.match(contents, /-name '\.turbo'/);
    assert.match(contents, /-name '\.cache'/);
    assert.match(contents, /-name '\.vite'/);
    assert.match(contents, /-name 'playwright-report'/);
    assert.match(contents, /-name 'test-results'/);
  }

  // The node-sast policy checker inside security/ci.sh contains an embedded Python filesystem walk.
  // Guard against regressions that would let it traverse `target/`/`node_modules/`/etc once builds
  // have run.
  const idx = securityCi.indexOf("SKIP_DIRNAMES");
  assert.ok(idx >= 0, "Expected security/ci.sh to define SKIP_DIRNAMES in the node-sast Python snippet.");
  const snippet = securityCi.slice(idx, idx + 400);
  for (const dir of [
    "node_modules",
    "target",
    ".pnpm-store",
    ".turbo",
    ".cache",
    ".vite",
    "dist",
    "build",
    "coverage",
    "test-results",
    "playwright-report",
  ]) {
    assert.ok(
      snippet.includes(dir),
      `Expected SKIP_DIRNAMES in security/ci.sh to include ${dir} (perf guardrail).\nSaw snippet:\n${snippet}`,
    );
  }
});
