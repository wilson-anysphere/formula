import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "security", "ci.sh");

test("security/ci.sh runs gitleaks in git mode with bounded log opts (perf guardrail)", () => {
  const contents = stripHashComments(readFileSync(scriptPath, "utf8"));
  assert.doesNotMatch(
    contents,
    /\bgitleaks detect[\s\S]*--no-git\b/,
    "Expected gitleaks scan to avoid --no-git (which scans the whole working tree incl. build outputs)",
  );
  assert.match(contents, /\bgitleaks detect[\s\S]*--log-opts\b/);
  assert.match(contents, /FORMULA_GITLEAKS_LOG_OPTS/);
});
