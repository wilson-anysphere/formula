import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripPythonComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-windows-installer-signatures.py");

test("check-windows-installer-signatures avoids recursive ** globs (perf guardrail)", () => {
  const contents = stripPythonComments(readFileSync(scriptPath, "utf8"));
  assert.doesNotMatch(contents, /\.glob\(\s*["']\*\*\//, "Expected no recursive Path.glob('**/...') usage");
});
