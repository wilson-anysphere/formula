import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

test("run-node-tests does not claim filter patterns matched nothing when matched files are skipped", () => {
  const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
  const runnerPath = path.join(repoRoot, "scripts", "run-node-tests.mjs");

  const result = spawnSync(
    process.execPath,
    [
      runnerPath,
      "node-test-runner.templateDynamicImportExternalDeps",
      "node-test-runner.importOptionsExternalDeps",
    ],
    { cwd: repoRoot, encoding: "utf8" },
  );

  assert.equal(result.status, 0);
  const output = `${result.stdout ?? ""}${result.stderr ?? ""}`;
  assert.ok(output.includes("[node:test] Filtering"), "expected runner to print its filtering banner");
  assert.ok(
    !output.includes("No node:test files matched:"),
    "runner should not claim patterns matched nothing when the matching files are skipped",
  );
});

