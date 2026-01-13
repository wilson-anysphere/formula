import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("generate-locale-function-tsv --check matches committed TSVs", () => {
  const proc = spawnSync(
    process.execPath,
    [path.join(repoRoot, "scripts", "generate-locale-function-tsv.js"), "--check"],
    {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        // Keep the generated header deterministic, regardless of the parent environment.
        SOURCE_DATE_EPOCH: "0",
      },
    }
  );

  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr || proc.stdout);
});

