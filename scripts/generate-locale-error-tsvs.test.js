import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("generate-locale-error-tsvs --check matches committed TSVs", () => {
  const proc = spawnSync(
    process.execPath,
    [path.join(repoRoot, "scripts", "generate-locale-error-tsvs.mjs"), "--check"],
    {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
      },
    },
  );

  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr || proc.stdout);
});

