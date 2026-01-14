import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("normalize-locale-function-sources --check matches committed source JSONs", () => {
  const proc = spawnSync(
    process.execPath,
    [path.join(repoRoot, "scripts", "normalize-locale-function-sources.js"), "--check"],
    {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        // Keep output deterministic if callers have unusual locale settings.
        LANG: "C",
        LC_ALL: "C",
      },
    }
  );

  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr || proc.stdout);
});

