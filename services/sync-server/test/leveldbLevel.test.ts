import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

import { requireLevelForYLeveldb } from "../src/leveldbLevel.js";

const require = createRequire(import.meta.url);
let hasYLeveldb = true;
try {
  require.resolve("y-leveldb");
} catch {
  // y-leveldb is an optional dependency.
  hasYLeveldb = false;
}

test(
  "requireLevelForYLeveldb resolves 'level' even when it's only a y-leveldb dependency",
  { skip: hasYLeveldb ? false : "y-leveldb not installed" },
  () => {
  const require = createRequire(import.meta.url);
  let directLevelOk = true;
  try {
    require("level");
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    directLevelOk = false;
    // Under pnpm, sync-server doesn't directly depend on `level`, so this is expected.
    assert.equal(code, "MODULE_NOT_FOUND");
  }

  const resolved = requireLevelForYLeveldb();
  assert.equal(typeof resolved, "function");

  // Sanity check the test expectation in pnpm installs; don't fail if deps are hoisted.
  if (!directLevelOk) {
    const yLeveldbRequire = createRequire(require.resolve("y-leveldb"));
    assert.equal(typeof yLeveldbRequire("level"), "function");
  }
  }
);
