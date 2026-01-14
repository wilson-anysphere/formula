import assert from "node:assert/strict";
import test from "node:test";

test("dependency-free node:test runner skips comment-wrapped dynamic imports of external deps", async () => {
  // If `node_modules` is absent, `apps/desktop/scripts/run-node-tests.mjs` should skip
  // this file (it depends on an external package). This line intentionally includes a
  // comment inside the `import()` call to ensure the runner strips comments while
  // scanning for dependency specifiers.
  const mod = await import(/* @vite-ignore */ "esbuild");
  const esbuild = mod?.build ? mod : mod?.default;
  assert.equal(typeof esbuild?.build, "function");
});

