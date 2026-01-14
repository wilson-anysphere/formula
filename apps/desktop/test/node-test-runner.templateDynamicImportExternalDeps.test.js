import assert from "node:assert/strict";
import test from "node:test";

test("dependency-free node:test runner skips external deps reached via template-literal dynamic imports", async () => {
  // This import uses a template literal + interpolation (common for cache-busting query strings).
  // The node:test runners should still treat it as a dependency edge and detect external deps
  // inside the imported module.
  const { loadEsbuild } = await import(`./fixtures/templateDynamicImportExternalDep.js?cacheBust=${Date.now()}`);
  const esbuild = await loadEsbuild();
  assert.equal(typeof esbuild?.build, "function");
});

