import assert from "node:assert/strict";
import test from "node:test";

test("dependency-free node:test runner skips external deps imported via import() options arg", async () => {
  // `import(specifier, options)` is valid Node syntax (used for import assertions / attributes).
  // Ensure the dependency scanners treat this as a dependency edge so environments without
  // `node_modules/` skip the file instead of crashing with ERR_MODULE_NOT_FOUND.
  const mod = await import("esbuild", {});
  const esbuild = mod?.build ? mod : mod?.default;
  assert.equal(typeof esbuild?.build, "function");
});

