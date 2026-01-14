import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

test("dependency-free node:test runner skips external deps reached via createRequire().resolve()", async () => {
  // When `node_modules/` is absent, the node:test runners should skip this file (it depends on
  // an external package). This intentionally loads the dependency through a createRequire alias
  // so the runners need to detect `const req = createRequire(...); req.resolve("pkg")` patterns.
  const requireFromHere = createRequire(import.meta.url);
  const esbuildPath = requireFromHere.resolve("esbuild");
  const mod = await import(pathToFileURL(esbuildPath).href);
  const esbuild = mod?.build ? mod : mod?.default;
  assert.equal(typeof esbuild?.build, "function");
});

