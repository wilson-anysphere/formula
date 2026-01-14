import assert from "node:assert/strict";
import test from "node:test";

test(
  "dependency-free node:test runner skips external deps reached via import.meta.resolve()",
  { skip: typeof import.meta.resolve !== "function" },
  async () => {
    // When `node_modules/` is absent, the node:test runners should skip this file (it depends on an
    // external package). This intentionally loads the dependency by first calling import.meta.resolve()
    // so the runners need to treat `import.meta.resolve("pkg")` as a dependency edge when scanning.
    const resolved = import.meta.resolve("esbuild");
    const mod = await import(resolved);
    const esbuild = mod?.build ? mod : mod?.default;
    assert.equal(typeof esbuild?.build, "function");
  },
);

