import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import test from "node:test";
import { fileURLToPath } from "node:url";

test("package export: can import @formula/power-query", async () => {
  const mod = await import("@formula/power-query");
  assert.equal(typeof mod.QueryEngine, "function");
  assert.equal(typeof mod.RefreshManager, "function");
  assert.equal(typeof mod.RefreshOrchestrator, "function");
  assert.equal(typeof mod.CacheManager, "function");
});

test("package export: can import @formula/power-query/node", async () => {
  const mod = await import("@formula/power-query/node");
  assert.equal(typeof mod.createNodeCryptoCacheProvider, "function");
  assert.equal(typeof mod.EncryptedFileSystemCacheStore, "function");
});

const require = createRequire(import.meta.url);
let tscPath = null;
try {
  tscPath = require.resolve("typescript/bin/tsc");
} catch {
  tscPath = null;
}

test(
  "package export: TypeScript can import public API",
  { skip: !tscPath },
  async () => {
    assert.ok(tscPath);
    const fixture = fileURLToPath(new URL("./package-import.smoke.ts", import.meta.url));

    const result = spawnSync(
      process.execPath,
      [
        tscPath,
        "--noEmit",
        "--pretty",
        "false",
        "--target",
        "ES2022",
        "--lib",
        "ES2022,DOM",
        "--module",
        "ESNext",
        "--moduleResolution",
        "Bundler",
        fixture,
      ],
      { encoding: "utf8" },
    );

    if (result.status !== 0) {
      throw new Error([result.stdout, result.stderr].filter(Boolean).join("\n"));
    }
  },
);
