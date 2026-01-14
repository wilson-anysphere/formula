import assert from "node:assert/strict";
import { statSync } from "node:fs";
import test from "node:test";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function hasTypeScriptDependency() {
  try {
    require.resolve("typescript");
    return true;
  } catch {
    return false;
  }
}

test(
  "node:test runner can execute TS runtime syntax via the transpile loader (parameter properties)",
  { skip: !hasTypeScriptDependency() },
  async () => {
    const mod = await import("./__fixtures__/resolve-ts-loader/param-prop.ts");
    const instance = new mod.ParamProp();
    assert.equal(instance.value, 42);

    const dirMod = await import("./__fixtures__/resolve-ts-loader/dir-import.ts");
    assert.equal(dirMod.getDirValue(), 99);

    const jsxMod = await import("./__fixtures__/resolve-ts-loader/jsx-import.ts");
    assert.equal(jsxMod.getJsxImportValue(), 42);
  },
);

test(
  "TS transpile loader reports TypeScript diagnostics on syntax errors",
  { skip: !hasTypeScriptDependency() },
  async () => {
    await assert.rejects(
      () => import("./__fixtures__/resolve-ts-loader/broken.ts"),
      (err) => {
        assert.ok(err instanceof SyntaxError);
        assert.match(err.message, /Failed to transpile TypeScript module/);
        assert.match(err.message, /broken\.ts/);
        assert.match(err.message, /TS\d+:/);
        return true;
      },
    );
  },
);

test(
  "TS transpile loader resolves @formula/* workspace packages when default resolution fails",
  { skip: !hasTypeScriptDependency() },
  async () => {
    const { resolve: resolveTsLoader } = await import("./resolve-ts-loader.mjs");

    const miss = new Error("ERR_MODULE_NOT_FOUND");
    /** @type {any} */ (miss).code = "ERR_MODULE_NOT_FOUND";
    const failingResolve = async () => {
      throw miss;
    };

    const resolved = await resolveTsLoader("@formula/collab-session#test", { parentURL: import.meta.url }, failingResolve);
    assert.ok(typeof resolved.url === "string" && resolved.url.includes("#test"));

    const url = new URL(resolved.url);
    url.search = "";
    url.hash = "";
    const resolvedPath = fileURLToPath(url);
    assert.ok(resolvedPath.startsWith(repoRoot), "expected resolved file to be within the repo");
    assert.ok(
      resolvedPath.includes(path.join("packages", "collab", "session")),
      "expected resolved file to be under packages/collab/session",
    );
    assert.ok(statSync(resolvedPath).isFile(), "expected resolved workspace entrypoint to exist");
  },
);
