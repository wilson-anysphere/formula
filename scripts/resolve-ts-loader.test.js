import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

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
