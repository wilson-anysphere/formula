import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { fromA1 as fromA1FromTs } from "../../../packages/spreadsheet-frontend/src/a1.ts";

test(
  "spreadsheet-frontend subpath exports are importable under Node ESM when executing TS sources directly",
  async () => {
    const root = await import("@formula/spreadsheet-frontend");
    const a1 = await import("@formula/spreadsheet-frontend/a1");
    const cache = await import("@formula/spreadsheet-frontend/cache");
    const grid = await import("@formula/spreadsheet-frontend/grid");

    assert.equal(typeof root.fromA1, "function");
    assert.equal(typeof a1.fromA1, "function");
    assert.deepEqual(a1.fromA1("B2"), { row0: 1, col0: 1 });
    assert.equal(typeof fromA1FromTs, "function");

    assert.equal(typeof cache.EngineCellCache, "function");
    assert.equal(typeof grid.EngineGridProvider, "function");
  },
);
