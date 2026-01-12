import assert from "node:assert/strict";
import test from "node:test";

// Include explicit `.ts` import specifiers so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { fromA1, toA1 } from "../packages/spreadsheet-frontend/src/index.ts";

test("spreadsheet-frontend TS sources are importable under Node ESM (strip-types)", () => {
  assert.equal(typeof toA1, "function");
  assert.equal(typeof fromA1, "function");

  assert.equal(toA1(0, 0), "A1");
  assert.deepEqual(fromA1("B2"), { row0: 1, col0: 1 });
});

