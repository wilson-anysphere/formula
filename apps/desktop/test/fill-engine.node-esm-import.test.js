import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { computeFillEdits as computeFillEditsFromTs } from "../../../packages/fill-engine/src/index.ts";

test("fill-engine is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/fill-engine");

  assert.equal(typeof mod.computeFillEdits, "function");
  assert.equal(typeof mod.shiftFormulaA1, "function");
  assert.equal(typeof computeFillEditsFromTs, "function");

  assert.equal(mod.shiftFormulaA1("=A1", 1, 0), "=A2");
});

