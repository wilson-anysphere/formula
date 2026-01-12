import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { normalizeFormulaText as normalizeFromTs } from "../../../packages/engine/src/backend/formula.ts";

test("engine/backend/formula is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/engine/backend/formula");

  assert.equal(typeof mod.isFormulaInput, "function");
  assert.equal(typeof mod.normalizeFormulaText, "function");
  assert.equal(typeof mod.normalizeFormulaTextOpt, "function");
  assert.equal(typeof normalizeFromTs, "function");

  assert.equal(mod.normalizeFormulaText("= 1 + 1"), "=1 + 1");
});

