import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
//
// (The package itself exports TypeScript sources directly.)
import { getSheetNameValidationErrorMessage as getMessageDirect } from "../packages/workbook-backend/src/index.ts";

test("workbook-backend is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/workbook-backend");

  assert.equal(typeof mod.getSheetNameValidationErrorMessage, "function");
  assert.equal(mod.getSheetNameValidationErrorMessage("Sheet1"), null);
  assert.equal(getMessageDirect("Sheet1"), null);
});

