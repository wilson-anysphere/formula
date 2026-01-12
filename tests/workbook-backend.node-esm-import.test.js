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

  // Basic happy-path.
  assert.equal(mod.getSheetNameValidationErrorMessage("Sheet1"), null);
  assert.equal(getMessageDirect("Sheet1"), null);

  // Runtime correctness smoke tests (these ensure the JS runtime implementation stays
  // in sync with the TS sources that provide types).
  assert.equal(mod.getSheetNameValidationErrorMessage(""), "sheet name cannot be blank");
  assert.equal(getMessageDirect(""), "sheet name cannot be blank");

  assert.equal(mod.getSheetNameValidationErrorMessage("'Budget"), "sheet name cannot begin or end with an apostrophe");
  assert.equal(getMessageDirect("'Budget"), "sheet name cannot begin or end with an apostrophe");

  assert.equal(mod.getSheetNameValidationErrorMessage("Bad:Name"), "sheet name contains invalid character `:`");
  assert.equal(getMessageDirect("Bad:Name"), "sheet name contains invalid character `:`");

  assert.equal(mod.getSheetNameValidationErrorMessage("budget", { existingNames: ["Budget"] }), "sheet name already exists");
  assert.equal(getMessageDirect("budget", { existingNames: ["Budget"] }), "sheet name already exists");
});
