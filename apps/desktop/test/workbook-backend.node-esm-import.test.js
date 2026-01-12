import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { WORKBOOK_BACKEND_REQUIRED_METHODS as requiredFromTs } from "../../../packages/workbook-backend/src/index.ts";

test("workbook-backend is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/workbook-backend");

  assert.ok(mod && typeof mod === "object");
  assert.deepEqual(mod.WORKBOOK_BACKEND_REQUIRED_METHODS, requiredFromTs);

  assert.equal(typeof mod.getSheetNameValidationError, "function");
  assert.equal(typeof mod.getSheetNameValidationErrorMessage, "function");
});
