import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { stripComments } from "../apps/desktop/test/sourceTextUtils.js";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
//
// (The package itself exports TypeScript sources directly.)
import { getSheetNameValidationErrorMessage as getMessageFromIndex } from "../packages/workbook-backend/src/index.ts";
import {
  EXCEL_MAX_SHEET_NAME_LEN as TS_EXCEL_MAX_SHEET_NAME_LEN,
  INVALID_SHEET_NAME_CHARACTERS as TS_INVALID_SHEET_NAME_CHARACTERS,
  getSheetNameValidationErrorMessage as getMessageFromTs,
} from "../packages/workbook-backend/src/sheetNameValidation.ts";
import {
  EXCEL_MAX_SHEET_NAME_LEN as JS_EXCEL_MAX_SHEET_NAME_LEN,
  INVALID_SHEET_NAME_CHARACTERS as JS_INVALID_SHEET_NAME_CHARACTERS,
  getSheetNameValidationErrorMessage as getMessageFromJs,
} from "../packages/workbook-backend/src/sheetNameValidation.js";

test("workbook-backend is importable under Node ESM when executing TS sources directly", async () => {
  // Guardrail: repo-level typecheck can run with TS configs that do *not* enable
  // `allowImportingTsExtensions`. `.ts` specifiers in source imports/exports would
  // then fail the build, so keep the public entrypoint free of `.ts` specifiers.
  const tsSpecifierRe =
    /(?:\bfrom\s+|\bimport\s*\(\s*|\bimport\s+)\s*['"]\.\.?\/[^'"\n]+?\.(?:ts|tsx)(?:[?#][^'"\n]*)?['"]/;
  const indexSrc = readFileSync(new URL("../packages/workbook-backend/src/index.ts", import.meta.url), "utf8");
  assert.ok(
    !tsSpecifierRe.test(stripComments(indexSrc)),
    "packages/workbook-backend/src/index.ts must not use .ts specifiers",
  );

  const mod = await import("@formula/workbook-backend");

  assert.equal(typeof mod.getSheetNameValidationErrorMessage, "function");

  // Basic happy-path.
  assert.equal(mod.getSheetNameValidationErrorMessage("Sheet1"), null);
  assert.equal(getMessageFromIndex("Sheet1"), null);

  // Runtime correctness smoke tests (these ensure the JS runtime implementation stays
  // in sync with the TS sources that provide types).
  assert.equal(mod.getSheetNameValidationErrorMessage(""), "sheet name cannot be blank");
  assert.equal(getMessageFromIndex(""), "sheet name cannot be blank");

  assert.equal(mod.getSheetNameValidationErrorMessage("'Budget"), "sheet name cannot begin or end with an apostrophe");
  assert.equal(getMessageFromIndex("'Budget"), "sheet name cannot begin or end with an apostrophe");

  assert.equal(mod.getSheetNameValidationErrorMessage("Bad:Name"), "sheet name contains invalid character `:`");
  assert.equal(getMessageFromIndex("Bad:Name"), "sheet name contains invalid character `:`");

  assert.equal(mod.getSheetNameValidationErrorMessage("budget", { existingNames: ["Budget"] }), "sheet name already exists");
  assert.equal(getMessageFromIndex("budget", { existingNames: ["Budget"] }), "sheet name already exists");

  // TS/JS parity: the repo keeps a `.ts` source version (for types) alongside a `.js`
  // runtime version (for Node ESM import specifiers). Keep their observable behavior in sync.
  assert.equal(TS_EXCEL_MAX_SHEET_NAME_LEN, JS_EXCEL_MAX_SHEET_NAME_LEN);
  assert.deepEqual([...TS_INVALID_SHEET_NAME_CHARACTERS], JS_INVALID_SHEET_NAME_CHARACTERS);

  // 31 character limit is measured in UTF-16 code units (JS `string.length`).
  // 🙂 is outside the BMP and takes two UTF-16 code units.
  const maxOk = `${"a".repeat(TS_EXCEL_MAX_SHEET_NAME_LEN - 2)}🙂`;
  assert.equal(maxOk.length, TS_EXCEL_MAX_SHEET_NAME_LEN);
  const tooLong = `${"a".repeat(TS_EXCEL_MAX_SHEET_NAME_LEN - 1)}🙂`;
  assert.equal(tooLong.length, TS_EXCEL_MAX_SHEET_NAME_LEN + 1);
  assert.equal(mod.getSheetNameValidationErrorMessage(maxOk), null);
  assert.equal(mod.getSheetNameValidationErrorMessage(tooLong), `sheet name cannot exceed ${TS_EXCEL_MAX_SHEET_NAME_LEN} characters`);

  const samples = [
    { name: "Sheet1", options: undefined },
    { name: "", options: undefined },
    { name: "'Budget", options: undefined },
    { name: "Bad:Name", options: undefined },
    { name: "budget", options: { existingNames: ["Budget"] } },
    { name: maxOk, options: undefined },
    { name: tooLong, options: undefined },
  ];
  for (const sample of samples) {
    assert.equal(getMessageFromTs(sample.name, sample.options), getMessageFromJs(sample.name, sample.options));
    assert.equal(mod.getSheetNameValidationErrorMessage(sample.name, sample.options), getMessageFromJs(sample.name, sample.options));
  }
});
