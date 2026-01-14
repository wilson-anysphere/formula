import assert from "node:assert/strict";
import test from "node:test";

import { normalizeFormula } from "../src/formula/normalize.js";

test("normalizeFormula preserves whitespace inside structured reference brackets", () => {
  const withSpace = normalizeFormula("=Table1[Total Amount]");
  const withoutSpace = normalizeFormula("=Table1[TotalAmount]");

  assert.equal(withSpace, "=TABLE1[TOTAL AMOUNT]");
  assert.equal(withoutSpace, "=TABLE1[TOTALAMOUNT]");
  assert.notEqual(withSpace, withoutSpace);
});

test("normalizeFormula preserves whitespace inside external workbook prefixes (including ]] escapes)", () => {
  // Workbook names can contain literal `]` escaped as `]]` and also contain spaces. These should
  // remain significant when doing best-effort normalization for semantic diffs.
  const withSpace = normalizeFormula("=[Book]] Name.xlsx]Sheet1!A1");
  const withoutSpace = normalizeFormula("=[Book]]Name.xlsx]Sheet1!A1");

  assert.equal(withSpace, "=[BOOK]] NAME.XLSX]SHEET1!A1");
  assert.equal(withoutSpace, "=[BOOK]]NAME.XLSX]SHEET1!A1");
  assert.notEqual(withSpace, withoutSpace);
});

