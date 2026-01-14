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

test("normalizeFormula handles workbook prefixes containing '[' (non-nesting) and continues normalization after the prefix", () => {
  // Workbook names may contain literal `[` characters. Unlike structured references, workbook
  // prefixes are not nested, so `[A1[Name.xlsx]Sheet1!A1` is valid and the bracket span ends at
  // the first (non-escaped) `]`.
  //
  // Regression: the fallback normalizer must still remove whitespace *after* the prefix (e.g. the
  // space after the comma), rather than treating the `[` inside the workbook name as introducing
  // nesting and bailing out early.
  const withSpace = normalizeFormula("=SUM([A1[Name.xlsx]Sheet1!A1, 1)");
  const withoutSpace = normalizeFormula("=SUM([A1[Name.xlsx]Sheet1!A1,1)");

  assert.equal(withSpace, "=SUM([A1[NAME.XLSX]SHEET1!A1,1)");
  assert.equal(withSpace, withoutSpace);
});

test("normalizeFormula handles workbook-scoped external defined names (no '!') and continues normalization after the prefix", () => {
  // Workbook-scoped external defined names do not use `!`, but we still need to detect the end of
  // the workbook prefix (which is non-nesting) so we can keep normalizing the rest of the formula.
  const withSpace = normalizeFormula("=SUM([A1[Name.xlsx]MyName, 1)");
  const withoutSpace = normalizeFormula("=SUM([A1[Name.xlsx]MyName,1)");

  assert.equal(withSpace, "=SUM([A1[NAME.XLSX]MYNAME,1)");
  assert.equal(withSpace, withoutSpace);
});
