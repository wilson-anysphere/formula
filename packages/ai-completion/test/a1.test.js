import assert from "node:assert/strict";
import test from "node:test";

import { normalizeCellRef, parseA1 } from "../src/a1.js";

test("normalizeCellRef accepts sheet-qualified A1 refs (Sheet1!A1)", () => {
  assert.deepEqual(normalizeCellRef("Sheet1!A1"), { row: 0, col: 0 });
});

test("normalizeCellRef accepts sheet-qualified A1 refs with quoted sheet names ('My Sheet'!$B$2)", () => {
  assert.deepEqual(normalizeCellRef("'My Sheet'!$B$2"), { row: 1, col: 1 });
});

test("normalizeCellRef accepts sheet-qualified A1 refs with escaped quotes ('Bob''s Sheet'!C3)", () => {
  assert.deepEqual(normalizeCellRef("'Bob''s Sheet'!C3"), { row: 2, col: 2 });
});

test("parseA1 returns null when the sheet-qualified cell portion is invalid", () => {
  assert.equal(parseA1("Sheet1!NotACell"), null);
  assert.equal(parseA1("'My Sheet'!A0"), null);
});

