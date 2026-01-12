import test from "node:test";
import assert from "node:assert/strict";

import { MockWorkbook } from "@formula/python-runtime/test-utils";

test("MockWorkbook validates sheet names (blank, invalid chars, length, apostrophes)", () => {
  const workbook = new MockWorkbook();

  assert.throws(() => workbook.create_sheet({ name: "" }), /sheet name cannot be blank/i);
  assert.throws(() => workbook.create_sheet({ name: "   " }), /sheet name cannot be blank/i);
  assert.throws(() => workbook.create_sheet({ name: "Bad:Name" }), /sheet name contains invalid character `:`/i);
  assert.throws(() => workbook.create_sheet({ name: "'Budget" }), /sheet name cannot begin or end with an apostrophe/i);
  assert.throws(() => workbook.create_sheet({ name: "Budget'" }), /sheet name cannot begin or end with an apostrophe/i);

  const tooLong = "A".repeat(32);
  assert.throws(() => workbook.create_sheet({ name: tooLong }), /sheet name cannot exceed 31 characters/i);
});

test("MockWorkbook enforces case-insensitive uniqueness for sheet names", () => {
  const workbook = new MockWorkbook();

  assert.throws(() => workbook.create_sheet({ name: "sheet1" }), /sheet name already exists/i);
});

