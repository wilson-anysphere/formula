import test from "node:test";
import assert from "node:assert/strict";

import { parseSpreadsheetCellKey } from "../packages/versioning/src/yjs/sheetState.js";

test("parseSpreadsheetCellKey: marks strict canonical keys as isCanonical=true", () => {
  assert.deepEqual(parseSpreadsheetCellKey("Sheet1:0:0"), {
    sheetId: "Sheet1",
    row: 0,
    col: 0,
    isCanonical: true,
  });
});

test("parseSpreadsheetCellKey: supports an out param for reuse", () => {
  const out = { sheetId: "", row: 0, col: 0, isCanonical: false };
  const r1 = parseSpreadsheetCellKey("Sheet1:0:0", undefined, out);
  assert.equal(r1, out);
  assert.deepEqual(out, { sheetId: "Sheet1", row: 0, col: 0, isCanonical: true });

  const r2 = parseSpreadsheetCellKey("Sheet1:0,1", undefined, out);
  assert.equal(r2, out);
  assert.deepEqual(out, { sheetId: "Sheet1", row: 0, col: 1, isCanonical: false });
});

test("parseSpreadsheetCellKey: marks non-canonical numeric encodings as isCanonical=false", () => {
  assert.deepEqual(parseSpreadsheetCellKey("Sheet1:00:0"), { sheetId: "Sheet1", row: 0, col: 0, isCanonical: false });
  assert.deepEqual(parseSpreadsheetCellKey("Sheet1:1e0:2"), { sheetId: "Sheet1", row: 1, col: 2, isCanonical: false });
  assert.deepEqual(parseSpreadsheetCellKey("Sheet1:1:2 "), { sheetId: "Sheet1", row: 1, col: 2, isCanonical: false });
});

test("parseSpreadsheetCellKey: supports legacy and r{row}c{col} encodings", () => {
  assert.deepEqual(parseSpreadsheetCellKey("Sheet1:0,0"), { sheetId: "Sheet1", row: 0, col: 0, isCanonical: false });
  assert.deepEqual(parseSpreadsheetCellKey("r0c0"), { sheetId: "Sheet1", row: 0, col: 0, isCanonical: false });
  assert.deepEqual(parseSpreadsheetCellKey("r1c2", { defaultSheetId: "Other" }), {
    sheetId: "Other",
    row: 1,
    col: 2,
    isCanonical: false,
  });
});

test("parseSpreadsheetCellKey: default-substituted sheet ids are not canonical", () => {
  assert.deepEqual(parseSpreadsheetCellKey(":0:0", { defaultSheetId: "Other" }), {
    sheetId: "Other",
    row: 0,
    col: 0,
    isCanonical: false,
  });
});

test("parseSpreadsheetCellKey: rejects unsupported key shapes", () => {
  assert.equal(parseSpreadsheetCellKey("Sheet1:0:0:extra"), null);
  assert.equal(parseSpreadsheetCellKey("Sheet1:-1:0"), null);
});
