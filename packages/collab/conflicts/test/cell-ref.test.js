import test from "node:test";
import assert from "node:assert/strict";

import { cellRefFromKey } from "../src/cell-ref.js";

test("cellRefFromKey supports r{row}c{col} convenience keys", () => {
  assert.deepEqual(cellRefFromKey("r0c2"), { sheetId: "Sheet1", row: 0, col: 2 });
});

test("cellRefFromKey supports canonical and legacy sheet keys", () => {
  assert.deepEqual(cellRefFromKey("Sheet1:0:1"), { sheetId: "Sheet1", row: 0, col: 1 });
  assert.deepEqual(cellRefFromKey("Sheet1:2,3"), { sheetId: "Sheet1", row: 2, col: 3 });
});

test("cellRefFromKey throws on invalid keys", () => {
  assert.throws(() => cellRefFromKey("not-a-cell"), /Invalid cell key/);
});

