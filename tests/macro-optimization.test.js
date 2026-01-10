import assert from "node:assert/strict";
import test from "node:test";

import { optimizeMacroActions } from "../apps/desktop/src/macro-recorder/index.js";

test("optimizeMacroActions merges dense rectangles into setRangeValues", () => {
  const actions = [
    { type: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
    { type: "setCellValue", sheetName: "Sheet1", address: "B1", value: 2 },
    { type: "setCellValue", sheetName: "Sheet1", address: "A2", value: 3 },
    { type: "setCellValue", sheetName: "Sheet1", address: "B2", value: 4 },
  ];

  const optimized = optimizeMacroActions(actions);
  assert.equal(optimized.length, 1);
  assert.deepEqual(optimized[0], {
    type: "setRangeValues",
    sheetName: "Sheet1",
    address: "A1:B2",
    values: [
      [1, 2],
      [3, 4],
    ],
  });
});

test("optimizeMacroActions collapses consecutive selections", () => {
  const actions = [
    { type: "setSelection", sheetName: "Sheet1", address: "A1" },
    { type: "setSelection", sheetName: "Sheet1", address: "B2" },
    { type: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
  ];

  const optimized = optimizeMacroActions(actions);
  assert.deepEqual(optimized[0], { type: "setSelection", sheetName: "Sheet1", address: "B2" });
  assert.deepEqual(optimized[1], { type: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 });
});

