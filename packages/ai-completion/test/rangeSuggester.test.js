import assert from "node:assert/strict";
import test from "node:test";

import { suggestRanges } from "../src/rangeSuggester.js";

function createColumnAContext(rowsWithValues) {
  const values = new Map();
  for (const [rowIndex, value] of rowsWithValues) {
    values.set(`${rowIndex},0`, value);
  }
  return {
    getCellValue(row, col) {
      return values.get(`${row},${col}`);
    },
  };
}

test("suggestRanges returns contiguous range above current cell for a column prefix", () => {
  const ctx = createColumnAContext([
    [0, 10],
    [1, 20],
    [2, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 3, col: 0 }, // row 4, below data
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A1:A3");
});
