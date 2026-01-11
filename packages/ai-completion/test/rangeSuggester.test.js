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

test("suggestRanges trims non-numeric header rows when the range is mostly numeric", () => {
  const ctx = createColumnAContext([
    [0, "Header"],
    [1, 10],
    [2, 20],
    [3, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 4, col: 0 }, // row 5, below data
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A2:A4");
});

test("suggestRanges preserves absolute column/row prefixes in A1 output", () => {
  const ctx = createColumnAContext([
    [0, 10],
    [1, 20],
    [2, 30],
  ]);

  const absCol = suggestRanges({
    currentArgText: "$A",
    cellRef: { row: 3, col: 0 }, // row 4, below data
    surroundingCells: ctx,
  });

  assert.equal(absCol[0].range, "$A1:$A3");
  assert.equal(absCol[1].range, "$A:$A");

  const absRow = suggestRanges({
    currentArgText: "A$1",
    cellRef: { row: 0, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(absRow[0].range, "A$1:A$3");

  const absBoth = suggestRanges({
    currentArgText: "$A$1",
    cellRef: { row: 0, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(absBoth[0].range, "$A$1:$A$3");
});
