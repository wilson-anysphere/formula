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

function createGridContext(cells) {
  const values = new Map();
  for (const [rowIndex, colIndex, value] of cells) {
    values.set(`${rowIndex},${colIndex}`, value);
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

test("suggestRanges suggests a 2D table range when adjacent columns form a rectangular block", () => {
  /** @type {Array<[number, number, any]>} */
  const cells = [];
  // Header row (row 1 in A1 notation).
  for (let c = 0; c < 4; c++) cells.push([0, c, `H${c + 1}`]);
  // Data rows 2..10.
  for (let r = 1; r < 10; r++) {
    for (let c = 0; c < 4; c++) {
      cells.push([r, c, r * 100 + c]);
    }
  }

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 10, col: 0 }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "A1:D10"),
    `Expected suggestions to contain A1:D10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges does not suggest a 2D table range when only one column is populated", () => {
  /** @type {Array<[number, number, any]>} */
  const cells = [];
  for (let r = 0; r < 10; r++) cells.push([r, 0, r === 0 ? "Header" : r]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 10, col: 0 }, // row 11, below the data
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    !suggestions.some((s) => /A\d+:[B-Z]/.test(s.range)),
    `Expected no multi-column A1 range suggestions, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});
