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

test("suggestRanges falls back to scanning down when no data exists above the current cell", () => {
  const ctx = createColumnAContext([
    [1, 10], // A2
    [2, 20], // A3
    [3, 30], // A4
    [4, 40], // A5
    [5, 50], // A6
    [6, 60], // A7
    [7, 70], // A8
    [8, 80], // A9
    [9, 90], // A10
    [10, 100], // A11
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 0, col: 0 }, // A1, above the data block
    surroundingCells: ctx,
  });

  assert.ok(suggestions.some((s) => s.range === "A2:A11"));
});

test("suggestRanges downward fallback includes same-row data when the active cell is in a different column", () => {
  const ctx = createColumnAContext([
    [1, 10], // A2
    [2, 20], // A3
    [3, 30], // A4
    [4, 40], // A5
    [5, 50], // A6
    [6, 60], // A7
    [7, 70], // A8
    [8, 80], // A9
    [9, 90], // A10
    [10, 100], // A11
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 1, col: 1 }, // B2, beside the data block start row
    surroundingCells: ctx,
  });

  assert.ok(suggestions.some((s) => s.range === "A2:A11"));
});

test("suggestRanges returns the full contiguous block when the active cell is inside the block (different column)", () => {
  const ctx = createColumnAContext(Array.from({ length: 10 }, (_, i) => [i, i + 1])); // A1..A10

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 4, col: 1 }, // B5, inside the A1..A10 block
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A1:A10");
});

test("suggestRanges in-block expansion respects maxScanRows (does not double the cap)", () => {
  const ctx = createColumnAContext(Array.from({ length: 1000 }, (_, i) => [i, i + 1])); // A1..A1000

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 500, col: 1 }, // B501, inside the large block
    surroundingCells: ctx,
    maxScanRows: 200,
  });

  // With a 200-row cap, we should not extend downward beyond the scanned window.
  assert.equal(suggestions[0].range, "A302:A501");
});

test("suggestRanges trims non-numeric header rows when scanning downwards", () => {
  const ctx = createColumnAContext([
    [1, "Header"], // A2
    [2, 10], // A3
    [3, 20], // A4
    [4, 30], // A5
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 0, col: 0 }, // A1, above the header+data block
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A3:A5");
});

test("suggestRanges does not return an invalid contiguous-above range when the scan cap is hit", () => {
  // Data exists, but is too far above the active row to be detected under a small maxScanRows cap.
  const ctx = createColumnAContext([[0, 1]]); // A1

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 10, col: 0 }, // A11, far below the data
    surroundingCells: ctx,
    maxScanRows: 3,
  });

  // Should fall back to the safe "entire column" suggestion only.
  assert.equal(suggestions.length, 1);
  assert.equal(suggestions[0].range, "A:A");
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

test("suggestRanges completes partial A1 range syntax (A1: -> A1:A10) and preserves absolute markers ($A$1:)", () => {
  const ctx = createColumnAContext(Array.from({ length: 10 }, (_, i) => [i, i + 1]));

  const suggestions = suggestRanges({
    currentArgText: "A1:",
    cellRef: { row: 20, col: 0 }, // arbitrary; explicit start cell bounds the scan
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A1:A10");

  const absSuggestions = suggestRanges({
    currentArgText: "$A$1:",
    cellRef: { row: 20, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(absSuggestions[0].range, "$A$1:$A$10");
});

test("suggestRanges preserves absolute markers from the typed end column token (A1:$A -> A1:$A3)", () => {
  const ctx = createColumnAContext([
    [0, 10],
    [1, 20],
    [2, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A1:$A",
    cellRef: { row: 10, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A1:$A3");
  assert.equal(suggestions[1].range, "A:$A");
});

test("suggestRanges supports partial end-column prefixes for multi-letter columns (AB1:A -> AB1:AB3)", () => {
  // Column AB is 0-based index 27 (A=0, B=1, ..., Z=25, AA=26, AB=27).
  const AB_COL = 27;
  const ctx = createGridContext([
    [0, AB_COL, 10],
    [1, AB_COL, 20],
    [2, AB_COL, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "AB1:A",
    cellRef: { row: 0, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "AB1:AB3");
});

test("suggestRanges completes partial column range syntax (A: -> A:A)", () => {
  const ctx = createColumnAContext([
    [0, 10],
    [1, 20],
    [2, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A:",
    cellRef: { row: 3, col: 0 }, // row 4, below data
    surroundingCells: ctx,
  });

  assert.equal(suggestions.length, 1);
  assert.equal(suggestions[0].range, "A:A");
});

test("suggestRanges is conservative for multi-column prefixes (A1:B -> no suggestions)", () => {
  const ctx = createColumnAContext([
    [0, 10],
    [1, 20],
    [2, 30],
  ]);

  const suggestions = suggestRanges({
    currentArgText: "A1:B",
    cellRef: { row: 3, col: 0 },
    surroundingCells: ctx,
  });

  assert.deepEqual(suggestions, []);
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

test("suggestRanges respects maxScanCols when expanding a 2D table range", () => {
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
    maxScanCols: 2,
  });

  assert.ok(
    suggestions.some((s) => s.range === "A1:B10"),
    `Expected suggestions to contain A1:B10 when maxScanCols=2, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
  assert.ok(
    !suggestions.some((s) => s.range === "A1:D10"),
    `Expected suggestions to not contain A1:D10 when maxScanCols=2, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges suggests a 2D table range when an explicit start cell is provided (A1)", () => {
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
    currentArgText: "A1",
    cellRef: { row: 10, col: 0 }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "A1:D10"),
    `Expected suggestions to contain A1:D10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges preserves end-column $ prefix in 2D table suggestions (A:$A -> A1:$D10)", () => {
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
    currentArgText: "A:$A",
    cellRef: { row: 10, col: 0 }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "A1:$D10"),
    `Expected suggestions to contain A1:$D10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges preserves end-column token casing for 2D table suggestions (AB1:a -> AB1:ad10)", () => {
  /** @type {Array<[number, number, any]>} */
  const cells = [];
  // Start at column AB (0-based 27) so the table expands into AC/AD.
  const startCol = 27; // AB
  // Header row across AB:AD.
  for (let c = 0; c < 3; c++) cells.push([0, startCol + c, `H${c + 1}`]);
  // Data rows 2..10.
  for (let r = 1; r < 10; r++) {
    for (let c = 0; c < 3; c++) {
      cells.push([r, startCol + c, r * 100 + c]);
    }
  }

  const suggestions = suggestRanges({
    // User started typing AB1:AB10, but only entered a lowercase prefix for the end column.
    currentArgText: "AB1:a",
    cellRef: { row: 10, col: startCol }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "AB1:ad10"),
    `Expected suggestions to contain AB1:ad10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges suggests a 2D table range across the Z->AA column boundary (Y1:AB10)", () => {
  /** @type {Array<[number, number, any]>} */
  const cells = [];
  const startCol = 24; // Y
  // Header row across Y:AB.
  for (let c = 0; c < 4; c++) cells.push([0, startCol + c, `H${c + 1}`]);
  // Data rows 2..10.
  for (let r = 1; r < 10; r++) {
    for (let c = 0; c < 4; c++) {
      cells.push([r, startCol + c, r * 100 + c]);
    }
  }

  const suggestions = suggestRanges({
    currentArgText: "Y",
    cellRef: { row: 10, col: startCol }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "Y1:AB10"),
    `Expected suggestions to contain Y1:AB10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges preserves absolute row/col prefixes for 2D table suggestions", () => {
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
    currentArgText: "$A$1",
    cellRef: { row: 10, col: 0 }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "$A$1:$D$10"),
    `Expected suggestions to contain $A$1:$D$10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges preserves lowercase column prefixes for 2D table suggestions", () => {
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
    currentArgText: "a",
    cellRef: { row: 10, col: 0 }, // row 11, below the table
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "a1:d10"),
    `Expected suggestions to contain a1:d10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges stops table expansion when encountering a gap (entirely empty column)", () => {
  /** @type {Array<[number, number, any]>} */
  const cells = [];
  // Header row across A:D (but C will be empty across all rows -> gap).
  cells.push([0, 0, "Key"]);
  cells.push([0, 1, "Value"]);
  // Intentionally omit any values for column C (index 2).
  cells.push([0, 3, "Ignored"]);

  // Data rows 2..10.
  for (let r = 1; r < 10; r++) {
    cells.push([r, 0, `K${r}`]);
    cells.push([r, 1, r]);
    // Column C gap.
    cells.push([r, 3, r * 1000]); // Should not be pulled in due to the gap at C.
  }

  const suggestions = suggestRanges({
    currentArgText: "A",
    cellRef: { row: 10, col: 0 }, // row 11, below the data
    surroundingCells: createGridContext(cells),
  });

  assert.ok(
    suggestions.some((s) => s.range === "A1:B10"),
    `Expected suggestions to contain A1:B10, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
  assert.ok(
    !suggestions.some((s) => s.range === "A1:D10"),
    `Expected suggestions to not contain A1:D10 due to gap column, got: ${suggestions.map((s) => s.range).join(", ")}`
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
    !suggestions.some((s) => /A\\d+:[B-Z]/.test(s.range)),
    `Expected no multi-column A1 range suggestions, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges returns suggestions for an empty argument using the active cell column", () => {
  const ctx = createGridContext([
    [0, 1, 10], // B1
    [1, 1, 20], // B2
    [2, 1, 30], // B3
  ]);

  const suggestions = suggestRanges({
    currentArgText: "",
    cellRef: { row: 3, col: 1 }, // B4
    surroundingCells: ctx,
  });

  assert.ok(
    suggestions.some((s) => s.range === "B:B"),
    `Expected suggestions to contain B:B, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.range === "B1:B3"),
    `Expected suggestions to contain B1:B3, got: ${suggestions.map((s) => s.range).join(", ")}`
  );
});

test("suggestRanges completes partial range syntax with an explicit end column token (A1:A -> A1:A10)", () => {
  const ctx = createColumnAContext(Array.from({ length: 10 }, (_, i) => [i, i + 1]));

  const suggestions = suggestRanges({
    currentArgText: "A1:A",
    cellRef: { row: 20, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0].range, "A1:A10");
});

test("suggestRanges returns no suggestions for columns beyond Excel max (XFD)", () => {
  const suggestions = suggestRanges({
    currentArgText: "ZZZ",
    cellRef: { row: 0, col: 0 },
    surroundingCells: createGridContext([]),
  });

  assert.deepEqual(suggestions, []);
});
