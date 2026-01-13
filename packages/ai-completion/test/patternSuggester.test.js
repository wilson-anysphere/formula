import assert from "node:assert/strict";
import test from "node:test";

import { suggestPatternValues } from "../src/patternSuggester.js";

function createContext(cells) {
  const values = new Map();
  for (const [row, col, value] of cells) {
    values.set(`${row},${col}`, value);
  }
  return {
    getCellValue(row, col) {
      return values.get(`${row},${col}`);
    },
  };
}

test("suggestPatternValues suggests repeated string matches in the current row", () => {
  const ctx = createContext([
    [5, 3, "Apple"],
    [5, 4, "Apple"],
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "Ap",
    cursorPosition: 2,
    cellRef: { row: 5, col: 5 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0]?.text, "Apple");
});

test("suggestPatternValues suggests repeated string matches in the current column", () => {
  const ctx = createContext([
    [3, 2, "Apple"],
    [4, 2, "Apple"],
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "Ap",
    cursorPosition: 2,
    cellRef: { row: 5, col: 2 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0]?.text, "Apple");
});

test("suggestPatternValues ranks closer matches above farther matches", () => {
  const ctx = createContext([
    [5, 4, "Bazooka"], // distance 1
    [5, 15, "Bar"], // distance 10
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "B",
    cursorPosition: 1,
    cellRef: { row: 5, col: 5 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0]?.text, "Bazooka");
  assert.equal(suggestions[1]?.text, "Bar");
});

test("suggestPatternValues returns no suggestions for formula inputs", () => {
  const ctx = createContext([
    [0, 0, "Apple"],
    [0, 1, "Apple"],
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "=Ap",
    cursorPosition: 3,
    cellRef: { row: 0, col: 2 },
    surroundingCells: ctx,
  });

  assert.deepEqual(suggestions, []);
});

test("suggestPatternValues suggests the next number in a simple column sequence", () => {
  const ctx = createContext([
    [0, 0, 1],
    [1, 0, 2],
    [2, 0, 3],
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "4",
    cursorPosition: 1,
    cellRef: { row: 3, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0]?.text, "4");
});

test("suggestPatternValues suggests the next number for an empty input (pure insertion)", () => {
  const ctx = createContext([
    [0, 0, 1],
    [1, 0, 2],
    [2, 0, 3],
  ]);

  const suggestions = suggestPatternValues({
    currentInput: "",
    cursorPosition: 0,
    cellRef: { row: 3, col: 0 },
    surroundingCells: ctx,
  });

  assert.equal(suggestions[0]?.text, "4");
});
