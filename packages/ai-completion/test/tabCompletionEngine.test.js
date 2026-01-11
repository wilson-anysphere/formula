import assert from "node:assert/strict";
import test from "node:test";

import { TabCompletionEngine } from "../src/tabCompletionEngine.js";

function createMockCellContext(valuesByA1) {
  // valuesByA1: { "A1": 1, ... }
  /** @type {Map<string, any>} */
  const map = new Map(Object.entries(valuesByA1));

  return {
    getCellValue(row, col) {
      const a1 = `${columnIndexToLetter(col)}${row + 1}`;
      return map.get(a1);
    },
  };
}

function columnIndexToLetter(index) {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let n = index;
  let out = "";
  while (n >= 0) {
    out = letters[n % 26] + out;
    n = Math.floor(n / 26) - 1;
  }
  return out;
}

test("Typing =VLO suggests VLOOKUP(", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=VLO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=VLOOKUP("),
    `Expected a VLOOKUP suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =XLO suggests XLOOKUP(", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=XLO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=XLOOKUP("),
    `Expected an XLOOKUP suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =SUM(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=SUM(A",
    cursorPosition: 6,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=SUM(A1:A10)"),
    `Expected a SUM range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("TabCompletionEngine caches suggestions by context key", async () => {
  let callCount = 0;
  const localModel = {
    async complete() {
      callCount++;
      return "+1";
    },
  };

  const engine = new TabCompletionEngine({ localModel, localModelTimeoutMs: 200 });

  const ctx = {
    currentInput: "=1+",
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  };

  const s1 = await engine.getSuggestions(ctx);
  const s2 = await engine.getSuggestions(ctx);

  assert.deepEqual(s1, s2);
  assert.equal(callCount, 1, "Expected local model to be called once due to caching");
});
