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

test("Typing =_xlfn.XLO suggests =_xlfn.XLOOKUP(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.XLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=_xlfn.XLOOKUP("),
    `Expected an _xlfn.XLOOKUP suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
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

test("Range suggestions do not auto-close parens when the function needs more args (VLOOKUP)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=VLOOKUP(A1, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=VLOOKUP(A1, A1:A10"),
    `Expected a VLOOKUP range suggestion without closing paren, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =COUNTIF(A suggests a range but does not auto-close parens", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=COUNTIF(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=COUNTIF(A1:A10"),
    `Expected a COUNTIF range suggestion without closing paren, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =MAX(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=MAX(A",
    cursorPosition: 6,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=MAX(A1:A10)"),
    `Expected a MAX range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =IRR(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=IRR(A",
    cursorPosition: 6,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=IRR(A1:A10)"),
    `Expected an IRR range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =TOD suggests TODAY() (zero-arg function inserts closing paren)", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=TOD",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=TODAY()"),
    `Expected a TODAY() suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Argument value suggestions use catalog arg_types (RANDBETWEEN suggests numbers)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANDBETWEEN(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=RANDBETWEEN(1"),
    `Expected a numeric argument suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
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
