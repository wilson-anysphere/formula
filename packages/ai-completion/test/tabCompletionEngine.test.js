import assert from "node:assert/strict";
import test from "node:test";

import { TabCompletionEngine } from "../src/tabCompletionEngine.js";
import { parsePartialFormula } from "../src/formulaPartialParser.js";
import { FunctionRegistry } from "../src/functionRegistry.js";

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

test("Typing = suggests starter functions like SUM(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.deepEqual(
    suggestions.map((s) => s.text),
    ["=SUM(", "=AVERAGE(", "=IF(", "=XLOOKUP(", "=VLOOKUP("],
    `Expected stable starter ordering, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TabCompletionEngine supports async parsePartialFormula overrides", async () => {
  const engine = new TabCompletionEngine({
    // Simulate a worker/WASM-backed partial parser that is async.
    parsePartialFormula: async (input) => {
      if (input !== "=VLO") return { isFormula: false, inFunctionCall: false };
      return {
        isFormula: true,
        inFunctionCall: false,
        functionNamePrefix: { text: "VLO", start: 1, end: 4 },
      };
    },
  });

  const suggestions = await engine.getSuggestions({
    currentInput: "=VLO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP("),
    `Expected async parser to yield VLOOKUP suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TabCompletionEngine falls back when async parsePartialFormula throws", async () => {
  let calls = 0;
  const engine = new TabCompletionEngine({
    parsePartialFormula: async () => {
      calls += 1;
      throw new Error("parser unavailable");
    },
  });

  const suggestions = await engine.getSuggestions({
    currentInput: "=VLO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(calls, 1);
  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP("),
    `Expected fallback parser to yield VLOOKUP suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =<space> suggests starter functions (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "= ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.deepEqual(
    suggestions.map((s) => s.text),
    ["= SUM(", "= AVERAGE(", "= IF(", "= XLOOKUP(", "= VLOOKUP("],
    `Expected stable starter ordering preserving the space, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing = suggests an extended starter list when maxSuggestions is increased", async () => {
  const engine = new TabCompletionEngine({ maxSuggestions: 7 });

  const currentInput = "=";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.deepEqual(
    suggestions.map((s) => s.text),
    ["=SUM(", "=AVERAGE(", "=IF(", "=XLOOKUP(", "=VLOOKUP(", "=INDEX(", "=MATCH("],
    `Expected extended starter ordering, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Backend completion client is not called for empty formulas (just '=')", async () => {
  let calls = 0;
  const completionClient = {
    async completeTabCompletion() {
      calls += 1;
      return "SHOULD_NOT_BE_USED";
    },
  };

  const engine = new TabCompletionEngine({ completionClient, completionTimeoutMs: 200 });

  const currentInput = "=";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(calls, 0);
  assert.ok(
    suggestions.some((s) => s.text === "=SUM("),
    `Expected a SUM starter suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

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

test("Function name completion works after ';' inside an array constant", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "={1;VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "={1;VLOOKUP("),
    `Expected VLOOKUP completion after ';', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Function name completion works after '{' inside an array constant", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "={VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "={VLOOKUP("),
    `Expected VLOOKUP completion after '{', got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =_xlfn.TAK suggests =_xlfn.TAKE(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.TAK";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.TAKE("),
    `Expected an _xlfn.TAKE suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.DRO suggests =_xlfn.DROP(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.DRO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.DROP("),
    `Expected an _xlfn.DROP suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.EXPA suggests =_xlfn.EXPAND(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.EXPA";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.EXPAND("),
    `Expected an _xlfn.EXPAND suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.TEXTSPL suggests =_xlfn.TEXTSPLIT(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.TEXTSPL";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.TEXTSPLIT("),
    `Expected an _xlfn.TEXTSPLIT suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.CHOOSECO suggests =_xlfn.CHOOSECOLS(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.CHOOSECO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.CHOOSECOLS("),
    `Expected an _xlfn.CHOOSECOLS suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.CHOOSERO suggests =_xlfn.CHOOSEROWS(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.CHOOSERO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.CHOOSEROWS("),
    `Expected an _xlfn.CHOOSEROWS suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =SUM(A suggests a contiguous range below the current cell when the formula is above the data block", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 2; r <= 11; r++) {
    values[`A${r}`] = r; // A2..A11 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=SUM(A",
    cursorPosition: 6,
    // Pretend we're on A1 (0-based row 0), above the data in column A.
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A2:A11)"),
    `Expected a SUM range suggestion for data below, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM(A suggests the full contiguous block when the formula is inside the block (different column)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=SUM(A",
    cursorPosition: 6,
    // Pretend we're on B5 (0-based row 4), inside the A1..A10 block.
    cellRef: { row: 4, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1:A10)"),
    `Expected a SUM range suggestion for the full block, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions work for subsequent args when ';' is used as the argument separator", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUM(A1; A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1; A1:A10)"),
    `Expected a SUM range suggestion for the 2nd arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions work for an empty subsequent arg when ';' is used as the argument separator", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  // Trailing space is common after typing a separator in the formula bar.
  const currentInput = "=SUM(A1; ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on A11 (0-based row 10), below the data.
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1; A1:A10)"),
    `Expected a SUM range suggestion for the 2nd (empty) arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM( suggests a contiguous range above the current cell using the active column", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUM(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on A11 (0-based row 10), below the data in column A.
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1:A10)"),
    `Expected a SUM range suggestion from an empty arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Empty-arg range defaults have slightly lower confidence than explicit prefixes", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) values[`A${r}`] = r;

  const fromPrefix = await engine.getSuggestions({
    currentInput: "=SUM(A",
    cursorPosition: 6,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });
  const prefixSuggestion = fromPrefix.find((s) => s.text === "=SUM(A1:A10)");
  assert.ok(prefixSuggestion, "Expected explicit-prefix suggestion to exist");

  const fromEmpty = await engine.getSuggestions({
    currentInput: "=SUM(",
    cursorPosition: 5,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });
  const emptySuggestion = fromEmpty.find((s) => s.text === "=SUM(A1:A10)");
  assert.ok(emptySuggestion, "Expected empty-arg suggestion to exist");

  assert.ok(
    (emptySuggestion.confidence ?? 0) < (prefixSuggestion.confidence ?? 0),
    `Expected empty-arg confidence to be lower (empty=${emptySuggestion.confidence}, typed=${prefixSuggestion.confidence})`
  );
});

test("Typing =TAKE(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=TAKE(A",
    cursorPosition: 7,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=TAKE(A1:A10)"),
    `Expected a TAKE range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =DROP(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=DROP(A",
    cursorPosition: 7,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=DROP(A1:A10)"),
    `Expected a DROP range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.TAKE(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.TAKE(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.TAKE(A1:A10)"),
    `Expected an _xlfn.TAKE range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.DROP(A suggests a contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.DROP(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.DROP(A1:A10)"),
    `Expected an _xlfn.DROP range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (CHOOSECOLS)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=CHOOSECOLS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=CHOOSECOLS(A1:A10"),
    `Expected a CHOOSECOLS range suggestion without closing paren, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Typing =TEXTSPLIT(A suggests a contiguous range above the current cell but does not auto-close parens", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=TEXTSPLIT(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TEXTSPLIT(A1:A10"),
    `Expected a TEXTSPLIT range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.TEXTSPLIT(A suggests a contiguous range above the current cell but does not auto-close parens", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.TEXTSPLIT(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.TEXTSPLIT(A1:A10"),
    `Expected an _xlfn.TEXTSPLIT range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (_xlfn.CHOOSECOLS)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.CHOOSECOLS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.CHOOSECOLS(A1:A10"),
    `Expected an _xlfn.CHOOSECOLS range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (CHOOSEROWS)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=CHOOSEROWS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=CHOOSEROWS(A1:A10"),
    `Expected a CHOOSEROWS range suggestion without closing paren, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (_xlfn.CHOOSEROWS)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.CHOOSEROWS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.CHOOSEROWS(A1:A10"),
    `Expected an _xlfn.CHOOSEROWS range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (EXPAND)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=EXPAND(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some(s => s.text === "=EXPAND(A1:A10"),
    `Expected an EXPAND range suggestion without closing paren, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Range suggestions do not auto-close parens when the function needs more args (_xlfn.EXPAND)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.EXPAND(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.EXPAND(A1:A10"),
    `Expected an _xlfn.EXPAND range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM($A suggests an absolute-column contiguous range above the current cell", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const suggestions = await engine.getSuggestions({
    currentInput: "=SUM($A",
    cursorPosition: 7,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM($A1:$A10)"),
    `Expected an absolute-column SUM range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM(A1:A10 suggests auto-closing parens when the range is already complete", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) values[`A${r}`] = r;

  const currentInput = "=SUM(A1:A10";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1:A10)"),
    `Expected a pure paren-close suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Auto-closing parens is not suggested when the function needs more args (VLOOKUP)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) values[`A${r}`] = r;

  const currentInput = "=VLOOKUP(A1, A1:A10";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  // No range candidates (the range is already complete) and VLOOKUP still requires
  // additional args, so don't suggest an auto-close.
  assert.equal(suggestions.length, 0);
});

test("Range suggestions do not auto-close parens when the function needs more args (VLOOKUP)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const emptyArgInput = "=VLOOKUP(A1, ";
  const emptyArgSuggestions = await engine.getSuggestions({
    currentInput: emptyArgInput,
    cursorPosition: emptyArgInput.length,
    // Pretend we're on A11 (0-based row 10), below the data in column A.
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    emptyArgSuggestions.some((s) => s.text === "=VLOOKUP(A1, A1:A10"),
    `Expected a VLOOKUP range suggestion from an empty arg without closing paren, got: ${emptyArgSuggestions
      .map((s) => s.text)
      .join(", ")}`
  );

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

test("Range suggestions work for ';' separators even when the formula contains decimal commas", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  // In semicolon locales, `,` is often used as the decimal separator.
  const currentInput = "=VLOOKUP(1,2; A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP(1,2; A1:A10"),
    `Expected a VLOOKUP range suggestion for the 2nd arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =VLOOKUP(A1, A suggests a 2D table range when adjacent columns form a table", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  // Header row.
  values["A1"] = "Key";
  values["B1"] = "Value1";
  values["C1"] = "Value2";
  values["D1"] = "Value3";
  // Data rows 2..10.
  for (let r = 2; r <= 10; r++) {
    values[`A${r}`] = `K${r}`;
    values[`B${r}`] = r * 10;
    values[`C${r}`] = r * 100;
    values[`D${r}`] = r * 1000;
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
    suggestions.some((s) => s.text === "=VLOOKUP(A1, A1:D10"),
    `Expected a VLOOKUP table range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("VLOOKUP table-range bias prefers a 2D range when the formula is above the table block", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  // A2:D11 (rows 2..11) contain a dense numeric table.
  for (let r = 2; r <= 11; r++) {
    values[`A${r}`] = r;
    values[`B${r}`] = r * 10;
    values[`C${r}`] = r * 100;
    values[`D${r}`] = r * 1000;
  }

  const currentInput = "=VLOOKUP(A1, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on A1 (0-based row 0), above the table.
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.equal(suggestions[0]?.text, "=VLOOKUP(A1, A2:D11");
});

test("Typing =FILTER(A suggests a 2D table range when adjacent columns form a table", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  // Header row.
  values["A1"] = "Key";
  values["B1"] = "Value1";
  values["C1"] = "Value2";
  values["D1"] = "Value3";
  // Data rows 2..10.
  for (let r = 2; r <= 10; r++) {
    values[`A${r}`] = `K${r}`;
    values[`B${r}`] = r * 10;
    values[`C${r}`] = r * 100;
    values[`D${r}`] = r * 1000;
  }

  const currentInput = "=FILTER(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FILTER(A1:D10"),
    `Expected a FILTER table range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =SUMIFS(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUMIFS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUMIFS(A1:A10"),
    `Expected a SUMIFS range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.SUMIFS(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=_xlfn.SUMIFS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.SUMIFS(A1:A10"),
    `Expected an _xlfn.SUMIFS range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SUMIFS repeating criteria_range suggestions do not auto-close parens (criteria2 still required)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = '=SUMIFS(A1:A10, A1:A10, ">5", A';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=SUMIFS(A1:A10, A1:A10, ">5", A1:A10'),
    `Expected a SUMIFS criteria_range2 suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =FILTER(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=FILTER(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FILTER(A1:A10"),
    `Expected a FILTER range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =BYROW(A suggests a range but does not auto-close parens (needs lambda)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=BYROW(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=BYROW(A1:A10"),
    `Expected a BYROW range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =REDUCE(A suggests a range but does not auto-close parens (needs lambda)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=REDUCE(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=REDUCE(A1:A10"),
    `Expected a REDUCE range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SCAN(A suggests a range but does not auto-close parens (needs lambda)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SCAN(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SCAN(A1:A10"),
    `Expected a SCAN range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =PERCENTILE(A suggests a range but does not auto-close parens (needs k)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=PERCENTILE(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=PERCENTILE(A1:A10"),
    `Expected a PERCENTILE range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =DSUM(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=DSUM(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=DSUM(A1:A10"),
    `Expected a DSUM range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =DSUM(A suggests a 2D table range when adjacent columns form a table", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  // Header row.
  values["A1"] = "Key";
  values["B1"] = "Value1";
  values["C1"] = "Value2";
  values["D1"] = "Value3";
  // Data rows 2..10.
  for (let r = 2; r <= 10; r++) {
    values[`A${r}`] = `K${r}`;
    values[`B${r}`] = r * 10;
    values[`C${r}`] = r * 100;
    values[`D${r}`] = r * 1000;
  }

  const currentInput = "=DSUM(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=DSUM(A1:D10"),
    `Expected a DSUM table range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =FORECAST.ETS(1, A suggests a range but does not auto-close parens (needs timeline)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=FORECAST.ETS(1, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS(1, A1:A10"),
    `Expected a FORECAST.ETS range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =WORKDAY(1, 5, A suggests a range and auto-closes (optional holidays arg)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=WORKDAY(1, 5, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=WORKDAY(1, 5, A1:A10)"),
    `Expected a WORKDAY holidays range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =ROWS(A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=ROWS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ROWS(A1:A10)"),
    `Expected a ROWS range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =TEXTJOIN(\",\",TRUE,A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = '=TEXTJOIN(",", TRUE, A';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTJOIN(",", TRUE, A1:A10)'),
    `Expected a TEXTJOIN range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUBTOTAL(9, A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUBTOTAL(9, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUBTOTAL(9, A1:A10)"),
    `Expected a SUBTOTAL range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =LARGE(A suggests a range but does not auto-close parens (needs k)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=LARGE(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=LARGE(A1:A10"),
    `Expected a LARGE range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =FORECAST.LINEAR(10, A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=FORECAST.LINEAR(10, A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.LINEAR(10, A1:A10"),
    `Expected a FORECAST.LINEAR range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =STDEV.S(A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=STDEV.S(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=STDEV.S(A1:A10)"),
    `Expected a STDEV.S range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =OFFSET(A suggests a range but does not auto-close parens (needs rows/cols)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=OFFSET(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=OFFSET(A1:A10"),
    `Expected an OFFSET range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =MMULT(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=MMULT(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=MMULT(A1:A10"),
    `Expected an MMULT range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =T.TEST(A suggests a range but does not auto-close parens (needs more args)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=T.TEST(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10"),
    `Expected a T.TEST range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =MODE.SNGL(A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=MODE.SNGL(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=MODE.SNGL(A1:A10)"),
    `Expected a MODE.SNGL range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =TRIMMEAN(A suggests a range but does not auto-close parens (needs percent)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=TRIMMEAN(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TRIMMEAN(A1:A10"),
    `Expected a TRIMMEAN range suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =HSTACK(A suggests a range and auto-closes (min args satisfied)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=HSTACK(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=HSTACK(A1:A10)"),
    `Expected an HSTACK range suggestion with closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =RAN suggests RAND() (another zero-arg function)", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=RAN",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=RAND()"),
    `Expected a RAND() suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
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

test("MATCH match_type suggests 0, 1, -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=MATCH(A1, A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1, A1:A10, 0"),
    `Expected MATCH to suggest match_type=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1, A1:A10, 1"),
    `Expected MATCH to suggest match_type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1, A1:A10, -1"),
    `Expected MATCH to suggest match_type=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("MATCH match_type suggestions work with ';' argument separators", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=MATCH(A1; A1:A10; ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1; A1:A10; 0"),
    `Expected MATCH to suggest match_type=0 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1; A1:A10; 1"),
    `Expected MATCH to suggest match_type=1 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=MATCH(A1; A1:A10; -1"),
    `Expected MATCH to suggest match_type=-1 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("XLOOKUP match_mode suggests 0, -1, 1, 2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=XLOOKUP(A1, A1:A10, B1:B10, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , 0"),
    `Expected XLOOKUP to suggest match_mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , -1"),
    `Expected XLOOKUP to suggest match_mode=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , 1"),
    `Expected XLOOKUP to suggest match_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , 2"),
    `Expected XLOOKUP to suggest match_mode=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("XLOOKUP match_mode suggestions work with ';' argument separators", async () => {
  const engine = new TabCompletionEngine();

  // Leave if_not_found blank so we're completing match_mode (5th arg).
  const currentInput = "=XLOOKUP(A1; A1:A10; B1:B10; ; ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1; A1:A10; B1:B10; ; 0"),
    `Expected XLOOKUP to suggest match_mode=0 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1; A1:A10; B1:B10; ; -1"),
    `Expected XLOOKUP to suggest match_mode=-1 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1; A1:A10; B1:B10; ; 1"),
    `Expected XLOOKUP to suggest match_mode=1 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1; A1:A10; B1:B10; ; 2"),
    `Expected XLOOKUP to suggest match_mode=2 with semicolon separators, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("XLOOKUP search_mode suggests 1, -1, 2, -2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=XLOOKUP(A1, A1:A10, B1:B10, , , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , , 1"),
    `Expected XLOOKUP to suggest search_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , , -1"),
    `Expected XLOOKUP to suggest search_mode=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , , 2"),
    `Expected XLOOKUP to suggest search_mode=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP(A1, A1:A10, B1:B10, , , -2"),
    `Expected XLOOKUP to suggest search_mode=-2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("VLOOKUP range_lookup suggests TRUE/FALSE with higher confidence", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:B10, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  const exact = suggestions.find((s) => s.text === "=VLOOKUP(A1, A1:B10, 2, FALSE");
  assert.ok(exact, `Expected VLOOKUP to suggest FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok(
    (exact?.confidence ?? 0) > 0.5,
    `Expected VLOOKUP/FALSE to have elevated confidence, got: ${exact?.confidence}`
  );

  const approx = suggestions.find((s) => s.text === "=VLOOKUP(A1, A1:B10, 2, TRUE");
  assert.ok(approx, `Expected VLOOKUP to suggest TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok(
    (approx?.confidence ?? 0) > 0.5,
    `Expected VLOOKUP/TRUE to have elevated confidence, got: ${approx?.confidence}`
  );
});

test("XMATCH match_mode suggests 0, -1, 1, 2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=XMATCH(A1, A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, 0"),
    `Expected XMATCH to suggest match_mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, -1"),
    `Expected XMATCH to suggest match_mode=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, 1"),
    `Expected XMATCH to suggest match_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, 2"),
    `Expected XMATCH to suggest match_mode=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("XMATCH search_mode suggests 1, -1, 2, -2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=XMATCH(A1, A1:A10, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, , 1"),
    `Expected XMATCH to suggest search_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, , -1"),
    `Expected XMATCH to suggest search_mode=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, , 2"),
    `Expected XMATCH to suggest search_mode=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH(A1, A1:A10, , -2"),
    `Expected XMATCH to suggest search_mode=-2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SORT sort_order suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SORT(A1:A10, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SORT(A1:A10, 1, 1"),
    `Expected SORT to suggest sort_order=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=SORT(A1:A10, 1, -1"),
    `Expected SORT to suggest sort_order=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TAKE rows suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TAKE(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TAKE(A1:A10, 1"),
    `Expected TAKE to suggest rows=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=TAKE(A1:A10, -1"),
    `Expected TAKE to suggest rows=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TEXTSPLIT ignore_empty suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTSPLIT("a,,b", ",", , ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("a,,b", ",", , TRUE'),
    `Expected TEXTSPLIT to suggest ignore_empty=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("a,,b", ",", , FALSE'),
    `Expected TEXTSPLIT to suggest ignore_empty=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TEXTSPLIT match_mode suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTSPLIT("aXb", "x", , FALSE, ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("aXb", "x", , FALSE, 0'),
    `Expected TEXTSPLIT to suggest match_mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("aXb", "x", , FALSE, 1'),
    `Expected TEXTSPLIT to suggest match_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TEXTJOIN ignore_empty suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTJOIN(",", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTJOIN(",", TRUE'),
    `Expected TEXTJOIN to suggest ignore_empty=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTJOIN(",", FALSE'),
    `Expected TEXTJOIN to suggest ignore_empty=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("UNIQUE by_col suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=UNIQUE(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=UNIQUE(A1:A10, FALSE"),
    `Expected UNIQUE to suggest by_col=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=UNIQUE(A1:A10, TRUE"),
    `Expected UNIQUE to suggest by_col=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("UNIQUE exactly_once suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=UNIQUE(A1:A10, FALSE, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=UNIQUE(A1:A10, FALSE, TRUE"),
    `Expected UNIQUE to suggest exactly_once=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=UNIQUE(A1:A10, FALSE, FALSE"),
    `Expected UNIQUE to suggest exactly_once=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SUBTOTAL function_num suggests 9 and 109", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUBTOTAL(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUBTOTAL(9"),
    `Expected SUBTOTAL to suggest function_num=9, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=SUBTOTAL(109"),
    `Expected SUBTOTAL to suggest function_num=109, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("AGGREGATE function_num suggests 9", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=AGGREGATE(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=AGGREGATE(9"),
    `Expected AGGREGATE to suggest function_num=9, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("AGGREGATE options suggests common values (0, 4, 6, 7)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=AGGREGATE(9, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=AGGREGATE(9, 0"),
    `Expected AGGREGATE to suggest options=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=AGGREGATE(9, 4"),
    `Expected AGGREGATE to suggest options=4, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=AGGREGATE(9, 6"),
    `Expected AGGREGATE to suggest options=6, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=AGGREGATE(9, 7"),
    `Expected AGGREGATE to suggest options=7, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("T.TEST tails suggests 1 and 2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=T.TEST(A1:A10, B1:B10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10, B1:B10, 1"),
    `Expected T.TEST to suggest tails=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10, B1:B10, 2"),
    `Expected T.TEST to suggest tails=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("T.TEST type suggests 1, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=T.TEST(A1:A10, B1:B10, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10, B1:B10, 2, 1"),
    `Expected T.TEST to suggest type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10, B1:B10, 2, 2"),
    `Expected T.TEST to suggest type=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=T.TEST(A1:A10, B1:B10, 2, 3"),
    `Expected T.TEST to suggest type=3, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("RANK.EQ order suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANK.EQ(10, A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=RANK.EQ(10, A1:A10, 0"),
    `Expected RANK.EQ to suggest order=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=RANK.EQ(10, A1:A10, 1"),
    `Expected RANK.EQ to suggest order=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("WEEKDAY return_type suggests 1, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=WEEKDAY(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=WEEKDAY(A1, 1"),
    `Expected WEEKDAY to suggest return_type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=WEEKDAY(A1, 2"),
    `Expected WEEKDAY to suggest return_type=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=WEEKDAY(A1, 3"),
    `Expected WEEKDAY to suggest return_type=3, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("WEEKNUM return_type suggests 1, 2, 21", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=WEEKNUM(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=WEEKNUM(A1, 1"),
    `Expected WEEKNUM to suggest return_type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=WEEKNUM(A1, 2"),
    `Expected WEEKNUM to suggest return_type=2, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=WEEKNUM(A1, 21"),
    `Expected WEEKNUM to suggest return_type=21, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("DAYS360 method suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=DAYS360(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  const us = suggestions.find((s) => s.text === "=DAYS360(A1, B1, FALSE");
  assert.ok(us, `Expected DAYS360 to suggest FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok((us?.confidence ?? 0) > 0.5, `Expected DAYS360/FALSE to have elevated confidence, got: ${us?.confidence}`);

  const eu = suggestions.find((s) => s.text === "=DAYS360(A1, B1, TRUE");
  assert.ok(eu, `Expected DAYS360 to suggest TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok((eu?.confidence ?? 0) > 0.5, `Expected DAYS360/TRUE to have elevated confidence, got: ${eu?.confidence}`);
});

test("YEARFRAC basis suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=YEARFRAC(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const basis of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=YEARFRAC(A1, B1, ${basis}`),
      `Expected YEARFRAC to suggest basis=${basis}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("FORECAST.ETS seasonality suggests 0, 1, 12, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS(A1, B1:B10, C1:C10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["0", "1", "12", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=FORECAST.ETS(A1, B1:B10, C1:C10, ${v}`),
      `Expected FORECAST.ETS to suggest seasonality=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("FORECAST.ETS data_completion suggests 1 and 0", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS(A1, B1:B10, C1:C10, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS(A1, B1:B10, C1:C10, , 1"),
    `Expected FORECAST.ETS to suggest data_completion=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS(A1, B1:B10, C1:C10, , 0"),
    `Expected FORECAST.ETS to suggest data_completion=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("FORECAST.ETS aggregation suggests common values (1, 7)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS(A1, B1:B10, C1:C10, , , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS(A1, B1:B10, C1:C10, , , 1"),
    `Expected FORECAST.ETS to suggest aggregation=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS(A1, B1:B10, C1:C10, , , 7"),
    `Expected FORECAST.ETS to suggest aggregation=7, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("FORECAST.ETS.CONFINT confidence_level suggests 0.95, 0.9, 0.99", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS.CONFINT(A1, B1:B10, C1:C10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["0.95", "0.9", "0.99"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=FORECAST.ETS.CONFINT(A1, B1:B10, C1:C10, ${v}`),
      `Expected FORECAST.ETS.CONFINT to suggest confidence_level=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("FORECAST.ETS.SEASONALITY data_completion suggests 1 and 0", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS.SEASONALITY(B1:B10, C1:C10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS.SEASONALITY(B1:B10, C1:C10, 1"),
    `Expected FORECAST.ETS.SEASONALITY to suggest data_completion=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS.SEASONALITY(B1:B10, C1:C10, 0"),
    `Expected FORECAST.ETS.SEASONALITY to suggest data_completion=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("LINEST const suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=LINEST(A1:A10, B1:B10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  const calc = suggestions.find((s) => s.text === "=LINEST(A1:A10, B1:B10, TRUE");
  assert.ok(calc, `Expected LINEST to suggest TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok((calc?.confidence ?? 0) > 0.5, `Expected LINEST/TRUE to have elevated confidence, got: ${calc?.confidence}`);

  const force0 = suggestions.find((s) => s.text === "=LINEST(A1:A10, B1:B10, FALSE");
  assert.ok(force0, `Expected LINEST to suggest FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok(
    (force0?.confidence ?? 0) > 0.5,
    `Expected LINEST/FALSE to have elevated confidence, got: ${force0?.confidence}`
  );
});

test("LINEST stats suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=LINEST(A1:A10, B1:B10, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=LINEST(A1:A10, B1:B10, , TRUE"),
    `Expected LINEST to suggest stats=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=LINEST(A1:A10, B1:B10, , FALSE"),
    `Expected LINEST to suggest stats=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("ADDRESS abs_num suggests 1, 4, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ADDRESS(1, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "4", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=ADDRESS(1, 2, ${v}`),
      `Expected ADDRESS to suggest abs_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("ADDRESS a1 suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ADDRESS(1, 2, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ADDRESS(1, 2, 1, TRUE"),
    `Expected ADDRESS to suggest a1=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=ADDRESS(1, 2, 1, FALSE"),
    `Expected ADDRESS to suggest a1=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("INDIRECT a1 suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=INDIRECT("A1", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=INDIRECT("A1", TRUE'),
    `Expected INDIRECT to suggest a1=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=INDIRECT("A1", FALSE'),
    `Expected INDIRECT to suggest a1=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("RANDARRAY whole_number suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANDARRAY(2, 3, 0, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=RANDARRAY(2, 3, 0, 1, TRUE"),
    `Expected RANDARRAY to suggest whole_number=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=RANDARRAY(2, 3, 0, 1, FALSE"),
    `Expected RANDARRAY to suggest whole_number=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CEILING.MATH mode suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CEILING.MATH(-5.5, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CEILING.MATH(-5.5, 2, 0"),
    `Expected CEILING.MATH to suggest mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CEILING.MATH(-5.5, 2, 1"),
    `Expected CEILING.MATH to suggest mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("FLOOR.MATH mode suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FLOOR.MATH(-5.5, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FLOOR.MATH(-5.5, 2, 0"),
    `Expected FLOOR.MATH to suggest mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=FLOOR.MATH(-5.5, 2, 1"),
    `Expected FLOOR.MATH to suggest mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("QUARTILE.INC quart suggests 1, 2, 3, 0, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=QUARTILE.INC(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const q of ["1", "2", "3", "0", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=QUARTILE.INC(A1:A10, ${q}`),
      `Expected QUARTILE.INC to suggest quart=${q}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("NORM.DIST cumulative suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=NORM.DIST(0, 0, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=NORM.DIST(0, 0, 1, TRUE"),
    `Expected NORM.DIST to suggest TRUE (cumulative), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=NORM.DIST(0, 0, 1, FALSE"),
    `Expected NORM.DIST to suggest FALSE (probability), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("NORM.S.DIST cumulative suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=NORM.S.DIST(0, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=NORM.S.DIST(0, TRUE"),
    `Expected NORM.S.DIST to suggest TRUE (cumulative), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=NORM.S.DIST(0, FALSE"),
    `Expected NORM.S.DIST to suggest FALSE (probability), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TabCompletionEngine caches suggestions by context key", async () => {
  let callCount = 0;
  const completionClient = {
    async completeTabCompletion() {
      callCount++;
      return "+1";
    },
  };

  const engine = new TabCompletionEngine({ completionClient, completionTimeoutMs: 200 });

  const ctx = {
    currentInput: "=1+",
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  };

  const s1 = await engine.getSuggestions(ctx);
  const s2 = await engine.getSuggestions(ctx);

  assert.deepEqual(s1, s2);
  assert.equal(callCount, 1, "Expected completion client to be called once due to caching");
});

test("parsePartialFormula ignores commas inside structured refs and array constants", () => {
  const registry = new FunctionRegistry();

  const structured = "=SUM(Table1[[#All],[Amount]]";
  const structuredParsed = parsePartialFormula(structured, structured.length, registry);
  assert.equal(structuredParsed.argIndex, 0);
  assert.equal(structuredParsed.currentArg?.text, "Table1[[#All],[Amount]]");

  const arrayConst = "=SUM({1,2},A";
  const arrayParsed = parsePartialFormula(arrayConst, arrayConst.length, registry);
  assert.equal(arrayParsed.argIndex, 1);
  assert.equal(arrayParsed.currentArg?.text, "A");
});

test("parsePartialFormula ignores apostrophes and parentheses inside structured refs", () => {
  const registry = new FunctionRegistry();

  const parenInColumnName = "=SUM(Table1[Amount (USD]";
  const parenParsed = parsePartialFormula(parenInColumnName, parenInColumnName.length, registry);
  assert.equal(parenParsed.inFunctionCall, true);
  assert.equal(parenParsed.functionName, "SUM");
  assert.equal(parenParsed.argIndex, 0);
  assert.equal(parenParsed.currentArg?.text, "Table1[Amount (USD]");

  const apostropheInColumnName = "=SUM(Table1[Bob's]";
  const apostropheParsed = parsePartialFormula(apostropheInColumnName, apostropheInColumnName.length, registry);
  assert.equal(apostropheParsed.inFunctionCall, true);
  assert.equal(apostropheParsed.functionName, "SUM");
  assert.equal(apostropheParsed.argIndex, 0);
  assert.equal(apostropheParsed.currentArg?.text, "Table1[Bob's]");
});

test("TabCompletionEngine cache busts when schemaProvider cache key changes", async () => {
  let callCount = 0;
  let schemaKey = "v1";

  const completionClient = {
    async completeTabCompletion() {
      callCount++;
      return "+1";
    },
  };

  const engine = new TabCompletionEngine({
    completionClient,
    completionTimeoutMs: 200,
    schemaProvider: {
      getCacheKey: () => schemaKey,
    },
  });

  const ctx = {
    currentInput: "=1+",
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  };

  await engine.getSuggestions(ctx);
  schemaKey = "v2";
  await engine.getSuggestions(ctx);

  assert.equal(callCount, 2, "Expected completion client to be called again when schema key changes");
});

test("TabCompletionEngine cache busts when surroundingCells cache key changes", async () => {
  let callCount = 0;
  let cellsKey = "cells:v1";

  const completionClient = {
    async completeTabCompletion() {
      callCount++;
      return "+1";
    },
  };

  const engine = new TabCompletionEngine({
    completionClient,
    completionTimeoutMs: 200,
  });

  const ctx = {
    currentInput: "=1+",
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: {
      ...createMockCellContext({}),
      getCacheKey: () => cellsKey,
    },
  };

  await engine.getSuggestions(ctx);
  cellsKey = "cells:v2";
  await engine.getSuggestions(ctx);

  assert.equal(callCount, 2, "Expected completion client to be called again when surrounding key changes");
});

test("Named ranges are suggested in range arguments (=SUM(Sal → SalesData)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [{ name: "SalesData", range: "Sheet1!A1:A10" }],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sal";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some(s => s.text === "=SUM(SalesData)"),
    `Expected a named-range suggestion, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("Named ranges preserve the typed prefix case (lowercase)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [{ name: "SalesData", range: "Sheet1!A1:A10" }],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(sal";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(salesData)"),
    `Expected a named-range suggestion that preserves prefix case, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Completion client request is structured and completion inserts at the cursor", async () => {
  /** @type {any} */
  let seenReq = null;

  const completionClient = {
    async completeTabCompletion(req) {
      seenReq = req;
      return "2";
    },
  };

  const engine = new TabCompletionEngine({ completionClient, completionTimeoutMs: 200 });

  const currentInput = "=1+";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(seenReq?.input, currentInput);
  assert.equal(seenReq?.cursorPosition, currentInput.length);
  assert.equal(seenReq?.cellA1, "A1");
  assert.equal(typeof seenReq?.signal?.aborted, "boolean");
  assert.ok(
    suggestions.some(s => s.text === "=1+2"),
    `Expected the completion to be inserted, got: ${suggestions.map(s => s.text).join(", ")}`
  );
});

test("previewEvaluator is called and preview metadata is attached", async () => {
  let calls = 0;
  /** @type {any} */
  let last = null;
  const previewEvaluator = (params) => {
    calls += 1;
    last = params;
    return "42";
  };

  const engine = new TabCompletionEngine();
  const currentInput = "=TOD";
  const suggestions = await engine.getSuggestions(
    {
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    },
    { previewEvaluator }
  );

  const today = suggestions.find((s) => s.text === "=TODAY()");
  assert.ok(today, `Expected TODAY() suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.equal(today.preview, "42");
  assert.ok(calls >= 1);
  assert.equal(last?.suggestion?.text, today.text);
});

test("Structured references are suggested from table schemas", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["Amount"] }],
    },
  });

  const currentInput = "=SUM(Tab";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Table1[Amount])"),
    `Expected a structured reference suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Structured references preserve the typed prefix case (lowercase)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["Amount"] }],
    },
  });

  const currentInput = "=SUM(tab";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(table1[Amount])"),
    `Expected a structured reference suggestion that preserves prefix case, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Structured references are not suggested when the user types '[' before the table name is complete", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["Amount"] }],
    },
  });

  // Completing this would require inserting missing characters *before* the '[',
  // which isn't representable as a pure insertion at the caret.
  const currentInput = "=SUM(Tab[";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
});

test("Sheet-name prefixes are suggested as SheetName! inside range args (=SUM(she → sheet2!) without auto-closing parens", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2", "My Sheet", "A1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(she";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => ["=SUM(sheet1!", "=SUM(sheet2!"].includes(s.text)),
    `Expected a sheet prefix suggestion ending with '!', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Quoted sheet-name prefixes are suggested as 'Sheet Name'! inside range args (=SUM('my → 'my Sheet'!) without auto-closing parens", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2", "My Sheet", "A1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('my";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('my Sheet'!"),
    `Expected a quoted sheet prefix suggestion without closing paren, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet prefixes are not suggested when the user hasn't started quotes for a sheet that needs them (=SUM(My Sheet)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["My Sheet"],
      getTables: () => [],
    },
  });

  // Do not attempt to "fix" missing quotes here (would not be a pure insertion).
  const currentInput = "=SUM(My Sheet";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
});

test("Sheet-qualified ranges are suggested when typing Sheet2!A", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A1:A10)"),
    `Expected a sheet-qualified range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges are suggested when typing Sheet2!A above the data block", async () => {
  const values = {};
  for (let r = 2; r <= 11; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 1 (0-based 0), above the data.
    cellRef: { row: 0, col: 0 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A2:A11)"),
    `Expected a sheet-qualified range suggestion for data below, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified partial range prefixes do not produce invalid insertions (Sheet2!A: avoids '::')", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A:";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    !suggestions.some((s) => s.text.includes("::")),
    `Expected no invalid '::' suggestions, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A:A)"),
    `Expected a whole-column completion for the partial 'A:' prefix, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified partial range prefixes do not emit non-insertions (Sheet2!A1: avoids trailing ':')", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A1:";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A1:A10)"),
    `Expected a completed A1:A10 range for the 'A1:' prefix, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    !suggestions.some((s) => s.text === "=SUM(Sheet2!A1:)"),
    `Expected no suggestions that do not extend the typed prefix, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified whole-column ranges still allow auto-closing parens (=SUM(Sheet2!A:A → ...))", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A:A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A:A)"),
    `Expected an auto-closed paren suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified complete ranges still allow auto-closing parens (=SUM(Sheet2!A1:A10 → ...))", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2", "My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!A1:A10";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A1:A10)"),
    `Expected an auto-closed paren suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Quoted sheet-qualified complete ranges still allow auto-closing parens (=SUM('My Sheet'!A1:A10 → ...))", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2", "My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('My Sheet'!A1:A10";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('My Sheet'!A1:A10)"),
    `Expected an auto-closed paren suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges work when the quoted sheet name contains a comma", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Jan,2024!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Jan,2024"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('Jan,2024'!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('Jan,2024'!A1:A10)"),
    `Expected a sheet-qualified range suggestion for a comma-containing sheet, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges preserve absolute column prefixes (Sheet2!$A → Sheet2!$A1:$A10)", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Sheet2!$A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!$A1:$A10)"),
    `Expected an absolute-column sheet-qualified range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges preserve the typed prefix case for sheet names", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(sheet2!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(sheet2!A1:A10)"),
    `Expected a sheet-qualified range suggestion that preserves prefix case, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges quote sheet names with spaces", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`My Sheet!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('My Sheet'!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('My Sheet'!A1:A10)"),
    `Expected a quoted sheet-qualified range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges escape apostrophes in sheet names", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Bob's Sheet!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Bob's Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('Bob''s Sheet'!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('Bob''s Sheet'!A1:A10)"),
    `Expected an escaped sheet-qualified range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified range suggestions do not attempt to add missing quotes (not a pure insertion)", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`My Sheet!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "My Sheet"],
      getTables: () => [],
    },
  });

  // We intentionally don't suggest quote-fixing completions here because adding
  // a leading quote would modify text *before* the cursor (the formula bar only
  // shows/apply "pure insertion" completions).
  const currentInput = "=SUM(My Sheet!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.equal(suggestions.length, 0);
});

test("Sheet-qualified range suggestions require quotes for sheet names that start with a digit", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`2024!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "2024"],
      getTables: () => [],
    },
  });

  const unquoted = "=SUM(2024!A";
  const unquotedSuggestions = await engine.getSuggestions({
    currentInput: unquoted,
    cursorPosition: unquoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });
  assert.equal(unquotedSuggestions.length, 0);

  const quoted = "=SUM('2024'!A";
  const quotedSuggestions = await engine.getSuggestions({
    currentInput: quoted,
    cursorPosition: quoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    quotedSuggestions.some((s) => s.text === "=SUM('2024'!A1:A10)"),
    `Expected a quoted numeric sheet range suggestion, got: ${quotedSuggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified range suggestions require quotes for sheet names that look like A1 refs (A1)", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`A1!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "A1"],
      getTables: () => [],
    },
  });

  const unquoted = "=SUM(A1!A";
  const unquotedSuggestions = await engine.getSuggestions({
    currentInput: unquoted,
    cursorPosition: unquoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });
  assert.equal(unquotedSuggestions.length, 0);

  const quoted = "=SUM('A1'!A";
  const quotedSuggestions = await engine.getSuggestions({
    currentInput: quoted,
    cursorPosition: quoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    quotedSuggestions.some((s) => s.text === "=SUM('A1'!A1:A10)"),
    `Expected a quoted A1 sheet range suggestion, got: ${quotedSuggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified range suggestions require quotes for reserved sheet names (TRUE)", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`TRUE!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "TRUE"],
      getTables: () => [],
    },
  });

  const unquoted = "=SUM(TRUE!A";
  const unquotedSuggestions = await engine.getSuggestions({
    currentInput: unquoted,
    cursorPosition: unquoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });
  assert.equal(unquotedSuggestions.length, 0);

  const quoted = "=SUM('TRUE'!A";
  const quotedSuggestions = await engine.getSuggestions({
    currentInput: quoted,
    cursorPosition: quoted.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    quotedSuggestions.some((s) => s.text === "=SUM('TRUE'!A1:A10)"),
    `Expected a quoted TRUE sheet range suggestion, got: ${quotedSuggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified ranges are not suggested when the sheet name prefix is incomplete (can't be a pure insertion)", async () => {
  const values = {};
  for (let r = 1; r <= 10; r++) values[`Sheet2!A${r}`] = r;

  const cellContext = {
    getCellValue(row, col, sheetName) {
      const sheet = sheetName ?? "Sheet1";
      const a1 = `${sheet}!${columnIndexToLetter(col)}${row + 1}`;
      return values[a1] ?? null;
    },
  };

  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(She!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.equal(suggestions.length, 0);
});

test("Sheet names are suggested as identifiers when typing =Sheet", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=Sheet";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Sheet1!" || s.text === "=Sheet2!"),
    `Expected a sheet-name identifier suggestion ending with '!', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet name suggestions preserve the typed prefix case (lowercase)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2"],
      getTables: () => [],
    },
  });

  const currentInput = "=shee";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text.startsWith("=shee") && s.text.endsWith("!")),
    `Expected a sheet-name suggestion that preserves prefix case, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet names that require quotes are not suggested as identifiers (=My Sheet is ignored)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=My";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(
    suggestions.filter((s) => s.text.endsWith("!")).length,
    0,
    `Expected no sheet-name suggestions ending with '!', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("getSuggestions never throws when cellRef is malformed", async () => {
  const engine = new TabCompletionEngine();
  const invalidRefs = [null, { row: "x" }, "not-a1"];

  for (const cellRef of invalidRefs) {
    const formulaInput = "=SUM(A";
    const formulaSuggestions = await engine.getSuggestions({
      currentInput: formulaInput,
      cursorPosition: formulaInput.length,
      cellRef,
      surroundingCells: createMockCellContext({}),
    });
    assert.ok(Array.isArray(formulaSuggestions), `Expected array for cellRef=${String(cellRef)}`);

    const valueInput = "x";
    const valueSuggestions = await engine.getSuggestions({
      currentInput: valueInput,
      cursorPosition: valueInput.length,
      cellRef,
      surroundingCells: createMockCellContext({ A2: "xray" }),
    });
    assert.ok(Array.isArray(valueSuggestions), `Expected array for cellRef=${String(cellRef)}`);
  }
});

test("Completion client request falls back to A1 when cellRef is invalid", async () => {
  /** @type {any} */
  let seenReq = null;
  const completionClient = {
    async completeTabCompletion(req) {
      seenReq = req;
      return "2";
    },
  };

  const engine = new TabCompletionEngine({ completionClient, completionTimeoutMs: 200 });

  const currentInput = "=1+";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: "not-a1",
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(seenReq?.input, currentInput);
  assert.equal(seenReq?.cursorPosition, currentInput.length);
  assert.equal(seenReq?.cellA1, "A1");
  assert.equal(typeof seenReq?.signal?.aborted, "boolean");
  assert.ok(Array.isArray(suggestions));
});

test("parsePartialFormula errors do not crash getSuggestions (falls back to pattern suggestions)", async () => {
  const engine = new TabCompletionEngine({
    parsePartialFormula() {
      throw new Error("boom");
    },
  });

  const currentInput = "ap";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 5, col: 0 },
    surroundingCells: createMockCellContext({ A5: "apple", A4: "apricot" }),
  });

  assert.ok(Array.isArray(suggestions));
  assert.ok(
    suggestions.some((s) => s.text === "apple" || s.text === "apricot"),
    `Expected pattern suggestions, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("getSuggestions is crash-proof for non-string currentInput", async () => {
  const engine = new TabCompletionEngine();

  const nonStringInputSuggestions = await engine.getSuggestions({
    // @ts-ignore - intentionally invalid
    currentInput: 123,
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });
  assert.equal(nonStringInputSuggestions.length, 0);
});
