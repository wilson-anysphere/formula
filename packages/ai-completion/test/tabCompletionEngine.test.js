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

test("TabCompletionEngine supports custom starter functions", async () => {
  const engine = new TabCompletionEngine({ starterFunctions: ["FOO(", "BAR("], maxSuggestions: 2 });

  const currentInput = "=";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.deepEqual(
    suggestions.map((s) => s.text),
    ["=FOO(", "=BAR("],
    `Expected custom starter ordering, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =HLO suggests HLOOKUP( and a modern XLOOKUP( alternative", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=HLO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=HLOOKUP("),
    `Expected a HLOOKUP( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP("),
    `Expected an XLOOKUP( alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =LOO suggests LOOKUP( and a modern XLOOKUP( alternative", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=LOO",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=LOOKUP("),
    `Expected a LOOKUP( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XLOOKUP("),
    `Expected an XLOOKUP( alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =loo suggests lookup( and a modern xlookup( alternative", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=loo",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=lookup("),
    `Expected a lookup( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=xlookup("),
    `Expected an xlookup( alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =MAT suggests MATCH( and a modern XMATCH( alternative", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=MAT",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=MATCH("),
    `Expected a MATCH( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=XMATCH("),
    `Expected an XMATCH( alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =mat suggests match( and a modern xmatch( alternative", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=mat",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=match("),
    `Expected a match( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=xmatch("),
    `Expected an xmatch( alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =CONCATE suggests CONCATENATE( and a modern CONCAT( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CONCATE";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CONCATENATE("),
    `Expected a CONCATENATE( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CONCAT("),
    `Expected a CONCAT( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =STDEVP suggests STDEVP( and a modern STDEV.P( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=STDEVP";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=STDEVP("),
    `Expected a STDEVP( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=STDEV.P("),
    `Expected a STDEV.P( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =stdevp suggests stdevp( and a modern stdev.p( alternative (lowercase)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=stdevp";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=stdevp("),
    `Expected a stdevp( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=stdev.p("),
    `Expected a stdev.p( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =VARP suggests VARP( and a modern VAR.P( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VARP";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VARP("),
    `Expected a VARP( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=VAR.P("),
    `Expected a VAR.P( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =NORMD suggests NORMDIST( and a modern NORM.DIST( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=NORMD";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=NORMDIST("),
    `Expected a NORMDIST( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=NORM.DIST("),
    `Expected a NORM.DIST( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =CHID suggests CHIDIST( and a modern CHISQ.DIST.RT( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHID";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CHIDIST("),
    `Expected a CHIDIST( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CHISQ.DIST.RT("),
    `Expected a CHISQ.DIST.RT( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =FINV suggests FINV( and a modern F.INV.RT( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FINV";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FINV("),
    `Expected a FINV( suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=F.INV.RT("),
    `Expected a F.INV.RT( modern alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Legacy distribution aliases suggest modern dotted function alternatives", async () => {
  const engine = new TabCompletionEngine();

  const cases = [
    { currentInput: "=CRITB", legacy: "=CRITBINOM(", modern: "=BINOM.INV(" },
    { currentInput: "=GAMMAD", legacy: "=GAMMADIST(", modern: "=GAMMA.DIST(" },
    { currentInput: "=EXPOND", legacy: "=EXPONDIST(", modern: "=EXPON.DIST(" },
    { currentInput: "=HYPGEOMD", legacy: "=HYPGEOMDIST(", modern: "=HYPGEOM.DIST(" },
    { currentInput: "=NEGBINOMD", legacy: "=NEGBINOMDIST(", modern: "=NEGBINOM.DIST(" },
    { currentInput: "=BETAD", legacy: "=BETADIST(", modern: "=BETA.DIST(" },
    { currentInput: "=BETAI", legacy: "=BETAINV(", modern: "=BETA.INV(" },
    { currentInput: "=LOGINV", legacy: "=LOGINV(", modern: "=LOGNORM.INV(" },
    { currentInput: "=TDI", legacy: "=TDIST(", modern: "=T.DIST.2T(" },
    { currentInput: "=TINV", legacy: "=TINV(", modern: "=T.INV.2T(" },
    { currentInput: "=ISOW", legacy: "=ISOWEEKNUM(", modern: "=ISO.WEEKNUM(" },
    { currentInput: "=FTEST", legacy: "=FTEST(", modern: "=F.TEST(" },
    { currentInput: "=ZTEST", legacy: "=ZTEST(", modern: "=Z.TEST(" },
    { currentInput: "=TTEST", legacy: "=TTEST(", modern: "=T.TEST(" },
  ];

  for (const { currentInput, legacy, modern } of cases) {
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    assert.ok(
      suggestions.some((s) => s.text === legacy),
      `Expected legacy function completion for ${currentInput} -> ${legacy}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
    assert.ok(
      suggestions.some((s) => s.text === modern),
      `Expected modern alternative for ${currentInput} -> ${modern}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("Typing =Vlo suggests Vlookup( (title-style casing)", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=Vlo",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Vlookup("),
    `Expected a Vlookup suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=Xlookup("),
    `Expected a title-case Xlookup modern alternative, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =Forecast.Et suggests Forecast.Ets( (segment title-style casing)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=Forecast.Et";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Forecast.Ets("),
    `Expected a Forecast.Ets suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =Foo. suggests Foo.Äbc( (Unicode segment title-style casing)", async () => {
  const functionRegistry = new FunctionRegistry([{ name: "FOO.ÄBC", args: [] }]);
  const engine = new TabCompletionEngine({ functionRegistry });

  const currentInput = "=Foo.";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Foo.Äbc("),
    `Expected a Foo.Äbc suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =Zä suggests Zählenwenn( (Unicode title-style casing)", async () => {
  const functionRegistry = new FunctionRegistry([{ name: "ZÄHLENWENN", args: [] }]);
  const engine = new TabCompletionEngine({ functionRegistry });

  const currentInput = "=Zä";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Zählenwenn("),
    `Expected a Zählenwenn suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("FunctionSpec.completionBoost biases function-name completion ranking", async () => {
  const functionRegistry = new FunctionRegistry([
    { name: "SUMIF", args: [] },
    // Same prefix length/overall length as SUMIF; without a boost, lexicographic tie-breaking would prefer SUMIF.
    { name: "SUMME", args: [], completionBoost: 0.05 },
  ]);
  const engine = new TabCompletionEngine({ functionRegistry, maxSuggestions: 2 });

  const suggestions = await engine.getSuggestions({
    currentInput: "=SU",
    cursorPosition: 3,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions[0]?.text, "=SUMME(");
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

test("Typing =vlo suggests vlookup(", async () => {
  const engine = new TabCompletionEngine();

  const suggestions = await engine.getSuggestions({
    currentInput: "=vlo",
    cursorPosition: 4,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=vlookup("),
    `Expected a vlookup suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=xlookup("),
    `Expected a lowercase xlookup modern alternative, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =_xlfn.VLO suggests =_xlfn.VLOOKUP( and a modern _xlfn.XLOOKUP( alternative", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.VLOOKUP("),
    `Expected an _xlfn.VLOOKUP suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.XLOOKUP("),
    `Expected an _xlfn.XLOOKUP alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.vlo suggests =_xlfn.vlookup( and a modern _xlfn.xlookup( alternative (lowercase)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.vlo";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.vlookup("),
    `Expected an _xlfn.vlookup suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.xlookup("),
    `Expected an _xlfn.xlookup alternative suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.Xlo suggests =_xlfn.Xlookup( (title-style casing)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.Xlo";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.Xlookup("),
    `Expected an _xlfn.Xlookup suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =_xlfn.Tak suggests =_xlfn.Take( (title-style casing)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.Tak";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.Take("),
    `Expected an _xlfn.Take suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =_xlfn.tak suggests =_xlfn.take(", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=_xlfn.tak";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=_xlfn.take("),
    `Expected an _xlfn.take suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Function name completion works after '@' (implicit intersection operator)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=@VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=@VLOOKUP("),
    `Expected VLOOKUP suggestion after '@', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Function name completion works after '&' (concatenation operator)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=A1&VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=A1&VLOOKUP("),
    `Expected VLOOKUP suggestion after '&', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Function name completion works after '>' (comparison operator)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=A1>VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=A1>VLOOKUP("),
    `Expected VLOOKUP suggestion after '>', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =LOG1 suggests LOG10( (function name looks like A1 cell ref)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=LOG1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=LOG10("),
    `Expected LOG10 suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Function name completion works inside range args (=SUM(OFFS → =SUM(OFFSET()", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(OFFS";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(OFFSET("),
    `Expected OFFSET( completion inside range arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Function name completion works inside range args after ':' (=SUM(A1:OFFS → =SUM(A1:OFFSET()", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(A1:OFFS";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1:OFFSET("),
    `Expected OFFSET( completion after ':', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range-colon function completion is conservative for 1-3 letter tokens (A1:OFF)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(A1:OFF";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.length === 0 || !suggestions.some((s) => s.text.includes("OFFSET(")),
    `Did not expect OFFSET( completion for short A1:OFF token, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range-colon function completion is conservative for A1-like tokens (A1:LOG1)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(A1:LOG1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.length === 0 || !suggestions.some((s) => s.text.includes("LOG10(")),
    `Did not expect LOG10( completion for A1-like A1:LOG1 token, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions work inside grouping parens (=SUM((A → =SUM((A1:A10)))", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUM((A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM((A1:A10))"),
    `Expected a grouped SUM range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions still work when grouping parens include whitespace (=SUM(( ␠ → =SUM(( ␠A1:A10)))", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUM(( ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on A11 (0-based row 10), below the data in column A.
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(( A1:A10))"),
    `Expected a grouped SUM range suggestion preserving whitespace, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions work with a unary '-' prefix (=SUM(-A → =SUM(-A1:A10))", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) {
    values[`A${r}`] = r; // A1..A10 contain numbers
  }

  const currentInput = "=SUM(-A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(-A1:A10)"),
    `Expected a unary '-' range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Range suggestions do not delete trailing whitespace (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) values[`A${r}`] = r;

  // The typed prefix ends with whitespace after a token ("A "), and any valid
  // completion would need to delete that whitespace (not a pure insertion).
  const currentInput = "=SUM(A ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.equal(suggestions.length, 0);
});

test("Sheet-prefix completions still work when the user types a space inside a quoted sheet name", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('My ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('My Sheet'!"),
    `Expected a quoted sheet-prefix completion for an in-name space, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Typing =SUM(A1 suggests auto-closing parens even when the referenced cell is empty", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(A1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1)"),
    `Expected a pure paren-close suggestion for a single-cell arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM(A1:A10 suggests auto-closing parens even when the range has no data", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(A1:A10";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(A1:A10)"),
    `Expected a pure paren-close suggestion even when the range is empty, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM(Sheet2!A1 suggests auto-closing parens without needing schema", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(Sheet2!A1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Sheet2!A1)"),
    `Expected a pure paren-close suggestion for a sheet-qualified cell ref, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM('My Sheet'!A1 suggests auto-closing parens without needing schema", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM('My Sheet'!A1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('My Sheet'!A1)"),
    `Expected a pure paren-close suggestion for a quoted sheet-qualified cell ref, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =SUM(Table1[Amount] suggests auto-closing parens without needing schema", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SUM(Table1[Amount]";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Table1[Amount])"),
    `Expected a pure paren-close suggestion for a structured reference, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =ABS(A1 suggests auto-closing parens for a value arg cell reference", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ABS(A1";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ABS(A1)"),
    `Expected ABS to suggest closing parens after a cell reference, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =ABS(Table1[Amount] suggests auto-closing parens for a structured reference value arg", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ABS(Table1[Amount]";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ABS(Table1[Amount])"),
    `Expected ABS to suggest closing parens after a structured reference, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =ABS(5 suggests auto-closing parens for a numeric literal", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ABS(5";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ABS(5)"),
    `Expected ABS to suggest closing parens after a numeric literal, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test('Typing =TEXT(A1,"yyyy-mm-dd" suggests auto-closing parens after a complete string enum', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXT(A1,"yyyy-mm-dd"';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXT(A1,"yyyy-mm-dd")'),
    `Expected TEXT to suggest closing parens after a complete string enum, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Typing =VLOOKUP(..., TRUE suggests auto-closing parens after a complete boolean literal", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:A10, 2, TRUE";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({ A1: 1 }),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP(A1, A1:A10, 2, TRUE)"),
    `Expected VLOOKUP to suggest closing parens after TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Auto-close parens works for decimal-comma numeric literals (semicolon separators)", async () => {
  const engine = new TabCompletionEngine();

  // `;` triggers semicolon-arg parsing; `,` remains inside the current arg (decimal comma locale style).
  const currentInput = "=IF(TRUE;1;1,2";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IF(TRUE;1;1,2)"),
    `Expected IF to suggest closing parens for a 1,2 literal, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Auto-close parens works for leading-dot numeric literals (.5)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=IF(TRUE;1;.5";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IF(TRUE;1;.5)"),
    `Expected IF to suggest closing parens for a .5 literal, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Auto-close parens works for percent literals (50%)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=IF(TRUE;1;50%";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IF(TRUE;1;50%)"),
    `Expected IF to suggest closing parens for a 50% literal, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("LARGE k suggests 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=LARGE(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected LARGE to suggest k=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect LARGE to suggest k=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SMALL k suggests 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SMALL(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected SMALL to suggest k=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect SMALL to suggest k=0, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Value-arg suggestions are pure insertions (ABS suggests left cell only when prefix matches)", async () => {
  const engine = new TabCompletionEngine();

  // Empty prefix: suggest left cell (A1) when editing in B1.
  const emptyInput = "=ABS(";
  const emptySuggestions = await engine.getSuggestions({
    currentInput: emptyInput,
    cursorPosition: emptyInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });
  assert.ok(
    emptySuggestions.some((s) => s.text === "=ABS(A1"),
    `Expected ABS to suggest left cell for empty arg, got: ${emptySuggestions.map((s) => s.text).join(", ")}`
  );

  // Matching prefix: still suggest.
  const aInput = "=ABS(A";
  const aSuggestions = await engine.getSuggestions({
    currentInput: aInput,
    cursorPosition: aInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });
  assert.ok(
    aSuggestions.some((s) => s.text === "=ABS(A1"),
    `Expected ABS to suggest A1 for the 'A' prefix, got: ${aSuggestions.map((s) => s.text).join(", ")}`
  );

  // Non-matching prefix: do not suggest (would require deleting typed text).
  const vInput = "=ABS(V";
  const vSuggestions = await engine.getSuggestions({
    currentInput: vInput,
    cursorPosition: vInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });
  assert.equal(vSuggestions.length, 0);
});

test("Function name completion works inside non-range args (=IF(VLO → =IF(VLOOKUP()", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=IF(VLO";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IF(VLOOKUP("),
    `Expected IF to suggest VLOOKUP( inside an arg, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Numeric argument suggestions work with a unary '-' prefix", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANDBETWEEN(-";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=RANDBETWEEN(-1"),
    `Expected RANDBETWEEN to suggest -1 after '-', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Value argument left-cell reference preserves the typed prefix (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ABS(A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're in column B so the cell to the left is A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ABS(A1"),
    `Expected ABS to suggest the left cell ref as a pure insertion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Value argument left-cell reference is not suggested when it would delete typed text (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ABS(C";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're in column B so the cell to the left is A1. Since the user typed "C",
    // suggesting "A1" would require deleting the "C".
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(
    suggestions.length,
    0,
    `Expected no suggestions (pure insertion), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test('DATEDIF unit suggests "d", "m", "y", "ym", "yd"', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=DATEDIF(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const unit of ['"d"', '"m"', '"y"', '"ym"', '"yd"']) {
    assert.ok(
      suggestions.some((s) => s.text === `=DATEDIF(A1, B1, ${unit}`),
      `Expected DATEDIF to suggest unit=${unit}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test('DATEDIF unit suggests "md" when typing the "\"m\" prefix', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=DATEDIF(A1, B1, "m';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=DATEDIF(A1, B1, "md"'),
    `Expected DATEDIF to suggest unit=\"md\" for the \"m\" prefix, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test('DATEDIF unit is not suggested for an unquoted prefix (not a pure insertion)', async () => {
  const engine = new TabCompletionEngine();

  // The curated enum entries are quoted strings (e.g. "d"). If the user hasn't started
  // the quote, inserting it would require modifying text before the caret.
  const currentInput = "=DATEDIF(A1, B1, d";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
});

test('CELL info_type suggests "address", "col", "row"', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CELL(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const infoType of ['"address"', '"col"', '"row"']) {
    assert.ok(
      suggestions.some((s) => s.text === `=CELL(${infoType}`),
      `Expected CELL to suggest info_type=${infoType}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test('INFO type_text suggests "osversion" and "system"', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=INFO(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const typeText of ['"osversion"', '"system"']) {
    assert.ok(
      suggestions.some((s) => s.text === `=INFO(${typeText}`),
      `Expected INFO to suggest type_text=${typeText}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("IMAGE sizing suggests 0, 1, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=IMAGE("https://example.com/cat.png", "cat", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const code of ["0", "1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=IMAGE(\"https://example.com/cat.png\", \"cat\", ${code}`),
      `Expected IMAGE to suggest sizing=${code}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test('NUMBERVALUE decimal_separator suggests "." and ","', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=NUMBERVALUE("1,23", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=NUMBERVALUE("1,23", "."'),
    `Expected NUMBERVALUE to suggest decimal_separator=\".\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=NUMBERVALUE("1,23", ","'),
    `Expected NUMBERVALUE to suggest decimal_separator=\",\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test('NUMBERVALUE group_separator suggests ",", ".", and " "', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=NUMBERVALUE("1.234,56", ",", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const sep of ['","', '"."', '" "']) {
    assert.ok(
      suggestions.some((s) => s.text === `=NUMBERVALUE("1.234,56", ",", ${sep}`),
      `Expected NUMBERVALUE to suggest group_separator=${sep}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test('NUMBERVALUE group_separator completes a quoted space when typing "\" \"', async () => {
  const engine = new TabCompletionEngine();

  // Trailing whitespace here is *inside* an unterminated string literal, so it should
  // still allow pure-insertion completions like `" "` -> `" "`.
  const currentInput = '=NUMBERVALUE("1.234,56", ",", " ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=NUMBERVALUE("1.234,56", ",", " "'),
    `Expected NUMBERVALUE to complete a quoted space, got: ${suggestions.map((s) => JSON.stringify(s.text)).join(", ")}`
  );
});

test('TEXT format_text suggests common format strings', async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TEXT(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const fmt of ['"0"', '"0.00"', '"0%"']) {
    assert.ok(
      suggestions.some((s) => s.text === `=TEXT(A1, ${fmt}`),
      `Expected TEXT to suggest format_text=${fmt}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
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

test("VLOOKUP col_index_num suggests 2, 1, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:B10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["2", "1", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected VLOOKUP to suggest col_index_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect VLOOKUP to suggest col_index_num=0, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("HLOOKUP row_index_num suggests 2, 1, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=HLOOKUP(A1, A1:B10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["2", "1", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected HLOOKUP to suggest row_index_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect HLOOKUP to suggest row_index_num=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("INDEX area_num suggests 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=INDEX(A1:B10, 1, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected INDEX to suggest area_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect INDEX to suggest area_num=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CHOOSE index_num suggests 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHOOSE(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected CHOOSE to suggest index_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect CHOOSE to suggest index_num=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("VLOOKUP range_lookup preserves typed casing for booleans (lowercase prefix)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:B10, 2, f";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP(A1, A1:B10, 2, false"),
    `Expected VLOOKUP to complete "f" -> "false", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("VLOOKUP range_lookup preserves typed casing for booleans (title-case prefix)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:B10, 2, Fa";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP(A1, A1:B10, 2, False"),
    `Expected VLOOKUP to complete \"Fa\" -> \"False\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("VLOOKUP range_lookup suggestions work inside grouping parens (preserves typed casing)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=VLOOKUP(A1, A1:B10, 2, (f";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=VLOOKUP(A1, A1:B10, 2, (false"),
    `Expected VLOOKUP to complete \"(f\" -> \"(false\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Argument value enum suggestions do not delete trailing whitespace after grouping parens (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  // User typed trailing whitespace after starting a grouped boolean literal. Any completion
  // would need to delete that whitespace, so the engine should return no suggestions.
  const currentInput = "=VLOOKUP(A1, A1:B10, 2, (F ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
});

test("Argument value enum suggestions do not delete trailing whitespace (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  // User typed trailing whitespace after starting a boolean literal. Any completion
  // would need to delete that whitespace, so the engine should return no suggestions.
  const currentInput = "=VLOOKUP(A1, A1:B10, 2, F ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
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

test("SORT sort_index suggests 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SORT(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected SORT to suggest sort_index=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect SORT to suggest sort_index=0, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("SORTBY repeating sort_order suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  // Test the second sort_order position to ensure repeating-group enum mapping works:
  // SORTBY(array, by_array1, sort_order1, by_array2, sort_order2, ...)
  const currentInput = "=SORTBY(A1:A10, B1:B10, 1, C1:C10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SORTBY(A1:A10, B1:B10, 1, C1:C10, 1"),
    `Expected SORTBY to suggest sort_order2=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=SORTBY(A1:A10, B1:B10, 1, C1:C10, -1"),
    `Expected SORTBY to suggest sort_order2=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Positive count args suggest 1, 2, 3 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const cases = [
    { name: "SEQUENCE rows", currentInput: "=SEQUENCE(" },
    { name: "SEQUENCE columns", currentInput: "=SEQUENCE(5, " },
    { name: "MAKEARRAY rows", currentInput: "=MAKEARRAY(" },
    { name: "MAKEARRAY columns", currentInput: "=MAKEARRAY(5, " },
    { name: "RANDARRAY rows", currentInput: "=RANDARRAY(" },
    { name: "RANDARRAY columns", currentInput: "=RANDARRAY(5, " },
    { name: "EXPAND rows", currentInput: "=EXPAND(A1:A10, " },
    { name: "EXPAND columns", currentInput: "=EXPAND(A1:A10, 5, " },
    { name: "WRAPROWS wrap_count", currentInput: "=WRAPROWS(A1:A10, " },
    { name: "WRAPCOLS wrap_count", currentInput: "=WRAPCOLS(A1:A10, " },
    { name: "OFFSET height", currentInput: "=OFFSET(A1, 0, 0, " },
    { name: "OFFSET width", currentInput: "=OFFSET(A1, 0, 0, 5, " },
  ];

  for (const { name, currentInput } of cases) {
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const v of ["1", "2", "3"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${v}`),
        `Expected ${name} to suggest ${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }

    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}0`),
      `Did not expect ${name} to suggest 0, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
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

test("TAKE rows enum suggestions work with a unary '-' prefix (no '--1')", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TAKE(A1:A10,-";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TAKE(A1:A10,-1"),
    `Expected TAKE to suggest rows=-1 after '-', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    !suggestions.some((s) => s.text === "=TAKE(A1:A10,--1"),
    `Did not expect TAKE to suggest '--1' after '-', got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CHOOSECOLS col_num suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHOOSECOLS(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSECOLS(A1:A10, 1"),
    `Expected CHOOSECOLS to suggest col_num=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSECOLS(A1:A10, -1"),
    `Expected CHOOSECOLS to suggest col_num=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CHOOSECOLS repeating col_num suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHOOSECOLS(A1:A10, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSECOLS(A1:A10, 1, 1"),
    `Expected CHOOSECOLS to suggest col_num2=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSECOLS(A1:A10, 1, -1"),
    `Expected CHOOSECOLS to suggest col_num2=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CHOOSEROWS row_num suggests 1 and -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHOOSEROWS(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSEROWS(A1:A10, 1"),
    `Expected CHOOSEROWS to suggest row_num=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CHOOSEROWS(A1:A10, -1"),
    `Expected CHOOSEROWS to suggest row_num=-1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TOCOL ignore suggests 0, 1, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TOCOL(A1:B2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["0", "1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=TOCOL(A1:B2, ${v}`),
      `Expected TOCOL to suggest ignore=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("TOCOL scan_by_column suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TOCOL(A1:B2, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TOCOL(A1:B2, , FALSE"),
    `Expected TOCOL to suggest scan_by_column=FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=TOCOL(A1:B2, , TRUE"),
    `Expected TOCOL to suggest scan_by_column=TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TOROW ignore suggests 0, 1, 2, 3", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TOROW(A1:B2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["0", "1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=TOROW(A1:B2, ${v}`),
      `Expected TOROW to suggest ignore=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("TEXTSPLIT col_delimiter suggests common delimiters", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTSPLIT("aXb", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("aXb", ","'),
    `Expected TEXTSPLIT to suggest col_delimiter=\",\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTSPLIT("aXb", " "'),
    `Expected TEXTSPLIT to suggest col_delimiter=\" \", got: ${suggestions.map((s) => JSON.stringify(s.text)).join(", ")}`
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

test("TEXTAFTER delimiter suggests common delimiters", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTAFTER("aXb", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", ","'),
    `Expected TEXTAFTER to suggest delimiter=\",\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", " "'),
    `Expected TEXTAFTER to suggest delimiter=\" \", got: ${suggestions.map((s) => JSON.stringify(s.text)).join(", ")}`
  );
});

test("TEXTAFTER instance_num suggests 1, 2, -1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTAFTER("aXbXc", "X", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "-1"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=TEXTAFTER("aXbXc", "X", ${v}`),
      `Expected TEXTAFTER to suggest instance_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("TEXTAFTER match_mode suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=TEXTAFTER("aXb", "x", 1, ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", "x", 1, 0'),
    `Expected TEXTAFTER to suggest match_mode=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", "x", 1, 1'),
    `Expected TEXTAFTER to suggest match_mode=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TEXTAFTER match_end suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  // Leave match_mode blank so we're completing match_end (5th arg).
  const currentInput = '=TEXTAFTER("aXb", "x", 1, , ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", "x", 1, , 0'),
    `Expected TEXTAFTER to suggest match_end=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTAFTER("aXb", "x", 1, , 1'),
    `Expected TEXTAFTER to suggest match_end=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("TEXTJOIN delimiter suggests common delimiters", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TEXTJOIN(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=TEXTJOIN(","'),
    `Expected TEXTJOIN to suggest delimiter=\",\", got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === '=TEXTJOIN(" "'),
    `Expected TEXTJOIN to suggest delimiter=\" \", got: ${suggestions.map((s) => JSON.stringify(s.text)).join(", ")}`
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

test("SUBSTITUTE instance_num suggests 1 and 2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=SUBSTITUTE(A1, "x", "y", ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected SUBSTITUTE to suggest instance_num=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect SUBSTITUTE to suggest instance_num=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Text start_num args suggest 1 and 2 (FIND/SEARCH/MID/REPLACE)", async () => {
  const engine = new TabCompletionEngine();

  const cases = [
    '=FIND("x", A1, ',
    '=SEARCH("x", A1, ',
    "=MID(A1, ",
    "=REPLACE(A1, ",
  ];

  for (const currentInput of cases) {
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const v of ["1", "2"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${v}`),
        `Expected ${currentInput}... to suggest ${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }

    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}0`),
      `Did not expect ${currentInput}... to suggest 0, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
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

test("TDIST tails suggests 1 and 2", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=TDIST(1, 10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=TDIST(1, 10, 1"),
    `Expected TDIST to suggest tails=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=TDIST(1, 10, 2"),
    `Expected TDIST to suggest tails=2, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("ROMAN form suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ROMAN(42, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const form of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=ROMAN(42, ${form}`),
      `Expected ROMAN to suggest form=${form}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
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

test("WORKDAY.INTL weekend suggests 1, 2, 7, 11, 17", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=WORKDAY.INTL(A1, 5, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "7", "11", "17"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=WORKDAY.INTL(A1, 5, ${v}`),
      `Expected WORKDAY.INTL to suggest weekend=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("NETWORKDAYS.INTL weekend suggests 1, 2, 7, 11, 17", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=NETWORKDAYS.INTL(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["1", "2", "7", "11", "17"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=NETWORKDAYS.INTL(A1, B1, ${v}`),
      `Expected NETWORKDAYS.INTL to suggest weekend=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
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

test("PRICE frequency suggests 2, 1, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PRICE(A1, B1, 0.05, 0.04, 100, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const f of ["2", "1", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=PRICE(A1, B1, 0.05, 0.04, 100, ${f}`),
      `Expected PRICE to suggest frequency=${f}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("PRICE basis suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PRICE(A1, B1, 0.05, 0.04, 100, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const basis of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=PRICE(A1, B1, 0.05, 0.04, 100, 2, ${basis}`),
      `Expected PRICE to suggest basis=${basis}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("ACCRINT frequency suggests 2, 1, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ACCRINT(A1, B1, C1, 0.05, 100, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const f of ["2", "1", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=ACCRINT(A1, B1, C1, 0.05, 100, ${f}`),
      `Expected ACCRINT to suggest frequency=${f}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("ACCRINT calc_method suggests TRUE/FALSE with meaning", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ACCRINT(A1, B1, C1, 0.05, 100, 2, 0, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  const fromIssue = suggestions.find((s) => s.text === "=ACCRINT(A1, B1, C1, 0.05, 100, 2, 0, FALSE");
  assert.ok(fromIssue, `Expected ACCRINT to suggest FALSE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok(
    (fromIssue?.confidence ?? 0) > 0.5,
    `Expected ACCRINT/FALSE to have elevated confidence, got: ${fromIssue?.confidence}`
  );

  const fromCoupon = suggestions.find((s) => s.text === "=ACCRINT(A1, B1, C1, 0.05, 100, 2, 0, TRUE");
  assert.ok(fromCoupon, `Expected ACCRINT to suggest TRUE, got: ${suggestions.map((s) => s.text).join(", ")}`);
  assert.ok(
    (fromCoupon?.confidence ?? 0) > 0.5,
    `Expected ACCRINT/TRUE to have elevated confidence, got: ${fromCoupon?.confidence}`
  );
});

test("PRICEDISC basis suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PRICEDISC(A1, B1, 0.05, 100, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const basis of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=PRICEDISC(A1, B1, 0.05, 100, ${basis}`),
      `Expected PRICEDISC to suggest basis=${basis}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("ODDFPRICE frequency suggests 2, 1, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ODDFPRICE(A1, B1, C1, D1, 0.05, 0.04, 100, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const f of ["2", "1", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=ODDFPRICE(A1, B1, C1, D1, 0.05, 0.04, 100, ${f}`),
      `Expected ODDFPRICE to suggest frequency=${f}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("INTRATE basis suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=INTRATE(A1, B1, 100, 110, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const basis of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=INTRATE(A1, B1, 100, 110, ${basis}`),
      `Expected INTRATE to suggest basis=${basis}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("COUPDAYBS frequency suggests 2, 1, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=COUPDAYBS(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const f of ["2", "1", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=COUPDAYBS(A1, B1, ${f}`),
      `Expected COUPDAYBS to suggest frequency=${f}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("COUPDAYBS basis suggests 0, 1, 2, 3, 4", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=COUPDAYBS(A1, B1, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const basis of ["0", "1", "2", "3", "4"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=COUPDAYBS(A1, B1, 2, ${basis}`),
      `Expected COUPDAYBS to suggest basis=${basis}, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("FunctionRegistry uses correct optional arg ordering for FORECAST.ETS.STAT", () => {
  const registry = new FunctionRegistry();
  const spec = registry.getFunction("FORECAST.ETS.STAT");
  assert.ok(spec, "Expected FORECAST.ETS.STAT to exist in registry");

  assert.equal(spec?.args?.[0]?.name, "values");
  assert.equal(spec?.args?.[1]?.name, "timeline");
  assert.equal(spec?.args?.[2]?.name, "seasonality");
  assert.equal(spec?.args?.[3]?.name, "data_completion");
  assert.equal(spec?.args?.[4]?.name, "aggregation");
  assert.equal(spec?.args?.[5]?.name, "statistic_type");
});

test("FORECAST.ETS.STAT statistic_type suggests 8 (RMSE) and 7 (MAE)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.ETS.STAT(B1:B10, C1:C10, 1, 1, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS.STAT(B1:B10, C1:C10, 1, 1, 1, 8"),
    `Expected FORECAST.ETS.STAT to suggest statistic_type=8, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.ETS.STAT(B1:B10, C1:C10, 1, 1, 1, 7"),
    `Expected FORECAST.ETS.STAT to suggest statistic_type=7, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Rounding significance args suggest common values (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const cases = [
    { name: "CEILING.MATH significance", currentInput: "=CEILING.MATH(A1, " },
    { name: "FLOOR.MATH significance", currentInput: "=FLOOR.MATH(A1, " },
    { name: "CEILING significance", currentInput: "=CEILING(A1, " },
    { name: "FLOOR significance", currentInput: "=FLOOR(A1, " },
    { name: "CEILING.PRECISE significance", currentInput: "=CEILING.PRECISE(A1, " },
    { name: "FLOOR.PRECISE significance", currentInput: "=FLOOR.PRECISE(A1, " },
    { name: "ISO.CEILING significance", currentInput: "=ISO.CEILING(A1, " },
    { name: "MROUND multiple", currentInput: "=MROUND(A1, " },
  ];

  for (const { name, currentInput } of cases) {
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const v of ["1", "0.1", "10"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${v}`),
        `Expected ${name} to suggest ${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }

    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}0`),
      `Did not expect ${name} to suggest 0, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("MOD divisor suggests 2 and 10 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=MOD(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["2", "10"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected MOD to suggest divisor=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  assert.ok(
    !suggestions.some((s) => s.text === `${currentInput}0`),
    `Did not expect MOD to suggest divisor=0, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("HYPGEOM.DIST scalar args suggest a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=HYPGEOM.DIST(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=HYPGEOM.DIST(A1"),
    `Expected HYPGEOM.DIST to suggest a left-cell reference, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("WORKDAY days suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=WORKDAY(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in C1 so the left-cell heuristic suggests B1.
    cellRef: { row: 0, col: 2 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=WORKDAY(A1, B1"),
    `Expected WORKDAY to suggest a left-cell reference for days, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("RANK.EQ number suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANK.EQ(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=RANK.EQ(A1"),
    `Expected RANK.EQ to suggest a left-cell reference for number, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("PERCENTRANK x suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PERCENTRANK(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in C1 so the left-cell heuristic suggests B1.
    cellRef: { row: 0, col: 2 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=PERCENTRANK(A1:A10, B1"),
    `Expected PERCENTRANK to suggest a left-cell reference for x, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("DELTA number1 suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=DELTA(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=DELTA(A1"),
    `Expected DELTA to suggest a left-cell reference for number1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("BESSELI x suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=BESSELI(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=BESSELI(A1"),
    `Expected BESSELI to suggest a left-cell reference for x, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("FORECAST.LINEAR x suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=FORECAST.LINEAR(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=FORECAST.LINEAR(A1"),
    `Expected FORECAST.LINEAR to suggest a left-cell reference for x, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SERIESSUM x suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SERIESSUM(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in B1 so the left-cell heuristic suggests A1.
    cellRef: { row: 0, col: 1 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SERIESSUM(A1"),
    `Expected SERIESSUM to suggest a left-cell reference for x, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SERIESSUM n suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SERIESSUM(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in C1 so the left-cell heuristic suggests B1.
    cellRef: { row: 0, col: 2 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SERIESSUM(A1, B1"),
    `Expected SERIESSUM to suggest a left-cell reference for n, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("INDEX row_num suggests a left-cell reference (value-like)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=INDEX(A1:B10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Place the caret in C1 so the left-cell heuristic suggests B1.
    cellRef: { row: 0, col: 2 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=INDEX(A1:B10, B1"),
    `Expected INDEX to suggest a left-cell reference for row_num, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("QUARTILE.EXC quart suggests 1, 2, 3 (no 0/4)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=QUARTILE.EXC(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const q of ["1", "2", "3"]) {
    assert.ok(
      suggestions.some((s) => s.text === `=QUARTILE.EXC(A1:A10, ${q}`),
      `Expected QUARTILE.EXC to suggest quart=${q}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  for (const q of ["0", "4"]) {
    assert.ok(
      !suggestions.some((s) => s.text === `=QUARTILE.EXC(A1:A10, ${q}`),
      `Did not expect QUARTILE.EXC to suggest quart=${q}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("PERCENTILE k suggests common values (including 0 and 1)", async () => {
  const engine = new TabCompletionEngine();

  for (const fn of ["PERCENTILE", "PERCENTILE.INC"]) {
    const currentInput = `=${fn}(A1:A10, `;
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const k of ["0.5", "0.25", "0.75", "0", "1"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${k}`),
        `Expected ${fn} to suggest k=${k}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }
  }
});

test("PERCENTILE.EXC k suggests common values (no 0/1)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PERCENTILE.EXC(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const k of ["0.5", "0.25", "0.75"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${k}`),
      `Expected PERCENTILE.EXC to suggest k=${k}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  for (const k of ["0", "1"]) {
    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}${k}`),
      `Did not expect PERCENTILE.EXC to suggest k=${k}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("PERCENTRANK significance suggests 3, 2, 1 (no 0)", async () => {
  const engine = new TabCompletionEngine();

  for (const fn of ["PERCENTRANK", "PERCENTRANK.INC", "PERCENTRANK.EXC"]) {
    const currentInput = `=${fn}(A1:A10, 5, `;
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const v of ["3", "2", "1"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${v}`),
        `Expected ${fn} to suggest significance=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }
    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}0`),
      `Did not expect ${fn} to suggest significance=0, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("LOG base suggests 10 and 2 (no 0/1)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=LOG(100, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  for (const v of ["10", "2"]) {
    assert.ok(
      suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Expected LOG to suggest base=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  for (const v of ["0", "1"]) {
    assert.ok(
      !suggestions.some((s) => s.text === `${currentInput}${v}`),
      `Did not expect LOG to suggest base=${v}, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("BASE/DECIMAL radix suggests common values (no 0/1)", async () => {
  const engine = new TabCompletionEngine();

  const cases = [
    { fn: "BASE", currentInput: "=BASE(255, " },
    { fn: "DECIMAL", currentInput: '=DECIMAL("FF", ' },
  ];

  for (const { fn, currentInput } of cases) {
    const suggestions = await engine.getSuggestions({
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    });

    for (const radix of ["10", "2", "16"]) {
      assert.ok(
        suggestions.some((s) => s.text === `${currentInput}${radix}`),
        `Expected ${fn} to suggest radix=${radix}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }

    for (const radix of ["0", "1"]) {
      assert.ok(
        !suggestions.some((s) => s.text === `${currentInput}${radix}`),
        `Did not expect ${fn} to suggest radix=${radix}, got: ${suggestions.map((s) => s.text).join(", ")}`
      );
    }
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

test("POISSON cumulative suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=POISSON(1, 2, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=POISSON(1, 2, TRUE"),
    `Expected POISSON to suggest TRUE (cumulative), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=POISSON(1, 2, FALSE"),
    `Expected POISSON to suggest FALSE (probability), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("WEIBULL cumulative suggests TRUE/FALSE", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=WEIBULL(1, 2, 3, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=WEIBULL(1, 2, 3, TRUE"),
    `Expected WEIBULL to suggest TRUE (cumulative), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=WEIBULL(1, 2, 3, FALSE"),
    `Expected WEIBULL to suggest FALSE (probability), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("PV rate suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PV(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=PV(A1"),
    `Expected PV to suggest A1 for rate, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("IRR guess suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=IRR(A1:A10, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IRR(A1:A10, B1"),
    `Expected IRR to suggest B1 for guess, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("PRICE rate suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PRICE(A1, B1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 3 }, // D1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=PRICE(A1, B1, C1"),
    `Expected PRICE to suggest C1 for rate, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("ADDRESS row_num suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=ADDRESS(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=ADDRESS(A1"),
    `Expected ADDRESS to suggest A1 for row_num, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("IMAGE height suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = '=IMAGE("https://example.com/cat.png", "cat", 1, ';
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === '=IMAGE("https://example.com/cat.png", "cat", 1, A1'),
    `Expected IMAGE to suggest A1 for height, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CHAR number suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CHAR(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CHAR(A1"),
    `Expected CHAR to suggest A1 for number, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("UNICHAR number suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=UNICHAR(";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 1 }, // B1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=UNICHAR(A1"),
    `Expected UNICHAR to suggest A1 for number, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("IMPOWER exponent suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=IMPOWER(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=IMPOWER(A1, B1"),
    `Expected IMPOWER to suggest B1 for exponent, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("SEQUENCE start suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=SEQUENCE(5, 1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SEQUENCE(5, 1, B1"),
    `Expected SEQUENCE to suggest B1 for start, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("RANDARRAY min suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=RANDARRAY(5, 5, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=RANDARRAY(5, 5, B1"),
    `Expected RANDARRAY to suggest B1 for min, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("DOLLAR decimals suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=DOLLAR(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=DOLLAR(A1, B1"),
    `Expected DOLLAR to suggest B1 for decimals, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("REPT number_times suggests a left-cell reference (value-like arg)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=REPT(A1, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 2 }, // C1
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=REPT(A1, B1"),
    `Expected REPT to suggest B1 for number_times, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("PMT type suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=PMT(0.05/12, 60, 10000, , ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=PMT(0.05/12, 60, 10000, , 0"),
    `Expected PMT to suggest type=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=PMT(0.05/12, 60, 10000, , 1"),
    `Expected PMT to suggest type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("CUMIPMT type suggests 0 and 1", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "=CUMIPMT(0.05/12, 60, 10000, 1, 12, ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=CUMIPMT(0.05/12, 60, 10000, 1, 12, 0"),
    `Expected CUMIPMT to suggest type=0, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.ok(
    suggestions.some((s) => s.text === "=CUMIPMT(0.05/12, 60, 10000, 1, 12, 1"),
    `Expected CUMIPMT to suggest type=1, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Named ranges title-case ALL-CAPS identifiers for a Unicode title-style prefix", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [{ name: "ZÄHLENWENN", range: "Sheet1!A1:A10" }],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Zä";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Zählenwenn)"),
    `Expected a Unicode title-cased named-range completion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Named ranges title-case ALL-CAPS identifiers across segments (underscore)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [{ name: "FOO_BAR", range: "Sheet1!A1:A10" }],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM(Foo_Ba";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Foo_Bar)"),
    `Expected a segment-aware title-cased named-range completion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Backend full-formula completions are only accepted when they are pure insertions", async () => {
  // 1) Accept a full-formula completion when it strictly extends the current input.
  {
    const completionClient = {
      async completeTabCompletion() {
        return "=1+2";
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

    assert.ok(
      suggestions.some((s) => s.text === "=1+2"),
      `Expected a full-formula backend completion to be accepted, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }

  // 2) Reject a full-formula completion that would rewrite user-typed characters (not a pure insertion).
  {
    const completionClient = {
      async completeTabCompletion() {
        return "=2";
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

    assert.equal(
      suggestions.length,
      0,
      `Expected backend rewrite completion to be dropped, got: ${suggestions.map((s) => s.text).join(", ")}`
    );
  }
});

test("TabCompletionEngine forwards AbortSignal to completionClient and aborts when caller cancels", async () => {
  let calls = 0;
  let sawAbort = false;

  const completionClient = {
    async completeTabCompletion(req) {
      calls += 1;
      return await new Promise((resolve) => {
        const done = () => {
          sawAbort = true;
          resolve("");
        };
        if (req?.signal?.aborted) return done();
        req?.signal?.addEventListener("abort", done, { once: true });
      });
    },
  };

  const engine = new TabCompletionEngine({ completionClient, completionTimeoutMs: 200 });

  const currentInput = "=1+";
  const controller = new AbortController();
  const promise = engine.getSuggestions(
    {
      currentInput,
      cursorPosition: currentInput.length,
      cellRef: { row: 0, col: 0 },
      surroundingCells: createMockCellContext({}),
    },
    { signal: controller.signal },
  );

  controller.abort();

  const suggestions = await promise;
  assert.equal(calls, 1);
  assert.equal(sawAbort, true);
  assert.ok(Array.isArray(suggestions));
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

test("Structured references do not delete trailing whitespace (pure insertion)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["Amount"] }],
    },
  });

  const currentInput = "=SUM(Tab ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.equal(suggestions.length, 0);
});

test("Structured references support column names with spaces (pure insertion)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["First Name"] }],
    },
  });

  const currentInput = "=SUM(Table1[First ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM(Table1[First Name])"),
    `Expected a structured ref completion for a spaced column name, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Structured references are suggested inside brackets at top level (=Table1[First ␠ → =Table1[First Name])", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1"],
      getTables: () => [{ name: "Table1", columns: ["First Name"] }],
    },
  });

  const currentInput = "=Table1[First ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 10, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Table1[First Name]"),
    `Expected a top-level structured ref completion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
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

test("Sheet-name prefixes are suggested inside grouping parens (=SUM((she → =SUM((sheet2!)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Sheet1", "Sheet2", "My Sheet", "A1"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM((she";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => ["=SUM((sheet1!", "=SUM((sheet2!"].includes(s.text)),
    `Expected a grouped sheet prefix suggestion ending with '!', got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Quoted sheet-name prefix completions escape apostrophes (Bob's Sheet → 'Bob''s Sheet'!)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["Bob's Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('Bo";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('Bob''s Sheet'!"),
    `Expected a quoted sheet-prefix suggestion with escaped apostrophe, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
  assert.equal(
    suggestions.filter((s) => s.text.endsWith("!)")).length,
    0,
    `Expected escaped quoted sheet-prefix suggestions to not auto-close parens, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Quoted sheet-name prefix completions preserve schema case when only \"'\" is typed", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "=SUM('";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=SUM('My Sheet'!"),
    `Expected sheet-prefix suggestion to preserve schema casing, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("Sheet-qualified ranges are suggested at top level (=Sheet2!A → =Sheet2!A1:A10)", async () => {
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

  const currentInput = "=Sheet2!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.ok(
    suggestions.some((s) => s.text === "=Sheet2!A1:A10"),
    `Expected a top-level sheet-qualified range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Top-level A1 range suggestions work (=A1: → =A1:A10)", async () => {
  const engine = new TabCompletionEngine();

  const values = {};
  for (let r = 1; r <= 10; r++) values[`A${r}`] = r;

  const currentInput = "=A1:";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: createMockCellContext(values),
  });

  assert.ok(
    suggestions.some((s) => s.text === "=A1:A10"),
    `Expected a top-level A1 range suggestion, got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Sheet-qualified VLOOKUP table_array prefers a 2D table range when adjacent columns form a table", async () => {
  const values = {};
  // Header row.
  values["Sheet2!A1"] = "Key";
  values["Sheet2!B1"] = "Value1";
  values["Sheet2!C1"] = "Value2";
  values["Sheet2!D1"] = "Value3";
  // Data rows 2..10.
  for (let r = 2; r <= 10; r++) {
    values[`Sheet2!A${r}`] = `K${r}`;
    values[`Sheet2!B${r}`] = r * 10;
    values[`Sheet2!C${r}`] = r * 100;
    values[`Sheet2!D${r}`] = r * 1000;
  }

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

  const currentInput = "=VLOOKUP(A1, Sheet2!A";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Pretend we're on row 11 (0-based 10), below the data.
    cellRef: { row: 10, col: 1 },
    surroundingCells: cellContext,
  });

  assert.equal(suggestions[0]?.text, "=VLOOKUP(A1, Sheet2!A1:D10");
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

test("Quoted sheet names are suggested as prefixes when the user starts a quote (=\'my → =\'my Sheet\'!)", async () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getNamedRanges: () => [],
      getSheetNames: () => ["My Sheet"],
      getTables: () => [],
    },
  });

  const currentInput = "='my";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 0, col: 0 },
    surroundingCells: createMockCellContext({}),
  });

  assert.ok(
    suggestions.some((s) => s.text === "='my Sheet'!"),
    `Expected a quoted sheet-name completion, got: ${suggestions.map((s) => s.text).join(", ")}`
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

test("buildCacheKey never throws when cache-key providers throw", () => {
  const engine = new TabCompletionEngine({
    schemaProvider: {
      getCacheKey: () => {
        throw new Error("boom");
      },
    },
  });

  const key = engine.buildCacheKey({
    // @ts-ignore - intentionally invalid
    currentInput: 123,
    cursorPosition: -5,
    // @ts-ignore - intentionally invalid
    cellRef: "not-a1",
    surroundingCells: {
      getCellValue() {
        return null;
      },
      getCacheKey() {
        throw new Error("boom");
      },
    },
  });

  assert.equal(typeof key, "string");
  assert.ok(key.length > 0);
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

test("Pattern suggestions preserve typed prefix casing (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "ap";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    cellRef: { row: 5, col: 0 },
    surroundingCells: createMockCellContext({ A5: "Apple" }),
  });

  assert.ok(
    suggestions.some((s) => s.text === "apple"),
    `Expected a case-preserving pattern suggestion (apple), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Pattern suggestions preserve the typed suffix when cursor is mid-string", async () => {
  const engine = new TabCompletionEngine();

  const currentInput = "apX";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: 2, // cursor after "ap"
    cellRef: { row: 5, col: 0 },
    surroundingCells: createMockCellContext({ A5: "Apple" }),
  });

  assert.ok(
    suggestions.some((s) => s.text === "appleX"),
    `Expected a pure-insertion pattern suggestion (appleX), got: ${suggestions.map((s) => s.text).join(", ")}`
  );
});

test("Pattern numeric suggestions do not delete trailing whitespace (pure insertion)", async () => {
  const engine = new TabCompletionEngine();

  // Pretend the column above is a stable numeric sequence (10, 11) so the pattern
  // suggester would normally propose 12. If the user typed a trailing space, we
  // must not emit a suggestion that would require deleting it.
  const currentInput = "1 ";
  const suggestions = await engine.getSuggestions({
    currentInput,
    cursorPosition: currentInput.length,
    // Current cell is A3 (0-based row 2) so A2/A1 are "above".
    cellRef: { row: 2, col: 0 },
    surroundingCells: createMockCellContext({ A1: 10, A2: 11 }),
  });

  assert.equal(
    suggestions.length,
    0,
    `Expected no suggestions (pure insertion), got: ${suggestions.map((s) => JSON.stringify(s.text)).join(", ")}`
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
