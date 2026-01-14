import assert from "node:assert/strict";
import test from "node:test";

import { parsePartialFormula } from "../src/formulaPartialParser.js";
import { FunctionRegistry } from "../src/functionRegistry.js";

test("parsePartialFormula treats ';' as an argument separator", () => {
  const registry = new FunctionRegistry();
  const input = "=SUM(A1;";
  const parsed = parsePartialFormula(input, input.length, registry);

  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.start, input.length);
  assert.equal(parsed.currentArg?.end, input.length);
  assert.equal(parsed.currentArg?.text, "");
});

test("parsePartialFormula ignores ';' inside array constants { ... }", () => {
  const registry = new FunctionRegistry();
  const input = "=SUM({1;2}; A";
  const parsed = parsePartialFormula(input, input.length, registry);

  // The semicolon inside the array constant is a row separator, not a function arg separator.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula prefers ';' as the separator when both ';' and ',' appear (decimal comma locale)", () => {
  const registry = new FunctionRegistry();
  const input = "=VLOOKUP(1,2; A";
  const parsed = parsePartialFormula(input, input.length, registry);

  // The comma is part of a decimal literal in many locales; the semicolon is the arg separator.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula allows function-name completion after ';' (array constant row separator)", () => {
  const registry = new FunctionRegistry();
  const input = "={1;VLO";
  const parsed = parsePartialFormula(input, input.length, registry);

  assert.equal(parsed.isFormula, true);
  assert.equal(parsed.inFunctionCall, false);
  assert.deepEqual(parsed.functionNamePrefix, { text: "VLO", start: 4, end: 7 });
});

test("parsePartialFormula allows function-name completion after '{' (array constant start)", () => {
  const registry = new FunctionRegistry();
  const input = "={VLO";
  const parsed = parsePartialFormula(input, input.length, registry);

  assert.equal(parsed.isFormula, true);
  assert.equal(parsed.inFunctionCall, false);
  assert.deepEqual(parsed.functionNamePrefix, { text: "VLO", start: 2, end: 5 });
});

test("parsePartialFormula tracks quoted sheet names inside braces (sheet names may contain '(')", () => {
  const registry = new FunctionRegistry();
  // Sheet name is literally: My (Sheet   (no closing ')', which is valid in Excel).
  // The '(' should not be treated as a function call paren even though we're inside `{}`.
  const input = "={'My (Sheet'!A";
  const parsed = parsePartialFormula(input, input.length, registry);

  assert.equal(parsed.isFormula, true);
  assert.equal(parsed.inFunctionCall, false);
  assert.equal(parsed.functionName, undefined);
  assert.equal(parsed.functionNamePrefix, undefined);
});

test("parsePartialFormula ignores ';' inside quoted sheet names", () => {
  const registry = new FunctionRegistry();
  const input = "=SUM('Jan;2024'!A1; A";
  const parsed = parsePartialFormula(input, input.length, registry);

  // Only the semicolon *after* the A1 reference should split args.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula ignores ';' inside string literals", () => {
  const registry = new FunctionRegistry();
  const input = '=SUM("a;b"; A';
  const parsed = parsePartialFormula(input, input.length, registry);

  // Only the semicolon *after* the string literal should split args.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula ignores ';' inside structured references", () => {
  const registry = new FunctionRegistry();
  const input = "=SUM(Table1[Amount;USD]; A";
  const parsed = parsePartialFormula(input, input.length, registry);

  // Only the semicolon *after* the structured ref should split args.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula ignores ';' inside nested function calls (depth > baseDepth)", () => {
  const registry = new FunctionRegistry();
  const input = "=SUM(IF(A1>0;A1;0); A";
  const parsed = parsePartialFormula(input, input.length, registry);

  // Only the semicolon after the nested IF(...) should split SUM args.
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "A");
});

test("parsePartialFormula supports non-ASCII function names (localized identifiers)", () => {
  const registry = new FunctionRegistry();
  const input = "=zählenwenn(A1;"; // COUNTIF in German Excel.
  const parsed = parsePartialFormula(input, input.length, registry);

  assert.equal(parsed.isFormula, true);
  assert.equal(parsed.inFunctionCall, true);
  assert.equal(parsed.functionName, "ZÄHLENWENN");
  assert.equal(parsed.argIndex, 1);
  assert.equal(parsed.currentArg?.text, "");
});
