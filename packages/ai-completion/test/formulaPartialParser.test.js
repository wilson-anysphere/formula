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

