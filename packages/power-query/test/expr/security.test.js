import assert from "node:assert/strict";
import test from "node:test";

import { bindExprColumns, evaluateExpr, parseFormula } from "../../src/expr/index.js";

test("expr security: rejects identifier-based sandbox escapes (values.constructor)", () => {
  assert.throws(
    () => parseFormula(`=values["con"+"structor"]["con"+"structor"]("return 7")`),
    /Unsupported identifier 'values'/,
  );
});

test("expr security: rejects property access attempts (.[constructor])", () => {
  assert.throws(() => parseFormula(`=[a].constructor`), /Unsupported character '\.'/);
});

test("expr security: rejects access to globals (globalThis)", () => {
  assert.throws(() => parseFormula(`=globalThis`), /Unsupported identifier 'globalThis'/);
});

test("expr security: treats 'constructor' as a column name when bracketed", () => {
  const expr = parseFormula("=[constructor] + 1");
  const bound = bindExprColumns(expr, (name) => {
    if (name === "constructor") return 0;
    throw new Error(`Unknown column ${name}`);
  });
  assert.equal(evaluateExpr(bound, [123]), 124);
});

test("expr security: rejects function calls (Function)", () => {
  assert.throws(() => parseFormula(`=Function("return 1")`), /Unsupported function 'Function'/);
});

test("expr security: rejects unknown whitelisted function names", () => {
  assert.throws(() => parseFormula(`=text_reverse("x")`), /Unsupported function 'text_reverse'/);
});
