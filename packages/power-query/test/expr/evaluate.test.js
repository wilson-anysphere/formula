import assert from "node:assert/strict";
import test from "node:test";

import { bindExprColumns, evaluateExpr, parseFormula } from "../../src/expr/index.js";

/**
 * @param {string} formula
 * @param {unknown[]} values
 * @param {Record<string, number>} columns
 */
function evalWithColumns(formula, values, columns) {
  const expr = parseFormula(formula);
  const bound = bindExprColumns(expr, (name) => {
    const idx = columns[name];
    if (idx == null) throw new Error(`Unknown column ${name}`);
    return idx;
  });
  return evaluateExpr(bound, values);
}

test("expr evaluate: arithmetic precedence", () => {
  assert.equal(evalWithColumns("=[A] + [B] * 2", [1, 2], { A: 0, B: 1 }), 5);
  assert.equal(evalWithColumns("=([A] + [B]) * 2", [1, 2], { A: 0, B: 1 }), 6);
});

test("expr evaluate: comparisons + ternary", () => {
  assert.equal(evalWithColumns('=[A] > 0 ? "pos" : "neg"', [1], { A: 0 }), "pos");
  assert.equal(evalWithColumns('=[A] > 0 ? "pos" : "neg"', [-1], { A: 0 }), "neg");
});

test("expr evaluate: exponent number literals", () => {
  const expr = parseFormula("=1e3 + 5");
  assert.equal(evaluateExpr(expr, []), 1005);
});

test("expr evaluate: string escapes", () => {
  const value = "a\nb\\c";
  const expr = parseFormula(JSON.stringify(value));
  assert.equal(evaluateExpr(expr, []), value);
});

test("expr evaluate: value placeholder _", () => {
  const expr = parseFormula("_ == null ? 0 : _");
  assert.equal(evaluateExpr(expr, [], null, null), 0);
  assert.equal(evaluateExpr(expr, [], null, 5), 5);
  assert.throws(() => evaluateExpr(expr, []), /Formula references '_' but no value was provided/);
});

test("expr evaluate: date() literal", () => {
  const expr = parseFormula('date("2020-01-01")');
  const value = evaluateExpr(expr, []);
  assert.ok(value instanceof Date);
  assert.equal(value.toISOString(), "2020-01-01T00:00:00.000Z");
});

test("expr evaluate: date equality compares by timestamp (not identity)", () => {
  assert.equal(evaluateExpr(parseFormula('date("2020-01-01") == date("2020-01-01")'), []), true);
  assert.equal(evaluateExpr(parseFormula('date("2020-01-01") != date("2020-01-01")'), []), false);
  assert.equal(evaluateExpr(parseFormula('date("2020-01-01") === date("2020-01-01")'), []), true);
  assert.equal(evaluateExpr(parseFormula('date("2020-01-01") !== date("2020-01-01")'), []), false);
});

test("expr evaluate: date() rejects invalid formats", () => {
  const expr = parseFormula('date("2020-02-30")');
  assert.throws(() => evaluateExpr(expr, []), /Invalid date literal/);
});
