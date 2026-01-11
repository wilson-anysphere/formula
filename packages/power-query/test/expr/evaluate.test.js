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

test("expr evaluate: text_* functions", () => {
  assert.equal(evalWithColumns('text_upper([A])', ["a"], { A: 0 }), "A");
  assert.equal(evalWithColumns('text_lower([A])', ["A"], { A: 0 }), "a");
  assert.equal(evalWithColumns('text_trim([A])', ["  a  "], { A: 0 }), "a");
  assert.equal(evalWithColumns('text_length([A])', ["abcd"], { A: 0 }), 4);
});

test("expr evaluate: text_contains is case-insensitive", () => {
  assert.equal(evalWithColumns('text_contains([A], "bar")', ["FooBar"], { A: 0 }), true);
  assert.equal(evalWithColumns('text_contains([A], "baz")', ["FooBar"], { A: 0 }), false);
});

test("expr evaluate: number_round()", () => {
  assert.equal(evaluateExpr(parseFormula("number_round(12.345, 1)"), []), 12.3);
  assert.equal(evaluateExpr(parseFormula("number_round(12.345)"), []), 12);
  assert.equal(evaluateExpr(parseFormula("number_round(1234, -2)"), []), 1200);
});

test("expr evaluate: date_from_text() + date_add_days()", () => {
  const expr = parseFormula('date_add_days(date_from_text("2020-01-01"), 2)');
  const value = evaluateExpr(expr, []);
  assert.ok(value instanceof Date);
  assert.equal(value.toISOString(), "2020-01-03T00:00:00.000Z");
});
