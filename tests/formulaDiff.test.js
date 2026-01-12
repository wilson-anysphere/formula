import test from "node:test";
import assert from "node:assert/strict";

import { diffFormula } from "../packages/versioning/src/formula/diff.js";

/**
 * @param {ReturnType<typeof diffFormula>["ops"]} ops
 */
function simplifyOps(ops) {
  return ops.map((op) => ({
    type: op.type,
    tokens: op.tokens.map((t) => `${t.type}:${t.value}`),
  }));
}

test("diffFormula: equal formulas with whitespace differences (normalize=true)", () => {
  const result = diffFormula("=SUM(A1:A10)", "=SUM( A1 : A10 )", { normalize: true });
  assert.equal(result.equal, true);
  assert.deepEqual(simplifyOps(result.ops), [
    {
      type: "equal",
      tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::", "ident:A10", "punct:)"],
    },
  ]);
});

test("diffFormula: single token replacement (A10 -> A12)", () => {
  const result = diffFormula("=SUM(A1:A10)", "=SUM(A1:A12)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    {
      type: "equal",
      tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::"],
    },
    { type: "delete", tokens: ["ident:A10"] },
    { type: "insert", tokens: ["ident:A12"] },
    { type: "equal", tokens: ["punct:)"] },
  ]);
});

test("diffFormula: function name change (SUM -> AVERAGE)", () => {
  const result = diffFormula("=SUM(A1:A10)", "=AVERAGE(A1:A10)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:="] },
    { type: "delete", tokens: ["ident:SUM"] },
    { type: "insert", tokens: ["ident:AVERAGE"] },
    { type: "equal", tokens: ["punct:(", "ident:A1", "op::", "ident:A10", "punct:)"] },
  ]);
});

test("diffFormula: insertion around ranges/punctuation", () => {
  const result = diffFormula("=SUM(A1:A10)", "=SUM(A1:A10,B1:B2)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::", "ident:A10"] },
    { type: "insert", tokens: ["punct:,", "ident:B1", "op::", "ident:B2"] },
    { type: "equal", tokens: ["punct:)"] },
  ]);
});

test("diffFormula: deletion around ranges/punctuation", () => {
  const result = diffFormula("=SUM(A1:A10,B1:B2)", "=SUM(A1:A10)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::", "ident:A10"] },
    { type: "delete", tokens: ["punct:,", "ident:B1", "op::", "ident:B2"] },
    { type: "equal", tokens: ["punct:)"] },
  ]);
});

test("diffFormula: null/empty handling", () => {
  assert.deepEqual(diffFormula(null, null), { equal: true, ops: [] });
  assert.deepEqual(diffFormula("", "   "), { equal: true, ops: [] });

  assert.deepEqual(simplifyOps(diffFormula(null, "=1+2").ops), [
    { type: "insert", tokens: ["op:=", "number:1", "op:+", "number:2"] },
  ]);

  assert.deepEqual(simplifyOps(diffFormula("=1+2", null).ops), [
    { type: "delete", tokens: ["op:=", "number:1", "op:+", "number:2"] },
  ]);
});

test("diffFormula: normalization ensures leading '=' by default", () => {
  const result = diffFormula("SUM(A1:A10)", "=SUM(A1:A10)");
  assert.equal(result.equal, true);
  assert.deepEqual(simplifyOps(result.ops), [
    {
      type: "equal",
      tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::", "ident:A10", "punct:)"],
    },
  ]);
});

