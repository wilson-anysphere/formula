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
  // Excel-style input handling treats a bare "=" as an empty formula.
  assert.deepEqual(diffFormula("=", null), { equal: true, ops: [] });

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

test("diffFormula: supports absolute references ($A$10 -> $A$12)", () => {
  const result = diffFormula("=SUM($A$1:$A$10)", "=SUM($A$1:$A$12)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:SUM", "punct:(", "ident:$A$1", "op::"] },
    { type: "delete", tokens: ["ident:$A$10"] },
    { type: "insert", tokens: ["ident:$A$12"] },
    { type: "equal", tokens: ["punct:)"] },
  ]);
});

test("diffFormula: supports comparison operators (>, >=)", () => {
  const result = diffFormula("=IF(A1>0,1,0)", "=IF(A1>=0,1,0)");
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:IF", "punct:(", "ident:A1"] },
    { type: "delete", tokens: ["op:>"] },
    { type: "insert", tokens: ["op:>="] },
    { type: "equal", tokens: ["number:0", "punct:,", "number:1", "punct:,", "number:0", "punct:)"] },
  ]);
});

test("diffFormula: normalize=true treats identifier case changes as equal", () => {
  const result = diffFormula("=sum(a1:a2)", "=SUM(A1:A2)", { normalize: true });
  assert.equal(result.equal, true);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "op::", "ident:A2", "punct:)"] },
  ]);
});

test("diffFormula: normalize=false treats identifier case changes as edits", () => {
  const result = diffFormula("=sum(a1:a2)", "=SUM(A1:A2)", { normalize: false });
  assert.equal(result.equal, false);
});

test("diffFormula: normalize=true treats comma/semicolon argument separators as equal", () => {
  const result = diffFormula("=SUM(A1;B1)", "=SUM(A1,B1)", { normalize: true });
  assert.equal(result.equal, true);
  // In normalized mode, equal ops should emit tokens from the *new* formula.
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:SUM", "punct:(", "ident:A1", "punct:,", "ident:B1", "punct:)"] },
  ]);
});

test("diffFormula: normalize=false treats comma/semicolon argument separators as edits", () => {
  const result = diffFormula("=SUM(A1;B1)", "=SUM(A1,B1)", { normalize: false });
  assert.equal(result.equal, false);
});

test("diffFormula: comma/semicolon are not interchangeable inside array constants", () => {
  // In array constants (`{...}`), Excel uses `,` as a column separator and `;` as a row separator.
  // Treating them as equivalent here would hide real structural changes.
  const result = diffFormula("={1,2;3,4}", "={1;2;3;4}", { normalize: true });
  assert.equal(result.equal, false);
});

test("diffFormula: normalize=true treats numeric formatting differences as equal", () => {
  const result = diffFormula("=1.0+2", "=1+2", { normalize: true });
  assert.equal(result.equal, true);
  // In normalized mode, equal ops should emit tokens from the *new* formula.
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "number:1", "op:+", "number:2"] },
  ]);
});

test("diffFormula: normalize=false treats numeric formatting differences as edits", () => {
  const result = diffFormula("=1.0+2", "=1+2", { normalize: false });
  assert.equal(result.equal, false);
});

test("diffFormula: string literals remain case-sensitive under normalization", () => {
  const result = diffFormula('=IF(A1="Hello",1,0)', '=IF(A1="hello",1,0)', { normalize: true });
  assert.equal(result.equal, false);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:IF", "punct:(", "ident:A1", "op:="] },
    { type: "delete", tokens: ["string:Hello"] },
    { type: "insert", tokens: ["string:hello"] },
    { type: "equal", tokens: ["punct:,", "number:1", "punct:,", "number:0", "punct:)"] },
  ]);
});

test("diffFormula: normalize=true treats quoted sheet-name case changes as equal", () => {
  const result = diffFormula("='My Sheet'!A1", "='MY SHEET'!A1", { normalize: true });
  assert.equal(result.equal, true);
  assert.deepEqual(simplifyOps(result.ops), [
    { type: "equal", tokens: ["op:=", "ident:MY SHEET", "op:!", "ident:A1"] },
  ]);
});

test("diffFormula: does not throw on unterminated string literals", () => {
  assert.doesNotThrow(() => diffFormula('=IF(A1="Hello', '=IF(A1="Hello!")'));
  const result = diffFormula('=IF(A1="Hello', '=IF(A1="Hello!")');
  assert.equal(result.equal, false);
});

test("diffFormula: long formulas avoid pathological diff work (guardrail)", () => {
  const terms = 1200;
  const oldFormula = `=${Array.from({ length: terms }, () => "A1").join("+")}`;
  const newFormula = `=${Array.from({ length: terms }, () => "B1").join("+")}`;
  const result = diffFormula(oldFormula, newFormula);
  assert.equal(result.equal, false);
  assert.equal(result.ops[0].type, "equal");
  assert.equal(result.ops[1].type, "delete");
  assert.equal(result.ops[2].type, "insert");
  // Ensure we didn't run Myers on the full token stream (which would usually
  // produce many alternating ops due to the repeated "+" tokens). The guardrail
  // fallback should produce a simple equal/delete/insert structure.
  assert.equal(result.ops.length, 3);
});
