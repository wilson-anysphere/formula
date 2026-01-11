import test from "node:test";
import assert from "node:assert/strict";

import { resolveFormulaConflict } from "../packages/collab/conflicts/src/formula-conflict-resolver.js";

test("resolveFormulaConflict: AST-equivalent formulas auto-resolve", () => {
  const decision = resolveFormulaConflict({
    localFormula: "=SUM(A1:A2)",
    remoteFormula: "=sum(a1:a2)"
  });

  assert.equal(decision.kind, "equivalent");
  assert.equal(decision.chosenFormula, "=sum(a1:a2)");
});

test("resolveFormulaConflict: prefers the extension (remote)", () => {
  const decision = resolveFormulaConflict({
    localFormula: "=A1+1",
    remoteFormula: "=A1+1+1"
  });

  assert.equal(decision.kind, "prefer-remote");
  assert.equal(decision.chosenFormula, "=A1+1+1");
});

test("resolveFormulaConflict: prefers the extension (local)", () => {
  const decision = resolveFormulaConflict({
    localFormula: "=A1+1+1",
    remoteFormula: "=A1+1"
  });

  assert.equal(decision.kind, "prefer-local");
  assert.equal(decision.chosenFormula, "=A1+1+1");
});

test("resolveFormulaConflict: surfaces true conflicts", () => {
  const decision = resolveFormulaConflict({
    localFormula: "=A1+1",
    remoteFormula: "=A1*2"
  });

  assert.equal(decision.kind, "conflict");
});

test("resolveFormulaConflict: empty vs non-empty formulas surface a conflict", () => {
  const decision = resolveFormulaConflict({
    localFormula: "",
    remoteFormula: "=A1"
  });

  assert.equal(decision.kind, "conflict");
});
