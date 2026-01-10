import test from "node:test";
import assert from "node:assert/strict";

import { buildStepTree, DebuggerState, flattenVisibleSteps } from "./steps.ts";
import { highlightsForStep } from "./highlight.ts";
import { explainError } from "./errorExplanation.ts";
import type { TraceNode } from "./types.ts";

test("e2e: open formula debugger for VLOOKUP and inspect steps + hover highlights", () => {
  const formula = "=VLOOKUP(A1,B1:C2,2,FALSE)";

  const spanA1Start = formula.indexOf("A1");
  const spanRangeStart = formula.indexOf("B1:C2");
  const spanFalseStart = formula.indexOf("FALSE");
  const spanColStart = formula.indexOf(",2,") + 1;

  const trace: TraceNode = {
    kind: { type: "function_call", name: "VLOOKUP" },
    span: { start: 1, end: formula.length },
    value: 19.99,
    children: [
      {
        kind: { type: "cell_ref" },
        span: { start: spanA1Start, end: spanA1Start + 2 },
        value: "Key-123",
        reference: { type: "cell", cell: "A1" },
        children: [],
      },
      {
        kind: { type: "range_ref" },
        span: { start: spanRangeStart, end: spanRangeStart + "B1:C2".length },
        value: { array: [["Key-123", 19.99], ["Key-456", 29.99]] },
        reference: { type: "range", range: "B1:C2" },
        children: [],
      },
      {
        kind: { type: "number" },
        span: { start: spanColStart, end: spanColStart + 1 },
        value: 2,
        children: [],
      },
      {
        kind: { type: "bool" },
        span: { start: spanFalseStart, end: spanFalseStart + "FALSE".length },
        value: false,
        children: [],
      },
    ],
  };

  const root = buildStepTree(formula, trace);
  assert.equal(root.text, "VLOOKUP(A1,B1:C2,2,FALSE)");
  assert.equal(root.children.length, 4);

  const state = new DebuggerState();
  let visible = flattenVisibleSteps(root, state.collapsed);
  assert.ok(visible.length > 1);

  state.toggle("0"); // collapse root: hides sub-steps.
  visible = flattenVisibleSteps(root, state.collapsed);
  assert.equal(visible.length, 1);

  state.toggle("0");
  const rangeStep = root.children[1];
  const highlights = highlightsForStep(rangeStep);
  assert.deepEqual([...highlights].sort(), ["B1", "B2", "C1", "C2"].sort());
});

test("error explanation uses trace context when available", () => {
  const formula = "=VLOOKUP(A1,B1:C2,2,FALSE)";
  const trace: TraceNode = {
    kind: { type: "function_call", name: "VLOOKUP" },
    span: { start: 1, end: formula.length },
    value: { error: "#N/A" },
    children: [
      {
        kind: { type: "cell_ref" },
        span: { start: formula.indexOf("A1"), end: formula.indexOf("A1") + 2 },
        value: "Missing",
        reference: { type: "cell", cell: "A1" },
        children: [],
      },
      {
        kind: { type: "range_ref" },
        span: { start: formula.indexOf("B1:C2"), end: formula.indexOf("B1:C2") + "B1:C2".length },
        value: { array: [["Key-123", 19.99], ["Key-456", 29.99]] },
        reference: { type: "range", range: "B1:C2" },
        children: [],
      },
      { kind: { type: "number" }, span: { start: 0, end: 0 }, value: 2, children: [] },
      { kind: { type: "bool" }, span: { start: 0, end: 0 }, value: false, children: [] },
    ],
  };

  const explanation = explainError(formula, trace);
  assert.ok(explanation);
  assert.match(explanation.problem, /not found/i);
  assert.match(explanation.problem, /B1:C2/);
});

