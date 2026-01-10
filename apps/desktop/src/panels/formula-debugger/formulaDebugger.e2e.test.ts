import { describe, expect, it } from "vitest";

import { DebuggerState, buildStepTree, flattenVisibleSteps } from "./steps.ts";
import { highlightsForStep } from "./highlight.ts";
import { explainError } from "./errorExplanation.ts";
import type { TraceNode } from "./types.ts";

describe("formula debugger e2e helpers", () => {
  it("opens VLOOKUP trace and inspects steps + hover highlights", () => {
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
    expect(root.text).toBe("VLOOKUP(A1,B1:C2,2,FALSE)");
    expect(root.children).toHaveLength(4);

    const state = new DebuggerState();
    let visible = flattenVisibleSteps(root, state.collapsed);
    expect(visible.length).toBeGreaterThan(1);

    state.toggle("0"); // collapse root: hides sub-steps.
    visible = flattenVisibleSteps(root, state.collapsed);
    expect(visible).toHaveLength(1);

    state.toggle("0");
    const rangeStep = root.children[1];
    const highlights = highlightsForStep(rangeStep);
    expect([...highlights].sort()).toEqual(["B1", "B2", "C1", "C2"].sort());
  });

  it("generates error explanation using trace context when available", () => {
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
          children: []
        },
        {
          kind: { type: "range_ref" },
          span: { start: formula.indexOf("B1:C2"), end: formula.indexOf("B1:C2") + "B1:C2".length },
          value: { array: [["Key-123", 19.99], ["Key-456", 29.99]] },
          reference: { type: "range", range: "B1:C2" },
          children: []
        },
        { kind: { type: "number" }, span: { start: 0, end: 0 }, value: 2, children: [] },
        { kind: { type: "bool" }, span: { start: 0, end: 0 }, value: false, children: [] }
      ]
    };

    const explanation = explainError(formula, trace);
    expect(explanation).toBeTruthy();
    expect(explanation?.problem).toMatch(/not found/i);
    expect(explanation?.problem).toMatch(/B1:C2/);
  });
});
