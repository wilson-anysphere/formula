import { describe, expect, it } from "vitest";

import { getFunctionCallContext, getFunctionHint } from "./functionContext.js";

describe("function context", () => {
  it("getFunctionCallContext returns innermost function + arg index", () => {
    const formula = '=IF(SUM(A1:A2) > 3, "yes", "no")';

    const insideSum = formula.indexOf("A1");
    const sumContext = getFunctionCallContext(formula, insideSum);
    expect(sumContext).toEqual({ name: "SUM", argIndex: 0 });

    const insideIfSecondArg = formula.indexOf('"yes"');
    const ifContext = getFunctionCallContext(formula, insideIfSecondArg);
    expect(ifContext).toEqual({ name: "IF", argIndex: 1 });
  });

  it("getFunctionHint uses signature mapping and marks active parameter", () => {
    const formula = "=IF(A1, B1, C1)";
    const cursor = formula.indexOf("B1") + 1;
    const hint = getFunctionHint(formula, cursor);
    expect(hint).toBeTruthy();
    expect(hint?.signature.name).toBe("IF");
    expect(hint?.context.argIndex).toBe(1);
    expect(hint?.parts.some((p) => p.kind === "paramActive")).toBe(true);
  });

  it("getFunctionHint prefers curated signatures when available (XLOOKUP)", () => {
    const formula = "=XLOOKUP(A1, B1, C1)";
    const cursor = formula.indexOf("B1") + 1;
    const hint = getFunctionHint(formula, cursor);
    expect(hint).toBeTruthy();
    expect(hint?.signature.name).toBe("XLOOKUP");
    expect(hint?.signature.params[0]?.name).toBe("lookup_value");
    expect(hint?.signature.summary).toContain("Looks up");
  });

  it("getFunctionHint prefers curated signatures when available (SEQUENCE)", () => {
    const formula = "=SEQUENCE(10, 2)";
    const cursor = formula.indexOf("2") + 1;
    const hint = getFunctionHint(formula, cursor);
    expect(hint).toBeTruthy();
    expect(hint?.signature.name).toBe("SEQUENCE");
    expect(hint?.signature.params[0]?.name).toBe("rows");
    expect(hint?.signature.params[1]?.name).toBe("columns");
  });
});
