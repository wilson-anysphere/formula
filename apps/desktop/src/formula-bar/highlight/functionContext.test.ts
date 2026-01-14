import { describe, expect, it } from "vitest";

import { getFunctionCallContext } from "./functionContext.js";
import { getFunctionSignature, signatureParts } from "./functionSignatures.js";

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

  it("getFunctionCallContext ignores commas inside structured references", () => {
    const formula = "=SUM(Table1[[#All],[Amount]], 1)";
    const cursor = formula.indexOf("Amount") + 1;
    const context = getFunctionCallContext(formula, cursor);
    expect(context).toEqual({ name: "SUM", argIndex: 0 });
  });

  it("getFunctionCallContext ignores commas inside array literals", () => {
    const formula = "=SUM({1,2,3}, 4)";
    const cursor = formula.indexOf("2");
    const context = getFunctionCallContext(formula, cursor);
    expect(context).toEqual({ name: "SUM", argIndex: 0 });
  });

  it("getFunctionCallContext supports semicolon argument separators", () => {
    const formula = "=IF(A1>0;1;2)";
    const insideTrue = formula.indexOf(";1;") + 1;
    const trueContext = getFunctionCallContext(formula, insideTrue);
    expect(trueContext).toEqual({ name: "IF", argIndex: 1 });

    const insideFalse = formula.indexOf(";2)") + 1;
    const falseContext = getFunctionCallContext(formula, insideFalse);
    expect(falseContext).toEqual({ name: "IF", argIndex: 2 });
  });

  it("signatureParts marks the active parameter (via getFunctionCallContext)", () => {
    const formula = "=IF(A1, B1, C1)";
    const cursor = formula.indexOf("B1") + 1;
    const ctx = getFunctionCallContext(formula, cursor);
    expect(ctx).toEqual({ name: "IF", argIndex: 1 });
    const sig = getFunctionSignature(ctx?.name ?? "");
    expect(sig).toBeTruthy();
    const parts = sig ? signatureParts(sig, ctx?.argIndex ?? 0) : [];
    expect(sig?.name).toBe("IF");
    expect(parts.some((p) => p.kind === "paramActive")).toBe(true);
  });

  it("getFunctionSignature prefers curated signatures when available (XLOOKUP)", () => {
    const formula = "=XLOOKUP(A1, B1, C1)";
    const cursor = formula.indexOf("B1") + 1;
    const ctx = getFunctionCallContext(formula, cursor);
    expect(ctx).toEqual({ name: "XLOOKUP", argIndex: 1 });
    const sig = getFunctionSignature(ctx?.name ?? "");
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("XLOOKUP");
    expect(sig?.params[0]?.name).toBe("lookup_value");
    expect(sig?.summary).toContain("Looks up");
  });

  it("getFunctionSignature prefers curated signatures when available (SEQUENCE)", () => {
    const formula = "=SEQUENCE(10, 2)";
    const cursor = formula.indexOf("2") + 1;
    const ctx = getFunctionCallContext(formula, cursor);
    expect(ctx).toEqual({ name: "SEQUENCE", argIndex: 1 });
    const sig = getFunctionSignature(ctx?.name ?? "");
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("SEQUENCE");
    expect(sig?.params[0]?.name).toBe("rows");
    expect(sig?.params[1]?.name).toBe("columns");
  });
});
