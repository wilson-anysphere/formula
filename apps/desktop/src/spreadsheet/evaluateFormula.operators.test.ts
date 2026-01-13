import { describe, expect, it } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula operators", () => {
  it("supports comparisons", () => {
    expect(evaluateFormula("=1>0", () => null)).toBe(true);
    expect(evaluateFormula("=1<0", () => null)).toBe(false);
    expect(evaluateFormula("=1=1", () => null)).toBe(true);
    expect(evaluateFormula("=1<>2", () => null)).toBe(true);
    expect(evaluateFormula('="a"="A"', () => null)).toBe(true);
  });

  it("supports string concatenation (&) with correct precedence", () => {
    expect(evaluateFormula('="a"&"b"', () => null)).toBe("ab");
    // Addition binds tighter than concatenation (Excel precedence).
    expect(evaluateFormula('="a"&1+1', () => null)).toBe("a2");
  });

  it("supports logical functions (AND/OR/NOT/IFERROR)", () => {
    expect(evaluateFormula("=AND(1>0, 2>0)", () => null)).toBe(true);
    expect(evaluateFormula("=AND(1>0, 2<0)", () => null)).toBe(false);
    expect(evaluateFormula("=OR(1>0, 2<0)", () => null)).toBe(true);
    expect(evaluateFormula("=NOT(1>0)", () => null)).toBe(false);
    expect(evaluateFormula('=IFERROR(#REF!, "fallback")', () => null)).toBe("fallback");
  });

  it("treats missing operands / trailing tokens as #VALUE!", () => {
    expect(evaluateFormula("=1+", () => null)).toBe("#VALUE!");
    expect(evaluateFormula("=1> ", () => null)).toBe("#VALUE!");
    expect(evaluateFormula("=1 2", () => null)).toBe("#VALUE!");
  });

  it("accepts semicolons as function argument separators", () => {
    expect(evaluateFormula("=SUM(1;2)", () => null)).toBe(3);
    expect(evaluateFormula("=IF(1>0; TRUE; FALSE)", () => null)).toBe(true);
  });
});
