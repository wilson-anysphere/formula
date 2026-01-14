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

  it("canonicalizes localized function names when a localeId is provided", () => {
    expect(evaluateFormula("=SUMME(1;2)", () => null, { localeId: "de-DE" })).toBe(3);
    expect(evaluateFormula("=MITTELWERT(1;2;3)", () => null, { localeId: "de-DE" })).toBe(2);
    expect(evaluateFormula("=WENN(1>0; TRUE; FALSE)", () => null, { localeId: "de-DE" })).toBe(true);

    expect(evaluateFormula("=SOMME(1;2)", () => null, { localeId: "fr-FR" })).toBe(3);
    expect(evaluateFormula("=MOYENNE(1;2;3)", () => null, { localeId: "fr-FR" })).toBe(2);

    expect(evaluateFormula("=SUMA(1;2)", () => null, { localeId: "es-ES" })).toBe(3);
    expect(evaluateFormula("=PROMEDIO(1;2;3)", () => null, { localeId: "es-ES" })).toBe(2);
  });

  it("parses decimal commas and thousands separators when a comma-decimal localeId is provided", () => {
    // de-DE uses `,` decimals + `.` thousands separators.
    expect(evaluateFormula("=1,5+2,5", () => null, { localeId: "de-DE" })).toBe(4);
    expect(evaluateFormula("=SUMME(1,5;2,5)", () => null, { localeId: "de-DE" })).toBe(4);
    expect(evaluateFormula("=1.234,5+0,5", () => null, { localeId: "de-DE" })).toBe(1235);
    expect(evaluateFormula("=1.234.567,5+0,5", () => null, { localeId: "de-DE" })).toBe(1234568);

    // fr-FR uses `,` decimals (thousands grouping is NBSP; we don't require it here).
    expect(evaluateFormula("=1,5+2,5", () => null, { localeId: "fr-FR" })).toBe(4);
    expect(evaluateFormula("=SOMME(1,5;2,5)", () => null, { localeId: "fr-FR" })).toBe(4);

    // es-ES uses `,` decimals.
    expect(evaluateFormula("=1,5+2,5", () => null, { localeId: "es-ES" })).toBe(4);
    expect(evaluateFormula("=SUMA(1,5;2,5)", () => null, { localeId: "es-ES" })).toBe(4);
  });
});
