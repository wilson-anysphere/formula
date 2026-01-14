import { describe, expect, it, vi } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula structured reference resolver", () => {
  it("evaluates structured references via resolveStructuredRefToReference", () => {
    const getCellValue = (addr: string) => {
      if (addr === "A2") return 1;
      if (addr === "A3") return 2;
      if (addr === "A4") return 3;
      return null;
    };

    const resolver = vi.fn((refText: string) => (refText === "Table1[Amount]" ? "A2:A4" : null));

    const value = evaluateFormula("=SUM(Table1[Amount])", getCellValue, {
      resolveStructuredRefToReference: resolver,
    });

    expect(value).toBe(6);
    expect(resolver).toHaveBeenCalledWith("Table1[Amount]");
  });

  it("supports sheet-qualified replacements returned by resolveStructuredRefToReference", () => {
    const getCellValue = (addr: string) => {
      if (addr === "Sheet1!A2") return 1;
      if (addr === "Sheet1!A3") return 2;
      return null;
    };

    const resolver = vi.fn((refText: string) => (refText === "Table1[#Data]" ? "Sheet1!A2:A3" : null));

    const value = evaluateFormula("=SUM(Table1[#Data])", getCellValue, {
      resolveStructuredRefToReference: resolver,
    });

    expect(value).toBe(3);
    expect(resolver).toHaveBeenCalledWith("Table1[#Data]");
  });
});

