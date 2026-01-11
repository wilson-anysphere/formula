import { describe, expect, it } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula error handling", () => {
  it("treats unknown #prefixed strings as text, not spreadsheet errors", () => {
    const getCellValue = (addr: string) => (addr === "A1" ? "#hashtag" : null);
    expect(evaluateFormula("=SUM(A1)", getCellValue, { cellAddress: "Sheet1!B1" })).toBe(0);
  });
});

