import { describe, expect, it } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula error handling", () => {
  it("treats unknown #prefixed strings as text, not spreadsheet errors", () => {
    const getCellValue = (addr: string) => (addr === "A1" ? "#hashtag" : null);
    expect(evaluateFormula("=SUM(A1)", getCellValue, { cellAddress: "Sheet1!B1" })).toBe(0);
  });

  it("fails fast when asked to materialize an oversized range", () => {
    let reads = 0;
    const getCellValue = (_addr: string) => {
      reads += 1;
      return 1;
    };

    // 2 columns x 6 rows = 12 cells. With a tiny cap, we should bail out before scanning.
    expect(evaluateFormula("=SUM(A1:B6)", getCellValue, { maxRangeCells: 10 })).toBe("#VALUE!");
    expect(reads).toBe(0);
  });
});
