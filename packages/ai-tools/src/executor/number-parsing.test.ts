import { describe, expect, it } from "vitest";

import { parseSpreadsheetNumber } from "./number-parsing.ts";

describe("parseSpreadsheetNumber", () => {
  it("parses common spreadsheet numeric formats", () => {
    expect(parseSpreadsheetNumber(123)).toBe(123);
    expect(parseSpreadsheetNumber("123")).toBe(123);
    expect(parseSpreadsheetNumber("1,200")).toBe(1200);
    expect(parseSpreadsheetNumber("$5")).toBe(5);
    expect(parseSpreadsheetNumber("$-5")).toBe(-5);
    expect(parseSpreadsheetNumber("(1,200)")).toBe(-1200);
    expect(parseSpreadsheetNumber("(-5)")).toBe(-5);
    expect(parseSpreadsheetNumber("($-5)")).toBe(-5);
    expect(parseSpreadsheetNumber("10%")).toBeCloseTo(0.1);
    expect(parseSpreadsheetNumber("(-10%)")).toBeCloseTo(-0.1);
    expect(parseSpreadsheetNumber(".5")).toBeCloseTo(0.5);
    expect(parseSpreadsheetNumber("1e3")).toBe(1000);
  });

  it("rejects malformed comma groupings and non-numeric strings", () => {
    expect(parseSpreadsheetNumber("1,2")).toBeNull();
    expect(parseSpreadsheetNumber("12,34")).toBeNull();
    expect(parseSpreadsheetNumber("foo")).toBeNull();
    expect(parseSpreadsheetNumber("")).toBeNull();
    expect(parseSpreadsheetNumber("Infinity")).toBeNull();
  });
});
