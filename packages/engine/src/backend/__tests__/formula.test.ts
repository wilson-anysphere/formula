import { describe, expect, it } from "vitest";

import { isFormulaInput, normalizeFormulaText, normalizeFormulaTextOpt } from "../formula.ts";

describe("formula helpers", () => {
  describe("normalizeFormulaText", () => {
    it("returns an empty string for empty or whitespace-only input", () => {
      expect(normalizeFormulaText("")).toBe("");
      expect(normalizeFormulaText("   ")).toBe("");
      expect(normalizeFormulaText("\n\t")).toBe("");
    });

    it("treats bare '=' (and whitespace around it) as empty", () => {
      expect(normalizeFormulaText("=")).toBe("");
      expect(normalizeFormulaText("   =   ")).toBe("");
    });

    it("ensures a leading '=' and trims whitespace", () => {
      expect(normalizeFormulaText("=1+1")).toBe("=1+1");
      expect(normalizeFormulaText("  =  SUM(A1:A3)  ")).toBe("=SUM(A1:A3)");
    });

    it("only strips a single leading '='", () => {
      expect(normalizeFormulaText("==1+1")).toBe("==1+1");
    });
  });

  describe("normalizeFormulaTextOpt", () => {
    it("returns null for empty formulas", () => {
      expect(normalizeFormulaTextOpt("")).toBeNull();
      expect(normalizeFormulaTextOpt("   ")).toBeNull();
      expect(normalizeFormulaTextOpt("=")).toBeNull();
      expect(normalizeFormulaTextOpt("   =   ")).toBeNull();
    });

    it("returns the normalized display form for non-empty formulas", () => {
      expect(normalizeFormulaTextOpt("1+1")).toBe("=1+1");
      expect(normalizeFormulaTextOpt("  =  SUM(A1:A3)  ")).toBe("=SUM(A1:A3)");
    });
  });

  describe("isFormulaInput", () => {
    it("treats leading whitespace + '=' as a formula indicator", () => {
      expect(isFormulaInput("=1+1")).toBe(true);
      expect(isFormulaInput("   =1+1")).toBe(true);
    });

    it("does not treat a bare '=' as formula input", () => {
      expect(isFormulaInput("=")).toBe(false);
      expect(isFormulaInput("   =   ")).toBe(false);
    });

    it("does not treat arbitrary strings as formulas", () => {
      expect(isFormulaInput("1+1")).toBe(false);
      expect(isFormulaInput("SUM(A1:A3)")).toBe(false);
    });
  });
});
