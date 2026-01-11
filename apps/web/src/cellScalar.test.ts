import { describe, expect, it } from "vitest";

import { isFormulaInput, parseCellScalarInput } from "./cellScalar";

describe("cellScalar helpers (web)", () => {
  describe("parseCellScalarInput", () => {
    it("returns null for empty/whitespace-only input", () => {
      expect(parseCellScalarInput("")).toBeNull();
      expect(parseCellScalarInput("   ")).toBeNull();
      expect(parseCellScalarInput("\n\t")).toBeNull();
    });

    it("treats bare '=' (and whitespace around it) as empty", () => {
      expect(parseCellScalarInput("=")).toBeNull();
      expect(parseCellScalarInput("   =   ")).toBeNull();
    });

    it("normalizes formulas to trimmed display form", () => {
      expect(parseCellScalarInput("=1+1")).toBe("=1+1");
      expect(parseCellScalarInput("  =  SUM(A1:A3)  ")).toBe("=SUM(A1:A3)");
      expect(parseCellScalarInput("   =1+1")).toBe("=1+1");
    });

    it("only strips a single leading '=' when normalizing", () => {
      expect(parseCellScalarInput("==1+1")).toBe("==1+1");
    });
  });

  describe("isFormulaInput", () => {
    it("treats strings starting with '=' as formulas", () => {
      expect(isFormulaInput("=1+1")).toBe(true);
      expect(isFormulaInput("==1+1")).toBe(true);
    });

    it("does not treat non-strings as formulas", () => {
      expect(isFormulaInput(null)).toBe(false);
      expect(isFormulaInput(123)).toBe(false);
      expect(isFormulaInput(true)).toBe(false);
    });
  });
});

