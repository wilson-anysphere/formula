import { describe, expect, it } from "vitest";
import { computeFillEdits, shiftFormulaA1, type CellRange, type FillSourceCell } from "../src/index";

describe("shiftFormulaA1", () => {
  it("shifts relative refs and preserves absolute refs", () => {
    expect(shiftFormulaA1("=A1+$B$1", 1, 0)).toBe("=A2+$B$1");
    expect(shiftFormulaA1("=$A1+B$1", 2, 3)).toBe("=$A3+E$1");
  });
});

describe("computeFillEdits", () => {
  it("extends numeric series vertically", () => {
    const sourceRange: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    const targetRange: CellRange = { startRow: 2, endRow: 4, startCol: 0, endCol: 1 };
    const sourceCells: FillSourceCell[][] = [
      [{ input: 1, value: 1 }],
      [{ input: 2, value: 2 }]
    ];

    const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode: "series" });
    expect(edits).toEqual([
      { row: 2, col: 0, value: 3 },
      { row: 3, col: 0, value: 4 }
    ]);
  });

  it("fills formulas with relative reference adjustment (pattern repetition)", () => {
    const sourceRange: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    const targetRange: CellRange = { startRow: 2, endRow: 4, startCol: 0, endCol: 1 };
    const sourceCells: FillSourceCell[][] = [
      [{ input: "=B1", value: null }],
      [{ input: "=B2", value: null }]
    ];

    const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode: "formulas" });
    expect(edits).toEqual([
      { row: 2, col: 0, value: "=B3" },
      { row: 3, col: 0, value: "=B4" }
    ]);
  });

  it("extends month names horizontally", () => {
    const sourceRange: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 2 };
    const targetRange: CellRange = { startRow: 0, endRow: 1, startCol: 2, endCol: 4 };
    const sourceCells: FillSourceCell[][] = [[{ input: "Jan", value: "Jan" }, { input: "Feb", value: "Feb" }]];

    const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode: "series" });
    expect(edits).toEqual([
      { row: 0, col: 2, value: "Mar" },
      { row: 0, col: 3, value: "Apr" }
    ]);
  });
});

