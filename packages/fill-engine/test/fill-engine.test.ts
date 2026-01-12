import { describe, expect, it } from "vitest";
import { computeFillEdits, shiftFormulaA1, type CellRange, type FillSourceCell } from "../src/index";

describe("shiftFormulaA1", () => {
  it("shifts relative refs and preserves absolute refs", () => {
    expect(shiftFormulaA1("=A1+$B$1", 1, 0)).toBe("=A2+$B$1");
    expect(shiftFormulaA1("=$A1+B$1", 2, 3)).toBe("=$A3+E$1");
  });

  it("drops the spill-range postfix when shifting creates a #REF!", () => {
    expect(shiftFormulaA1("=A1#", 0, -1)).toBe("=#REF!");
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

  it("extends numeric suffix text series vertically", () => {
    const sourceRange: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    const targetRange: CellRange = { startRow: 2, endRow: 4, startCol: 0, endCol: 1 };
    const sourceCells: FillSourceCell[][] = [
      [{ input: "Item 1", value: "Item 1" }],
      [{ input: "Item 3", value: "Item 3" }]
    ];

    const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode: "series" });
    expect(edits).toEqual([
      { row: 2, col: 0, value: "Item 5" },
      { row: 3, col: 0, value: "Item 7" }
    ]);
  });

  it("preserves numeric suffix padding when extending text series vertically", () => {
    const sourceRange: CellRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 };
    const targetRange: CellRange = { startRow: 2, endRow: 4, startCol: 0, endCol: 1 };
    const sourceCells: FillSourceCell[][] = [
      [{ input: "Item 01", value: "Item 01" }],
      [{ input: "Item 03", value: "Item 03" }]
    ];

    const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode: "series" });
    expect(edits).toEqual([
      { row: 2, col: 0, value: "Item 05" },
      { row: 3, col: 0, value: "Item 07" }
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
