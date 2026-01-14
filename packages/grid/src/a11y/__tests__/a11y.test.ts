import { describe, expect, it } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { describeActiveCellLabel, describeCellForA11y, toA1Address, toColumnName } from "../a11y";

describe("a11y helpers", () => {
  it("converts 0-based columns to Excel names", () => {
    expect(toColumnName(0)).toBe("A");
    expect(toColumnName(25)).toBe("Z");
    expect(toColumnName(26)).toBe("AA");
    expect(toColumnName(27)).toBe("AB");
    expect(toColumnName(51)).toBe("AZ");
    expect(toColumnName(52)).toBe("BA");
    expect(toColumnName(701)).toBe("ZZ");
    expect(toColumnName(702)).toBe("AAA");
  });

  it("converts 0-based coordinates to A1 addresses", () => {
    expect(toA1Address(0, 0)).toBe("A1");
    expect(toA1Address(0, 25)).toBe("Z1");
    expect(toA1Address(0, 26)).toBe("AA1");
    expect(toA1Address(0, 27)).toBe("AB1");
    expect(toA1Address(9, 27)).toBe("AB10");
  });

  it("describes no selection", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: null })
    };

    expect(describeCellForA11y({ selection: null, range: null, provider, headerRows: 0, headerCols: 0 })).toBe("No cell selected.");
  });

  it("describes a selected cell with headers configured", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 1 && col === 1 ? "hello" : null })
    };

    expect(describeCellForA11y({ selection: { row: 1, col: 1 }, range: null, provider, headerRows: 1, headerCols: 1 })).toBe(
      "Active cell A1, value hello. Selection none."
    );
    expect(describeActiveCellLabel({ row: 1, col: 1 }, provider, 1, 1)).toBe("Cell A1, value hello.");
  });

  it("formats a single-cell selection range as A1 (not A1:A1)", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 1 && col === 1 ? "hello" : null })
    };

    expect(
      describeCellForA11y({
        selection: { row: 1, col: 1 },
        range: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
        provider,
        headerRows: 1,
        headerCols: 1
      })
    ).toBe("Active cell A1, value hello. Selection A1.");
  });

  it("describes header cells by row/column indices when outside the data region", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 0 && col === 1 ? "Header" : null })
    };

    expect(
      describeCellForA11y({
        selection: { row: 0, col: 1 },
        range: { startRow: 0, endRow: 1, startCol: 1, endCol: 2 },
        provider,
        headerRows: 1,
        headerCols: 1
      })
    ).toBe("Active cell row 1, column 2, value Header. Selection row 1, column 2.");
  });

  it("describes a range selection as A1:B2", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 1 && col === 1 ? "hello" : null })
    };

    expect(
      describeCellForA11y({
        selection: { row: 1, col: 1 },
        range: { startRow: 1, endRow: 3, startCol: 1, endCol: 3 },
        provider,
        headerRows: 1,
        headerCols: 1
      })
    ).toBe("Active cell A1, value hello. Selection A1:B2.");
  });

  it("uses image alt text when the cell value is blank", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 1 && col === 1
          ? { row, col, value: null, image: { imageId: "img1", altText: "Logo" } }
          : { row, col, value: null }
    };

    expect(describeCellForA11y({ selection: { row: 1, col: 1 }, range: null, provider, headerRows: 1, headerCols: 1 })).toBe(
      "Active cell A1, value Logo. Selection none."
    );
  });

  it("trims image alt text when the cell value is blank", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 1 && col === 1
          ? { row, col, value: null, image: { imageId: "img1", altText: "  Logo  " } }
          : { row, col, value: null }
    };

    expect(describeCellForA11y({ selection: { row: 1, col: 1 }, range: null, provider, headerRows: 1, headerCols: 1 })).toBe(
      "Active cell A1, value Logo. Selection none."
    );
    expect(describeActiveCellLabel({ row: 1, col: 1 }, provider, 1, 1)).toBe("Cell A1, value Logo.");
  });

  it("falls back to [Image] when the cell value is blank and no alt text is present", () => {
    const provider: CellProvider = {
      getCell: (row, col) => (row === 1 && col === 1 ? { row, col, value: null, image: { imageId: "img1" } } : { row, col, value: null })
    };

    expect(describeCellForA11y({ selection: { row: 1, col: 1 }, range: null, provider, headerRows: 1, headerCols: 1 })).toBe(
      "Active cell A1, value [Image]. Selection none."
    );
  });

  it("includes image alt text in the active-cell label when the cell value is blank", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 1 && col === 1
          ? { row, col, value: null, image: { imageId: "img1", altText: "Logo" } }
          : { row, col, value: null }
    };

    expect(describeActiveCellLabel({ row: 1, col: 1 }, provider, 1, 1)).toBe("Cell A1, value Logo.");
  });

  it("falls back to [Image] in the active-cell label when the cell value is blank and no alt text is present", () => {
    const provider: CellProvider = {
      getCell: (row, col) => (row === 1 && col === 1 ? { row, col, value: null, image: { imageId: "img1" } } : { row, col, value: null })
    };

    expect(describeActiveCellLabel({ row: 1, col: 1 }, provider, 1, 1)).toBe("Cell A1, value [Image].");
  });
});
