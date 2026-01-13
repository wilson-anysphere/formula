import { describe, expect, it } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { describeActiveCellLabel, describeCell, toA1Address, toColumnName } from "../a11y";

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

    expect(describeCell(null, null, provider, 0, 0)).toBe("No cell selected.");
  });

  it("describes a selected cell with headers configured", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 1 && col === 1 ? "hello" : null })
    };

    expect(describeCell({ row: 1, col: 1 }, null, provider, 1, 1)).toBe("Active cell A1, value hello. Selection none.");
    expect(describeActiveCellLabel({ row: 1, col: 1 }, provider, 1, 1)).toBe("Cell A1, value hello.");
  });

  it("describes a range selection as A1:B2", () => {
    const provider: CellProvider = {
      getCell: (row, col) => ({ row, col, value: row === 1 && col === 1 ? "hello" : null })
    };

    expect(
      describeCell(
        { row: 1, col: 1 },
        { startRow: 1, endRow: 3, startCol: 1, endCol: 3 },
        provider,
        1,
        1
      )
    ).toBe("Active cell A1, value hello. Selection A1:B2.");
  });

  it("uses image alt text when the cell value is blank", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 1 && col === 1
          ? { row, col, value: null, image: { imageId: "img1", altText: "Logo" } }
          : { row, col, value: null }
    };

    expect(describeCell({ row: 1, col: 1 }, null, provider, 1, 1)).toBe("Active cell A1, value Logo. Selection none.");
  });

  it("falls back to [Image] when the cell value is blank and no alt text is present", () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 1 && col === 1
          ? { row, col, value: null, image: { imageId: "img1" } }
          : { row, col, value: null }
    };

    expect(describeCell({ row: 1, col: 1 }, null, provider, 1, 1)).toBe("Active cell A1, value [Image]. Selection none.");
  });
});
