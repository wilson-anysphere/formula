import { describe, expect, it } from "vitest";

import { rangeToA1 } from "./a1";
import { computeCurrentRegionRange } from "./currentRegion";
import type { CellCoord, Range } from "./types";

describe("computeCurrentRegionRange", () => {
  it("returns the bounding rectangle of the connected component of non-empty cells (4-neighborhood)", () => {
    const cells = new Map<string, { value: unknown; formula: string | null }>();
    const setValue = (row: number, col: number, value: unknown, formula: string | null = null) => {
      cells.set(`${row},${col}`, { value, formula });
    };

    // Non-rectangular "cross" shape around B2 (1,1).
    setValue(1, 1, "center");
    setValue(0, 1, "up");
    setValue(2, 1, "down");
    setValue(1, 0, "left");
    // Formula-only cells should count as non-empty.
    setValue(1, 2, null, "=A1");

    // Isolated cell elsewhere should not affect the current region for the active cell.
    setValue(4, 4, "isolated");

    const data = {
      getUsedRange(): Range | null {
        if (cells.size === 0) return null;
        let minRow = Infinity;
        let minCol = Infinity;
        let maxRow = -Infinity;
        let maxCol = -Infinity;
        for (const key of cells.keys()) {
          const [rowStr, colStr] = key.split(",");
          const row = Number(rowStr);
          const col = Number(colStr);
          minRow = Math.min(minRow, row);
          minCol = Math.min(minCol, col);
          maxRow = Math.max(maxRow, row);
          maxCol = Math.max(maxCol, col);
        }
        return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
      },
      isCellEmpty(cell: CellCoord): boolean {
        const entry = cells.get(`${cell.row},${cell.col}`);
        return entry == null || (entry.value == null && entry.formula == null);
      },
    };

    const range = computeCurrentRegionRange({ row: 1, col: 1 }, data, { maxRows: 100, maxCols: 100 });
    expect(rangeToA1(range)).toBe("A1:C3");
  });

  it("selects a region when the active cell is empty but adjacent to non-empty cells", () => {
    const cells = new Map<string, { value: unknown; formula: string | null }>();
    const setValue = (row: number, col: number, value: unknown, formula: string | null = null) => {
      cells.set(`${row},${col}`, { value, formula });
    };

    // L-shape (A1, B1, A2). Active cell B2 is empty but adjacent to both A2/B1.
    setValue(0, 0, "A1");
    setValue(0, 1, "B1");
    setValue(1, 0, "A2");

    const data = {
      getUsedRange(): Range | null {
        let minRow = Infinity;
        let minCol = Infinity;
        let maxRow = -Infinity;
        let maxCol = -Infinity;
        for (const key of cells.keys()) {
          const [rowStr, colStr] = key.split(",");
          const row = Number(rowStr);
          const col = Number(colStr);
          minRow = Math.min(minRow, row);
          minCol = Math.min(minCol, col);
          maxRow = Math.max(maxRow, row);
          maxCol = Math.max(maxCol, col);
        }
        return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
      },
      isCellEmpty(cell: CellCoord): boolean {
        const entry = cells.get(`${cell.row},${cell.col}`);
        return entry == null || (entry.value == null && entry.formula == null);
      },
    };

    const range = computeCurrentRegionRange({ row: 1, col: 1 }, data, { maxRows: 100, maxCols: 100 });
    expect(rangeToA1(range)).toBe("A1:B2");
  });

  it("falls back to the active cell if the active cell is empty and has no non-empty neighbor", () => {
    const data = {
      getUsedRange: () => null,
      isCellEmpty: () => true,
    };

    const range = computeCurrentRegionRange({ row: 10, col: 5 }, data, { maxRows: 100, maxCols: 100 });
    expect(rangeToA1(range)).toBe("F11");
  });
});
