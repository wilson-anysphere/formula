import { describe, expect, it } from "vitest";

import { cellToA1, rangeToA1 } from "./a1";
import { navigateSelectionByKey } from "./navigation";
import type { GridLimits } from "./types";
import { addCellToSelection, createSelection, extendSelectionToCell, selectAll, selectColumns, selectRows, setActiveCell } from "./selection";
import { SheetModel } from "../sheet/sheetModel";

const limits: GridLimits = { maxRows: 100, maxCols: 50 };

describe("selection model", () => {
  it("creates a single-cell selection with anchor=active", () => {
    const s = createSelection({ row: 0, col: 0 }, limits);
    expect(s.type).toBe("cell");
    expect(cellToA1(s.active)).toBe("A1");
    expect(cellToA1(s.anchor)).toBe("A1");
    expect(rangeToA1(s.ranges[0])).toBe("A1");
  });

  it("setActiveCell collapses selection to a single cell and resets anchor", () => {
    const start = createSelection({ row: 0, col: 0 }, limits);
    const s = setActiveCell(start, { row: 5, col: 2 }, limits);
    expect(s.type).toBe("cell");
    expect(cellToA1(s.active)).toBe("C6");
    expect(cellToA1(s.anchor)).toBe("C6");
    expect(rangeToA1(s.ranges[0])).toBe("C6");
  });

  it("extendSelectionToCell uses anchor/active semantics to produce a rectangular range", () => {
    const start = createSelection({ row: 0, col: 0 }, limits);
    const s = extendSelectionToCell(start, { row: 2, col: 3 }, limits);
    expect(s.type).toBe("range");
    expect(cellToA1(s.anchor)).toBe("A1");
    expect(cellToA1(s.active)).toBe("D3");
    expect(rangeToA1(s.ranges[0])).toBe("A1:D3");
  });

  it("addCellToSelection creates a multi-range selection (Ctrl/Cmd+click)", () => {
    const start = createSelection({ row: 0, col: 0 }, limits);
    const s = addCellToSelection(start, { row: 4, col: 4 }, limits);
    expect(s.type).toBe("multi");
    expect(s.ranges).toHaveLength(2);
    expect(rangeToA1(s.ranges[0])).toBe("A1");
    expect(rangeToA1(s.ranges[1])).toBe("E5");
    expect(cellToA1(s.active)).toBe("E5");
  });

  it("selectRows selects full width for the requested rows", () => {
    const start = createSelection({ row: 2, col: 2 }, limits);
    const s = selectRows(start, 3, 5, {}, limits);
    expect(s.type).toBe("row");
    expect(rangeToA1(s.ranges[0])).toBe("A4:AX6");
  });

  it("selectColumns selects full height for the requested columns", () => {
    const start = createSelection({ row: 2, col: 2 }, limits);
    const s = selectColumns(start, 1, 3, {}, limits);
    expect(s.type).toBe("column");
    expect(rangeToA1(s.ranges[0])).toBe("B1:D100");
  });

  it("selectAll selects the entire sheet", () => {
    const s = selectAll(limits);
    expect(s.type).toBe("all");
    expect(rangeToA1(s.ranges[0])).toBe("A1:AX100");
  });
});

describe("keyboard navigation", () => {
  it("Arrow keys move the active cell and collapse to a single-cell selection", () => {
    const sheet = new SheetModel();
    const start = createSelection({ row: 0, col: 0 }, limits);
    const next = navigateSelectionByKey(start, "ArrowRight", { shift: false, primary: false }, sheet, limits);
    expect(next).not.toBeNull();
    expect(next?.type).toBe("cell");
    expect(cellToA1(next!.active)).toBe("B1");
  });

  it("Ctrl+End jumps to the bottom-right used cell", () => {
    const sheet = new SheetModel();
    sheet.setCellValue({ row: 0, col: 0 }, "A1");
    sheet.setCellValue({ row: 9, col: 9 }, "J10");

    const start = createSelection({ row: 0, col: 0 }, limits);
    const next = navigateSelectionByKey(start, "End", { shift: false, primary: true }, sheet, limits);
    expect(next).not.toBeNull();
    expect(cellToA1(next!.active)).toBe("J10");
  });

  it("Ctrl+Shift+Arrow extends selection to the edge of data", () => {
    const sheet = new SheetModel();
    sheet.setCellValue({ row: 0, col: 0 }, "A1");
    sheet.setCellValue({ row: 0, col: 3 }, "D1");

    const start = createSelection({ row: 0, col: 0 }, limits);
    const next = navigateSelectionByKey(start, "ArrowRight", { shift: true, primary: true }, sheet, limits);
    expect(next).not.toBeNull();
    expect(next!.type).toBe("range");
    expect(rangeToA1(next!.ranges[0])).toBe("A1:D1");
    expect(cellToA1(next!.active)).toBe("D1");
    expect(cellToA1(next!.anchor)).toBe("A1");
  });
});

