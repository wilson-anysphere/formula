import { describe, expect, it } from "vitest";

import { evaluateFormattingSelectionSize } from "../selectionSizeGuard.js";

describe("evaluateFormattingSelectionSize", () => {
  const excelLimits = { maxRows: 1_048_576, maxCols: 16_384 };

  it("allows small non-band selections", () => {
    const res = evaluateFormattingSelectionSize([{ startRow: 0, endRow: 99, startCol: 0, endCol: 99 }], excelLimits);
    expect(res.allowed).toBe(true);
    expect(res.totalCells).toBe(10_000);
    expect(res.allRangesBand).toBe(false);
  });

  it("blocks large non-band selections", () => {
    const res = evaluateFormattingSelectionSize([{ startRow: 0, endRow: 499, startCol: 0, endCol: 499 }], excelLimits);
    expect(res.totalCells).toBe(250_000);
    expect(res.allowed).toBe(false);
    expect(res.allRangesBand).toBe(false);
  });

  it("allows full-column bands even when huge", () => {
    const res = evaluateFormattingSelectionSize(
      [{ startRow: 0, endRow: excelLimits.maxRows - 1, startCol: 0, endCol: 0 }],
      excelLimits,
    );
    expect(res.totalCells).toBe(excelLimits.maxRows);
    expect(res.allowed).toBe(true);
    expect(res.allRangesBand).toBe(true);
  });

  it("allows full-row bands even when huge", () => {
    const res = evaluateFormattingSelectionSize(
      [{ startRow: 0, endRow: 9, startCol: 0, endCol: excelLimits.maxCols - 1 }],
      excelLimits,
    );
    expect(res.totalCells).toBe(10 * excelLimits.maxCols);
    expect(res.allowed).toBe(true);
    expect(res.allRangesBand).toBe(true);
  });

  it("blocks very large full-row selections that exceed the band row cap", () => {
    const res = evaluateFormattingSelectionSize(
      [{ startRow: 0, endRow: 60_000, startCol: 0, endCol: excelLimits.maxCols - 1 }],
      excelLimits,
    );
    expect(res.allowed).toBe(false);
    expect(res.allRangesBand).toBe(false);
  });

  it("allows full-sheet selections even though they exceed the band row cap", () => {
    const res = evaluateFormattingSelectionSize(
      [
        {
          startRow: 0,
          endRow: excelLimits.maxRows - 1,
          startCol: 0,
          endCol: excelLimits.maxCols - 1,
        },
      ],
      excelLimits,
    );
    expect(res.allowed).toBe(true);
    expect(res.allRangesBand).toBe(true);
  });

  it("blocks mixed selections when total cells exceed the cap", () => {
    const res = evaluateFormattingSelectionSize(
      [
        { startRow: 0, endRow: excelLimits.maxRows - 1, startCol: 0, endCol: 0 }, // full column
        { startRow: 0, endRow: 0, startCol: 1, endCol: 1 }, // single cell
      ],
      excelLimits,
    );
    expect(res.totalCells).toBe(excelLimits.maxRows + 1);
    expect(res.allRangesBand).toBe(false);
    expect(res.allowed).toBe(false);
  });
});
