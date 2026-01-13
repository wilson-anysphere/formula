import { describe, expect, it, vi } from "vitest";

import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider cache sizing", () => {
  it("evicts cells when sheetCacheMaxSize is exceeded", () => {
    const doc = {
      getCell: vi.fn((_sheetId: string, coord: { row: number; col: number }) => ({
        value: `${coord.row},${coord.col}`,
        formula: null,
      })),
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "sheet-1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
      sheetCacheMaxSize: 2,
    });

    provider.getCell(0, 0);
    provider.getCell(0, 1);
    provider.getCell(0, 2);
    expect(doc.getCell).toHaveBeenCalledTimes(3);

    // LRU cache should be capped to maxSize.
    expect(provider.getCacheStats().sheetCache).toEqual({ size: 2, max: 2 });

    // (0,0) should have been evicted (least recently used) and must be refetched.
    provider.getCell(0, 0);
    expect(doc.getCell).toHaveBeenCalledTimes(4);
  });

  it("keeps default cache sizes when not configured", () => {
    const doc = {
      getCell: vi.fn(() => ({ value: "x", formula: null })),
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "sheet-1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const stats = provider.getCacheStats();
    expect(stats.sheetCache.max).toBe(50_000);
    expect(stats.resolvedFormatCache.max).toBe(10_000);
    expect(stats.sheetColResolvedFormatCache.max).toBe(10_000);
    expect(stats.sheetRowResolvedFormatCache.max).toBe(10_000);
    expect(stats.sheetRunResolvedFormatCache.max).toBe(10_000);
    expect(stats.sheetCellResolvedFormatCache.max).toBe(10_000);
  });
});

