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

  it("dispose clears caches and unsubscribes from document changes", () => {
    const unsubscribe = vi.fn();
    const doc = {
      getCell: vi.fn(() => ({ value: "x", formula: null })),
      on: vi.fn(() => unsubscribe),
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
      sheetCacheMaxSize: 10,
    });

    // Populate cache and attach a doc subscription via provider.subscribe.
    provider.getCell(0, 0);
    const off = provider.subscribe(() => {});
    expect(doc.on).toHaveBeenCalledTimes(1);

    provider.dispose();
    expect(unsubscribe).toHaveBeenCalledTimes(1);

    // Should not repopulate caches after dispose.
    expect(provider.getCell(0, 0)).toBeNull();
    expect(doc.getCell).toHaveBeenCalledTimes(1);
    expect(provider.getCacheStats().sheetCache.size).toBe(0);

    // Unsubscribing after dispose should be a no-op.
    expect(() => off()).not.toThrow();
  });
});
