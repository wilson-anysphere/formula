import { describe, expect, it, vi } from "vitest";

import { DocumentCellProvider } from "../documentCellProvider.js";

type CellState = { value: unknown; formula: string | null; styleId?: number };

function createProvider(options: {
  getSheetId: () => string;
  getCell: (sheetId: string, coord: { row: number; col: number }) => CellState | null;
  headerRows?: number;
  headerCols?: number;
  rowCount?: number;
  colCount?: number;
  sheetCacheMaxSize?: number;
  getComputedValue?: (cell: { row: number; col: number }) => string | number | boolean | null;
  getCommentMeta?: (row: number, col: number) => { resolved: boolean } | null;
}) {
  const headerRows = options.headerRows ?? 1;
  const headerCols = options.headerCols ?? 1;
  const doc = {
    getCell: vi.fn(options.getCell),
    on: vi.fn(() => () => {}),
  };

  const provider = new DocumentCellProvider({
    document: doc as any,
    getSheetId: options.getSheetId,
    headerRows,
    headerCols,
    rowCount: options.rowCount ?? headerRows + 10,
    colCount: options.colCount ?? headerCols + 10,
    showFormulas: () => false,
    getComputedValue: options.getComputedValue ?? (() => null),
    getCommentMeta: options.getCommentMeta,
    sheetCacheMaxSize: options.sheetCacheMaxSize,
  });

  return { provider, doc };
}

describe("DocumentCellProvider (shared grid)", () => {
  it("caches within a sheet (hit/miss correctness)", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
    });

    const first = provider.getCell(1, 1);
    expect(first?.value).toBe("hello");
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    const second = provider.getCell(1, 1);
    expect(second).toBe(first);
    expect(doc.getCell).toHaveBeenCalledTimes(1);
  });

  it("caches header cells", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "should-not-be-called", formula: null }),
    });

    const header = provider.getCell(0, 5);
    expect(header?.value).toBe("E");
    expect(doc.getCell).toHaveBeenCalledTimes(0);

    const headerAgain = provider.getCell(0, 5);
    expect(headerAgain).toBe(header);
    expect(doc.getCell).toHaveBeenCalledTimes(0);
  });

  it("does not collide across sheets", () => {
    let activeSheet = "sheet-1";
    const values = new Map<string, string>([
      ["sheet-1", "one"],
      ["sheet-2", "two"],
    ]);

    const { provider, doc } = createProvider({
      getSheetId: () => activeSheet,
      getCell: (sheetId) => ({ value: values.get(sheetId)!, formula: null }),
    });

    const sheet1Cell = provider.getCell(1, 1);
    expect(sheet1Cell?.value).toBe("one");
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    activeSheet = "sheet-2";
    const sheet2Cell = provider.getCell(1, 1);
    expect(sheet2Cell?.value).toBe("two");
    // If caches collided, this would still be 1 and we'd get the sheet-1 cell back.
    expect(doc.getCell).toHaveBeenCalledTimes(2);
  });

  it("invalidateDocCells evicts only impacted keys for small ranges", () => {
    let activeSheet = "sheet-1";
    const state = new Map<string, string>();
    state.set("0,0", "A");
    state.set("0,1", "B");

    const { provider, doc } = createProvider({
      getSheetId: () => activeSheet,
      getCell: (_sheetId, coord) => {
        return { value: state.get(`${coord.row},${coord.col}`) ?? null, formula: null };
      },
    });

    const a1 = provider.getCell(1, 1);
    const b1 = provider.getCell(1, 2);
    expect(a1?.value).toBe("A");
    expect(b1?.value).toBe("B");
    expect(doc.getCell).toHaveBeenCalledTimes(2);

    // Update backing state for only A1 and invalidate just that doc cell.
    state.set("0,0", "A2");
    provider.invalidateDocCells({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

    const a1After = provider.getCell(1, 1);
    const b1After = provider.getCell(1, 2);

    expect(a1After?.value).toBe("A2");
    expect(b1After).toBe(b1);
    expect(b1After?.value).toBe("B");

    // A1 was re-fetched, B1 was served from cache.
    expect(doc.getCell).toHaveBeenCalledTimes(3);

    // Sanity: switching sheets shouldn't affect the invalidation we just performed.
    activeSheet = "sheet-2";
    state.set("0,0", "S2");
    const sheet2 = provider.getCell(1, 1);
    expect(sheet2?.value).toBe("S2");
  });

  it("invalidateAll clears per-sheet caches", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
    });

    const first = provider.getCell(1, 1);
    expect(first?.value).toBe("hello");
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    // Cache hit
    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    provider.invalidateAll();

    // Cache miss after invalidation
    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(2);
  });

  it("invalidateDocCells clears caches on large invalidations", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
    });

    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    // Large region (50x50=2500 cells) triggers a full cache clear.
    provider.invalidateDocCells({ startRow: 0, endRow: 50, startCol: 0, endCol: 50 });

    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(2);
  });

  it("invalidateDocCells falls back to invalidateAll for huge invalidations", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
    });

    // Populate the cache with some entries.
    provider.getCell(1, 1);
    provider.getCell(2, 2);
    expect(doc.getCell).toHaveBeenCalledTimes(2);

    const updates: any[] = [];
    provider.subscribe((update) => updates.push(update));

    const sheetCache = (provider as any).sheetCaches.get("sheet-1");
    expect(sheetCache).toBeTruthy();
    const keysSpy = vi.spyOn(sheetCache, "keys");

    // Invalidate an enormous region to trigger the invalidateAll heuristic.
    provider.invalidateDocCells({ startRow: 0, endRow: 1_000, startCol: 0, endCol: 1_000 });

    expect(keysSpy).not.toHaveBeenCalled();
    expect(updates).toEqual([{ type: "invalidateAll" }]);
    // `invalidateAll` drops per-sheet caches entirely.
    expect((provider as any).sheetCaches.size).toBe(0);

    // Cache miss after invalidation.
    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(3);
  });

  it("invalidateDocCells still scans the LRU for medium-large invalidations", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
    });

    provider.getCell(1, 1);
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    const updates: any[] = [];
    provider.subscribe((update) => updates.push(update));

    const sheetCache = (provider as any).sheetCaches.get("sheet-1");
    expect(sheetCache).toBeTruthy();
    const keysSpy = vi.spyOn(sheetCache, "keys");

    // 200 * 300 = 60k cells: above the direct-evict cutoff (50k) but below the "huge" invalidation threshold.
    provider.invalidateDocCells({ startRow: 0, endRow: 200, startCol: 0, endCol: 300 });

    expect(keysSpy).toHaveBeenCalled();
    expect(updates).toEqual([
      {
        type: "cells",
        range: { startRow: 1, endRow: 201, startCol: 1, endCol: 301 },
      },
    ]);

    // Cache map is preserved (we didn't drop all per-sheet caches).
    expect((provider as any).sheetCaches.size).toBe(1);
  });

  it("invalidateDocCells uses the configured sheetCacheMaxSize when deciding to invalidateAll", () => {
    const sheetCacheMaxSize = 200_000;
    const { provider } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
      // Make the grid big enough to cover the invalidation region.
      rowCount: 1_000,
      colCount: 1_000,
      sheetCacheMaxSize,
    });

    // Populate cache so `sheetCaches.get(sheetId)` exists.
    provider.getCell(1, 1);

    const updates: any[] = [];
    provider.subscribe((update) => updates.push(update));

    const sheetCache = (provider as any).sheetCaches.get("sheet-1");
    expect(sheetCache).toBeTruthy();
    const keysSpy = vi.spyOn(sheetCache, "keys");

    // 500x500 = 250k cells (~125% of sheetCacheMaxSize). This should take the invalidateAll path.
    provider.invalidateDocCells({ startRow: 0, endRow: 500, startCol: 0, endCol: 500 });

    expect(keysSpy).not.toHaveBeenCalled();
    expect(updates).toEqual([{ type: "invalidateAll" }]);
    expect((provider as any).sheetCaches.size).toBe(0);
  });

  it("looks up comment metadata by numeric coords (no A1 string conversion)", () => {
    const meta = { resolved: false };
    const getCommentMeta = vi.fn((row: number, col: number) => {
      if (row === 0 && col === 0) return meta;
      return null;
    });

    const { provider } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
      getCommentMeta,
    });

    // Grid (1,1) maps to document (0,0) when header rows/cols are enabled.
    const withComment = provider.getCell(1, 1);
    expect(withComment).not.toBeNull();
    expect(getCommentMeta).toHaveBeenCalledWith(0, 0);
    expect(withComment).toMatchObject({ comment: { resolved: false } });
    // Ensure DocumentCellProvider reuses the meta object (no per-cell allocations).
    expect((withComment as any).comment).toBe(meta);

    // Grid (2,2) maps to document (1,1); our meta provider returns null there.
    const withoutComment = provider.getCell(2, 2);
    expect(withoutComment).not.toBeNull();
    expect(getCommentMeta).toHaveBeenCalledWith(1, 1);
    expect("comment" in (withoutComment as any)).toBe(false);
  });

  it("does not request comment metadata for header cells", () => {
    const getCommentMeta = vi.fn(() => ({ resolved: false }));

    const { provider } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null }),
      getCommentMeta,
    });

    // Header row / col cells should not trigger comment lookups.
    provider.getCell(0, 5);
    provider.getCell(5, 0);
    provider.getCell(0, 0);
    expect(getCommentMeta).toHaveBeenCalledTimes(0);
  });

  it("reuses the coord object for getComputedValue calls (no per-cell allocations)", () => {
    const seen: Array<{ row: number; col: number }> = [];
    const getComputedValue = vi.fn((coord: { row: number; col: number }) => {
      seen.push({ row: coord.row, col: coord.col });
      return coord.col;
    });

    const { provider } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: null, formula: "=1" }),
      getComputedValue,
    });

    // Grid (1,1) -> doc (0,0) when headers are enabled.
    const a1 = provider.getCell(1, 1);
    // Grid (1,2) -> doc (0,1).
    const b1 = provider.getCell(1, 2);

    expect(a1?.value).toBe(0);
    expect(b1?.value).toBe(1);

    expect(getComputedValue).toHaveBeenCalledTimes(2);
    const firstArg = getComputedValue.mock.calls[0]?.[0];
    const secondArg = getComputedValue.mock.calls[1]?.[0];
    expect(firstArg).toBeTruthy();
    expect(secondArg).toBeTruthy();
    expect(firstArg).toBe(secondArg);
    expect(seen).toEqual([
      { row: 0, col: 0 },
      { row: 0, col: 1 },
    ]);
  });

  it("skips resolved format cache lookups for default style tuples (no per-cell key strings)", () => {
    const sheetId = "sheet-1";
    const doc = {
      getCell: vi.fn(() => ({ value: "hello", formula: null, styleId: 0 })),
      getCellFormatStyleIds: vi.fn(() => [0, 0, 0, 0, 0]),
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const cacheGetSpy = vi.spyOn((provider as any).resolvedFormatCache, "get");

    const cell = provider.getCell(1, 1);
    expect(cell?.value).toBe("hello");
    expect(doc.getCellFormatStyleIds).toHaveBeenCalledTimes(1);
    expect(cacheGetSpy).not.toHaveBeenCalled();
  });

  it("caches sheet-only formatting without using the resolved-format LRU cache", () => {
    const sheetId = "sheet-1";
    const styleTableGet = vi.fn((id: number) => {
      if (id === 1) return { numberFormat: "0%" };
      return {};
    });
    const doc = {
      getCell: vi.fn(() => ({ value: 1, formula: null, styleId: 0 })),
      getCellFormatStyleIds: vi.fn(() => [1, 0, 0, 0, 0]),
      styleTable: { get: styleTableGet },
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const cacheGetSpy = vi.spyOn((provider as any).resolvedFormatCache, "get");

    const a1 = provider.getCell(1, 1);
    const b2 = provider.getCell(2, 2);

    expect(a1?.value).toBe("100%");
    expect(b2?.value).toBe("100%");

    // `getCellFormatStyleIds` is still consulted per cell, but resolved formatting is cached by sheet default style id.
    expect(doc.getCellFormatStyleIds).toHaveBeenCalledTimes(2);
    // Only the first cell should require reading the style table; the second should hit the cached resolved format.
    expect(styleTableGet).toHaveBeenCalledTimes(2); // once for sheetStyle, once for resolveStyle cache
    expect(cacheGetSpy).not.toHaveBeenCalled();
  });

  it("caches column-only formatting without using the resolved-format LRU cache", () => {
    const sheetId = "sheet-1";
    const styleTableGet = vi.fn((id: number) => {
      if (id === 1) return { numberFormat: "0%" };
      return {};
    });
    const doc = {
      getCell: vi.fn(() => ({ value: 1, formula: null, styleId: 0 })),
      getCellFormatStyleIds: vi.fn(() => [0, 0, 1, 0, 0]),
      styleTable: { get: styleTableGet },
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const cacheGetSpy = vi.spyOn((provider as any).resolvedFormatCache, "get");

    const a1 = provider.getCell(1, 1);
    expect(a1?.value).toBe("100%");
    const callsAfterFirst = styleTableGet.mock.calls.length;

    const a2 = provider.getCell(2, 1);
    expect(a2?.value).toBe("100%");
    expect(styleTableGet).toHaveBeenCalledTimes(callsAfterFirst);
    expect(cacheGetSpy).not.toHaveBeenCalled();
  });

  it("caches row-only formatting without using the resolved-format LRU cache", () => {
    const sheetId = "sheet-1";
    const styleTableGet = vi.fn((id: number) => {
      if (id === 1) return { numberFormat: "0%" };
      return {};
    });
    const doc = {
      getCell: vi.fn(() => ({ value: 1, formula: null, styleId: 0 })),
      getCellFormatStyleIds: vi.fn(() => [0, 1, 0, 0, 0]),
      styleTable: { get: styleTableGet },
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const cacheGetSpy = vi.spyOn((provider as any).resolvedFormatCache, "get");

    const a1 = provider.getCell(1, 1);
    expect(a1?.value).toBe("100%");
    const callsAfterFirst = styleTableGet.mock.calls.length;

    const b1 = provider.getCell(1, 2);
    expect(b1?.value).toBe("100%");
    expect(styleTableGet).toHaveBeenCalledTimes(callsAfterFirst);
    expect(cacheGetSpy).not.toHaveBeenCalled();
  });
});
