import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider formatting integration", () => {
  it("maps resolved DocumentController styles into @formula/grid CellStyle", () => {
    const doc = new DocumentController();

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 200;
    const docCols = 10;

    // Apply formatting to a full-height column range. DocumentController should treat
    // this as a column formatting layer update (not per-cell deltas).
    const EXCEL_MAX_ROW = 1_048_576 - 1;
    doc.setRangeFormat("Sheet1", { start: { row: 0, col: 0 }, end: { row: EXCEL_MAX_ROW, col: 0 } }, {
      fill: { pattern: "solid", fgColor: "#FFFFFF00" },
      font: { bold: true, color: "#FF00FF00", size: 12, name: "  Arial  " },
      alignment: { horizontal: "right", wrapText: true }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const deepRow = 150;
    const cell = provider.getCell(headerRows + deepRow, headerCols + 0);
    expect(cell).not.toBeNull();
    expect(cell?.style).toEqual(
      expect.objectContaining({
        fontWeight: "700",
        fontFamily: "Arial",
        fontSize: 16,
        textAlign: "end",
        wrapMode: "word"
      })
    );

    // Ensure we convert Excel-style `#AARRGGBB` into a canvas-compatible CSS color string.
    // (The exact representation can vary: `rgba(...)` or `#RRGGBB` are both acceptable.)
    expect(cell?.style?.fill).toMatch(/^(rgba\(|#[0-9a-f]{6}$)/i);
    expect(cell?.style?.fill).not.toMatch(/^#[0-9a-f]{8}$/i);
    expect(cell?.style?.color).toMatch(/^(rgba\(|#[0-9a-f]{6}$)/i);
    expect(cell?.style?.color).not.toMatch(/^#[0-9a-f]{8}$/i);
  });

  it("maps vertical alignment + text rotation from DocumentController styles into grid CellStyle", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { vertical: "top", textRotation: 45 } });

    const headerRows = 1;
    const headerCols = 1;

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell?.style?.verticalAlign).toBe("top");
    expect(cell?.style?.rotationDeg).toBe(45);
  });

  it('maps alignment.vertical "center" to grid verticalAlign "middle"', () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { vertical: "center" } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 3,
      colCount: 3,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.style?.verticalAlign).toBe("middle");
  });

  it('maps alignment.vertical "bottom" to grid verticalAlign "bottom"', () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { vertical: "bottom" } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 3,
      colCount: 3,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.style?.verticalAlign).toBe("bottom");
  });

  it("prefers alignment.textRotation over alignment.rotation", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { textRotation: 45, rotation: 30 } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 3,
      colCount: 3,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.style?.rotationDeg).toBe(45);
  });

  it("falls back to alignment.rotation when textRotation is not finite", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { textRotation: Number.NaN, rotation: 30 } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 3,
      colCount: 3,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell?.style?.rotationDeg).toBe(30);
  });

  it("clamps rotation values and handles Excel vertical-text sentinel 255", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setCellValue("Sheet1", "A2", "hello");
    doc.setCellValue("Sheet1", "A3", "hello");

    doc.setRangeFormat("Sheet1", "A1", { alignment: { textRotation: 999 } });
    doc.setRangeFormat("Sheet1", "A2", { alignment: { textRotation: -999 } });
    doc.setRangeFormat("Sheet1", "A3", { alignment: { rotation: 255 } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 5,
      colCount: 3,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    expect(provider.getCell(1, 1)?.style?.rotationDeg).toBe(180);
    expect(provider.getCell(2, 1)?.style?.rotationDeg).toBe(-180);
    expect(provider.getCell(3, 1)?.style?.rotationDeg).toBe(90);
  });

  it("falls back fill/justify horizontal alignment to a deterministic textAlign value (and preserves semantics)", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "Fill");
    doc.setCellValue("Sheet1", "A2", "Justify");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { horizontal: "fill" } });
    doc.setRangeFormat("Sheet1", "A2", { alignment: { horizontal: "justify" } });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 3 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const fillCell = provider.getCell(headerRows, headerCols);
    const justifyCell = provider.getCell(headerRows + 1, headerCols);

    expect(fillCell?.style?.textAlign).toBe("start");
    expect((fillCell?.style as any)?.horizontalAlign).toBe("fill");

    expect(justifyCell?.style?.textAlign).toBe("start");
    expect((justifyCell?.style as any)?.horizontalAlign).toBe("justify");
  });

  it("supports snake_case alignment.horizontal_alignment (formula-model serialization) when mapping styles", () => {
    const doc: any = {
      getCell: (_sheetId: string, _coord: { row: number; col: number }) => ({ value: "x", formula: null, styleId: 1 }),
      styleTable: {
        get: (id: number) => (id === 1 ? { alignment: { horizontal_alignment: "fill" } } : {})
      }
    };

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell?.style?.textAlign).toBe("start");
    expect((cell?.style as any)?.horizontalAlign).toBe("fill");
  });

  it("supports flat/clipboard-ish style keys when mapping styleId -> CellStyle", () => {
    const doc = new DocumentController();

    // Back-compat: some historical snapshots/tests/clipboard round-trips store flat keys.
    doc.setRangeFormat("Sheet1", "A1", { bold: true, backgroundColor: "#FFFFFF00" });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell).not.toBeNull();
    expect(cell?.style?.fontWeight).toBe("700");
    expect(cell?.style?.fill).toBe("#ffff00");
  });

  it("supports additional flat/clipboard-ish aliases (snake_case and legacy keys)", () => {
    const doc = new DocumentController();

    // Back-compat: some historical snapshots/tests/clipboard round-trips store flat keys.
    doc.setRangeFormat("Sheet1", "A1", {
      bold: true,
      backgroundColor: "#FFFFFF00",
      font_size: 12,
      font_color: "#FF00FF00",
      horizontal_align: "right",
    });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell).not.toBeNull();
    expect(cell?.style?.fontWeight).toBe("700");
    expect(cell?.style?.fill).toBe("#ffff00");
    expect(cell?.style?.fontSize).toBe(16);
    expect(cell?.style?.color).toBe("#00ff00");
    expect(cell?.style?.textAlign).toBe("end");
  });

  it("maps diagonal border flags from DocumentController styles into grid CellStyle.diagonalBorders", () => {
    const doc = new DocumentController();

    doc.setCellValue("Sheet1", "A1", "diag");
    doc.setRangeFormat("Sheet1", "A1", {
      border: {
        diagonal: { style: "thin", color: "#FF0000FF" },
        diagonalDown: true
      }
    });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell).not.toBeNull();
    expect(cell?.style?.diagonalBorders).toEqual({
      down: { width: 1, style: "solid", color: "rgba(0,0,255,1)" }
    });
    // Ensure we convert Excel-style `#AARRGGBB` into a canvas-compatible CSS color string.
    expect((cell?.style as any)?.diagonalBorders?.down?.color).not.toMatch(/^#[0-9a-f]{8}$/i);
  });

  it("maps formula-model snake_case formatting keys into @formula/grid CellStyle", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");

    doc.setRangeFormat("Sheet1", "A1", {
      font: { size_100pt: 1200 },
      fill: { fg_color: "#FFFFFF00" },
      alignment: { wrap_text: true }
    });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell).not.toBeNull();
    expect(cell?.style?.fill).toBe("#ffff00");
    expect(cell?.style?.wrapMode).toBe("word");
    expect(cell?.style?.fontSize).toBeCloseTo(16, 5);
  });

  it("maps Excel alignment.indent into the shared-grid indent primitive", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "hello");
    doc.setRangeFormat("Sheet1", "A1", { alignment: { indent: 2 } });

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 2 + headerRows,
      colCount: 2 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(headerRows, headerCols);
    expect(cell).not.toBeNull();
    // 8px per indent level (see DocumentCellProvider).
    expect((cell?.style as any)?.textIndentPx).toBe(16);
  });

  it("caches resolved @formula/grid CellStyle by getCellFormatStyleIds tuple", () => {
    const getCellFormat = vi.fn(() => ({ font: { bold: true } }));
    const doc = {
      getCell: vi.fn(() => ({ value: "x", formula: null, styleId: 0 })),
      getCellFormatStyleIds: vi.fn(() => [0, 0, 0, 1] as [number, number, number, number]),
      getCellFormat,
      on: vi.fn(() => () => {})
    };

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    // Different coordinates but identical formatting tuple.
    provider.getCell(0, 0);
    provider.getCell(0, 1);
    provider.getCell(1, 0);

    expect(getCellFormat).toHaveBeenCalledTimes(1);
  });

  it("emits grid invalidation updates for format-layer-only deltas", () => {
    const doc = new DocumentController();

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 100;
    const docCols = 20;

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const updates: any[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    // Format a full-height column; this should emit format-layer-only deltas (no per-cell deltas)
    // and trigger a redraw for that column across all visible rows.
    const EXCEL_MAX_ROW = 1_048_576 - 1;
    doc.setRangeFormat("Sheet1", { start: { row: 0, col: 0 }, end: { row: EXCEL_MAX_ROW, col: 0 } }, { font: { bold: true } });

    unsubscribe();

    expect(updates.length).toBeGreaterThan(0);
    expect(updates[0]).toEqual({
      type: "cells",
      range: {
        startRow: headerRows,
        endRow: headerRows + docRows,
        startCol: headerCols,
        endCol: headerCols + 1
      }
    });
  });

  it("does not clear other sheet caches when a large invalidation occurs", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "one");
    doc.setCellValue("Sheet2", "A1", "two");

    let activeSheet = "Sheet1";

    const headerRows = 1;
    const headerCols = 1;

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => activeSheet,
      headerRows,
      headerCols,
      rowCount: 10 + headerRows,
      colCount: 10 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const spy = vi.spyOn(doc, "getCell");

    expect(provider.getCell(headerRows, headerCols)?.value).toBe("one");
    activeSheet = "Sheet2";
    expect(provider.getCell(headerRows, headerCols)?.value).toBe("two");
    expect(spy).toHaveBeenCalledTimes(2);

    // This is large enough to trigger the provider's "large invalidation" path.
    activeSheet = "Sheet1";
    provider.invalidateDocCells({ startRow: 0, endRow: 50, startCol: 0, endCol: 50 });

    // Sheet1 cache is cleared.
    expect(provider.getCell(headerRows, headerCols)?.value).toBe("one");
    expect(spy).toHaveBeenCalledTimes(3);

    // Sheet2 cache remains intact.
    activeSheet = "Sheet2";
    expect(provider.getCell(headerRows, headerCols)?.value).toBe("two");
    expect(spy).toHaveBeenCalledTimes(3);
  });

  it("emits grid invalidation updates for range-run formatting deltas", () => {
    class FakeDocument {
      private readonly listeners = new Set<(payload: any) => void>();

      on(event: string, listener: (payload: any) => void): () => void {
        if (event !== "change") return () => {};
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
      }

      emitChange(payload: any): void {
        for (const listener of this.listeners) listener(payload);
      }
    }

    const doc = new FakeDocument();

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 100;
    const docCols = 20;

    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const updates: any[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    doc.emitChange({
      deltas: [],
      sheetViewDeltas: [],
      rangeRunDeltas: [{ sheetId: "Sheet1", col: 2, startRow: 10, endRowExclusive: 20, beforeRuns: [], afterRuns: [] }]
    });

    unsubscribe();

    expect(updates.length).toBeGreaterThan(0);
    const update = updates[0];

    if (update.type === "invalidateAll") {
      expect(update.type).toBe("invalidateAll");
      return;
    }

    expect(update).toEqual({
      type: "cells",
      range: {
        startRow: headerRows + 10,
        endRow: headerRows + 20,
        startCol: headerCols + 2,
        endCol: headerCols + 3
      }
    });
  });

  it("does not invalidate for pure sheetViewDeltas", () => {
    class FakeDocument {
      private readonly listeners = new Set<(payload: any) => void>();

      getCell = vi.fn(() => ({ value: "hello", formula: null }));

      on(event: string, listener: (payload: any) => void): () => void {
        if (event !== "change") return () => {};
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
      }

      emitChange(payload: any): void {
        for (const listener of this.listeners) listener(payload);
      }
    }

    const doc = new FakeDocument();

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 10 + headerRows,
      colCount: 10 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const updates: any[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    // Prime cache for A1 (doc 0,0 => grid 1,1).
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("hello");
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    // Cache hit
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("hello");
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    // Emit a sheet-view-only delta; this should not invalidate the provider.
    doc.emitChange({
      deltas: [],
      sheetViewDeltas: [{ sheetId: "Sheet1", frozenRows: 1 }],
      rowStyleDeltas: [],
      colStyleDeltas: [],
      sheetStyleDeltas: [],
      formatDeltas: [],
      rangeRunDeltas: [],
      recalc: false
    });

    unsubscribe();

    expect(updates).toEqual([]);
    // Cache should still be intact (no extra doc reads).
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("hello");
    expect(doc.getCell).toHaveBeenCalledTimes(1);
  });

  it("invalidates for sheetViewDeltas that change mergedRanges", () => {
    class FakeDocument {
      private readonly listeners = new Set<(payload: any) => void>();

      mergedRanges: any[] = [];
      getMergedRanges = vi.fn(() => this.mergedRanges);
      getCell = vi.fn(() => ({ value: "hello", formula: null }));

      on(event: string, listener: (payload: any) => void): () => void {
        if (event !== "change") return () => {};
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
      }

      emitChange(payload: any): void {
        for (const listener of this.listeners) listener(payload);
      }
    }

    const doc = new FakeDocument();

    const headerRows = 1;
    const headerCols = 1;
    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: 10 + headerRows,
      colCount: 10 + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const updates: any[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    doc.mergedRanges = [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]; // A1:B1 (inclusive)
    doc.emitChange({
      deltas: [],
      sheetViewDeltas: [
        {
          sheetId: "Sheet1",
          before: { frozenRows: 0, frozenCols: 0 },
          after: { frozenRows: 0, frozenCols: 0, mergedRanges: doc.mergedRanges }
        }
      ],
      rowStyleDeltas: [],
      colStyleDeltas: [],
      sheetStyleDeltas: [],
      formatDeltas: [],
      rangeRunDeltas: [],
      recalc: false
    });

    unsubscribe();

    expect(updates).toEqual([{ type: "invalidateAll" }]);

    // Ensure merged range queries resolve with header offsets and exclusive end coords.
    expect(provider.getMergedRangeAt(headerRows + 0, headerCols + 1)).toEqual({
      startRow: headerRows + 0,
      endRow: headerRows + 1,
      startCol: headerCols + 0,
      endCol: headerCols + 2
    });
    expect(provider.getMergedRangesInRange({ startRow: headerRows, endRow: headerRows + 1, startCol: headerCols, endCol: headerCols + 2 })).toEqual(
      [{ startRow: headerRows + 0, endRow: headerRows + 1, startCol: headerCols + 0, endCol: headerCols + 2 }]
    );
  });

  it("does not emit additional range invalidations after invalidateAll is triggered", () => {
    class FakeDocument {
      private readonly listeners = new Set<(payload: any) => void>();

      getCell = vi.fn(() => ({ value: "hello", formula: null }));

      on(event: string, listener: (payload: any) => void): () => void {
        if (event !== "change") return () => {};
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
      }

      emitChange(payload: any): void {
        for (const listener of this.listeners) listener(payload);
      }
    }

    const doc = new FakeDocument();

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 1000;
    const docCols = 1000;
    const provider = new DocumentCellProvider({
      document: doc as any,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const updates: any[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    // Prime a sheet cache so `invalidateDocCells` would otherwise take the scan path.
    provider.getCell(headerRows + 0, headerCols + 0);
    expect(doc.getCell).toHaveBeenCalledTimes(1);

    doc.emitChange({
      deltas: [],
      sheetViewDeltas: [],
      // Create two invalidation spans. The row span is huge (1M cells) and should trigger invalidateAll,
      // and we should not additionally emit a separate column invalidation after that.
      rowStyleDeltas: [{ sheetId: "Sheet1", startRow: 0, endRowExclusive: docRows }],
      colStyleDeltas: [{ sheetId: "Sheet1", col: 0 }],
      sheetStyleDeltas: [],
      formatDeltas: [],
      rangeRunDeltas: []
    });

    unsubscribe();

    expect(updates).toEqual([{ type: "invalidateAll" }]);
  });

  it("caches layered-format style resolution by style-id tuple", () => {
    const doc = new DocumentController();

    // Apply column formatting so many cells share the same effective style tuple
    // (sheet default + row + col + cell).
    doc.setColFormat("Sheet1", 0, { font: { bold: true } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 500,
      colCount: 2,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const convertSpy = vi.spyOn(provider as any, "convertDocStyleToGridStyle");

    for (let row = 0; row < 250; row++) {
      const cell = provider.getCell(row, 0);
      expect(cell?.style?.fontWeight).toBe("700");
    }

    // Without tuple caching we'd convert once per cell because DocumentController.getCellFormat()
    // creates a fresh merged object every call (defeating the WeakMap cache).
    expect(convertSpy).toHaveBeenCalledTimes(1);
  });

  it("honors explicit-false overrides when merging layered formatting", () => {
    const doc = new DocumentController();

    doc.setSheetFormat("Sheet1", {
      font: { bold: true, italic: true, underline: true, strike: true },
      alignment: { wrapText: true }
    });

    // Explicit `false` should clear inherited `true` values.
    doc.setRowFormat("Sheet1", 0, {
      font: { bold: false, italic: false, underline: false, strike: false },
      alignment: { wrapText: false }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 0,
      headerCols: 0,
      rowCount: 2,
      colCount: 1,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cleared = provider.getCell(0, 0);
    expect(cleared?.style).toBeUndefined();

    const inherited = provider.getCell(1, 0);
    expect(inherited?.style).toEqual(
      expect.objectContaining({
        fontWeight: "700",
        fontStyle: "italic",
        underline: true,
        strike: true,
        wrapMode: "word"
      })
    );
  });

  it("keeps other rows cached when invalidating a full-width row range", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "row1");
    doc.setCellValue("Sheet1", "A2", "row2");

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 10;
    const docCols = 2001; // > 1000 so row-wide invalidation is considered "large"

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const getCellSpy = vi.spyOn(doc, "getCell");

    // Prime cache for two rows (same column).
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("row1");
    expect(provider.getCell(headerRows + 1, headerCols + 0)?.value).toBe("row2");
    expect(getCellSpy).toHaveBeenCalledTimes(2);

    // Invalidate row 0 across all columns (doc range).
    provider.invalidateDocCells({ startRow: 0, endRow: 1, startCol: 0, endCol: docCols });

    // Row 0 should be refetched...
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("row1");
    expect(getCellSpy).toHaveBeenCalledTimes(3);

    // ...but row 1 should stay cached.
    expect(provider.getCell(headerRows + 1, headerCols + 0)?.value).toBe("row2");
    expect(getCellSpy).toHaveBeenCalledTimes(3);
  });

  it("keeps other columns cached when invalidating a full-height column range", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "colA");
    doc.setCellValue("Sheet1", "B1", "colB");

    const headerRows = 1;
    const headerCols = 1;
    // > 50k so column-wide invalidation is considered "large" (would previously clear the whole sheet cache).
    const docRows = 60_000;
    const docCols = 2;

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const getCellSpy = vi.spyOn(doc, "getCell");

    // Prime cache for two columns (same row).
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("colA");
    expect(provider.getCell(headerRows + 0, headerCols + 1)?.value).toBe("colB");
    expect(getCellSpy).toHaveBeenCalledTimes(2);

    // Invalidate column 0 across all rows (doc range).
    provider.invalidateDocCells({ startRow: 0, endRow: docRows, startCol: 0, endCol: 1 });

    // Column 0 should be refetched...
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("colA");
    expect(getCellSpy).toHaveBeenCalledTimes(3);

    // ...but column 1 should stay cached.
    expect(provider.getCell(headerRows + 0, headerCols + 1)?.value).toBe("colB");
    expect(getCellSpy).toHaveBeenCalledTimes(3);
  });

  it("keeps other columns cached when invalidating multiple full-height columns", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "colA");
    doc.setCellValue("Sheet1", "P1", "colP");

    const headerRows = 1;
    const headerCols = 1;
    // 6000 rows x 10 columns = 60k cells (>50k direct-eviction threshold). This would previously
    // clear the entire sheet cache because the invalidation isn't "thin enough" (width > 4).
    const docRows = 6000;
    const docCols = 20;

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const getCellSpy = vi.spyOn(doc, "getCell");

    // Prime cache for a touched column and an untouched column (same row).
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("colA");
    expect(provider.getCell(headerRows + 0, headerCols + 15)?.value).toBe("colP");
    expect(getCellSpy).toHaveBeenCalledTimes(2);

    // Invalidate columns 0-9 across all rows (doc range).
    provider.invalidateDocCells({ startRow: 0, endRow: docRows, startCol: 0, endCol: 10 });

    // Touched column should be refetched...
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("colA");
    expect(getCellSpy).toHaveBeenCalledTimes(3);

    // ...but untouched column should stay cached.
    expect(provider.getCell(headerRows + 0, headerCols + 15)?.value).toBe("colP");
    expect(getCellSpy).toHaveBeenCalledTimes(3);
  });

  it("does not drop cached cells for large invalidations that do not intersect the cached region", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "cached");

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 2000;
    const docCols = 2000;

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const getCellSpy = vi.spyOn(doc, "getCell");

    // Prime cache for A1.
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("cached");
    expect(getCellSpy).toHaveBeenCalledTimes(1);

    // Invalidate a large 300x300 region far away from A1. This is >50k cells, so we should
    // avoid clearing the whole sheet cache (A1 remains cached).
    provider.invalidateDocCells({ startRow: 500, endRow: 800, startCol: 500, endCol: 800 });

    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("cached");
    expect(getCellSpy).toHaveBeenCalledTimes(1);
  });

  it("keeps other rows cached when invalidating multiple full-width rows", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "row1");
    doc.setCellValue("Sheet1", "A6", "row6");

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 10;
    const docCols = 16_384; // Excel max, ensures 4-row full-width invalidation exceeds 50k cells

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows,
      headerCols,
      rowCount: docRows + headerRows,
      colCount: docCols + headerCols,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const getCellSpy = vi.spyOn(doc, "getCell");

    // Prime cache for an affected row and an unaffected row.
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("row1");
    expect(provider.getCell(headerRows + 5, headerCols + 0)?.value).toBe("row6");
    expect(getCellSpy).toHaveBeenCalledTimes(2);

    // Invalidate rows 0-3 across all columns (doc range). This touches 65,536 cells, but only
    // 4 rows, so we should not clear the entire sheet cache.
    provider.invalidateDocCells({ startRow: 0, endRow: 4, startCol: 0, endCol: docCols });

    // Affected row should be refetched...
    expect(provider.getCell(headerRows + 0, headerCols + 0)?.value).toBe("row1");
    expect(getCellSpy).toHaveBeenCalledTimes(3);

    // ...but unaffected row should stay cached.
    expect(provider.getCell(headerRows + 5, headerCols + 0)?.value).toBe("row6");
    expect(getCellSpy).toHaveBeenCalledTimes(3);
  });
});
