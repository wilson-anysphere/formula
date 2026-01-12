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
      font: { bold: true, color: "#FF00FF00", size: 12, name: "Arial" },
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

  it("supports flat/clipboard-ish style keys when mapping styleId -> CellStyle", () => {
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
});
