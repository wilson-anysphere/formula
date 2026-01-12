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
      down: { width: 1, style: "solid", color: "#0000ff" }
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
});
