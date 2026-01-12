import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider formatting integration", () => {
  it("maps resolved DocumentController styles into @formula/grid CellStyle", () => {
    const doc = new DocumentController();

    // Back-compat for older controller builds: map the existing style table into the
    // new `getCellFormat()` API shape expected by DocumentCellProvider.
    (doc as any).getCellFormat ??= (sheetId: string, coord: { row: number; col: number }) => {
      const cell = doc.getCell(sheetId, coord);
      const styleId = typeof cell?.styleId === "number" ? cell.styleId : 0;
      return styleId === 0 ? null : doc.styleTable.get(styleId);
    };

    const headerRows = 1;
    const headerCols = 1;
    const docRows = 200;
    const docCols = 10;

    // Apply formatting to the full height of column A within this sheet.
    doc.setRangeFormat(
      "Sheet1",
      { start: { row: 0, col: 0 }, end: { row: docRows - 1, col: 0 } },
      {
        fill: { pattern: "solid", fgColor: "#FFFFFF00" },
        font: { bold: true, color: "#FF00FF00", size: 12, name: "Arial" },
        alignment: { horizontal: "right", wrapText: true }
      }
    );

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

  it("emits grid invalidation updates for format-layer-only deltas", () => {
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
      rowStyleDeltas: [{ sheetId: "Sheet1", row: 10 }]
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
        endRow: headerRows + 11,
        startCol: headerCols,
        endCol: headerCols + docCols
      }
    });
  });
});
