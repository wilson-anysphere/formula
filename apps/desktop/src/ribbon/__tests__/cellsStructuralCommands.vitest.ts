/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { executeCellsStructuralRibbonCommand } from "../cellsStructuralCommands";

describe("executeCellsStructuralRibbonCommand", () => {
  beforeEach(() => {
    document.body.innerHTML = `<div id=\"toast-root\"></div>`;
  });

  it("inserts sheet rows based on a full-row band selection", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "top");
    doc.setCellValue("Sheet1", "A3", "moved");

    let focused = 0;
    const app = {
      isEditing: () => false,
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      getSelectionRanges: () => [
        // Full-row band over 2 rows (within the current grid limits).
        { startRow: 1, endRow: 2, startCol: 0, endCol: 4 },
      ],
      getActiveCell: () => ({ row: 1, col: 0 }),
      getGridLimits: () => ({ maxRows: 10, maxCols: 5 }),
      focus: () => {
        focused += 1;
      },
    };

    const handled = executeCellsStructuralRibbonCommand(app as any, "home.cells.insert.insertSheetRows");
    expect(handled).toBe(true);
    expect(focused).toBe(1);

    expect(doc.getCell("Sheet1", "A1").value).toBe("top");
    // A3 (row 2) shifts down by 2 -> A5 (row 4).
    expect(doc.getCell("Sheet1", "A3").value).toBe(null);
    expect(doc.getCell("Sheet1", "A5").value).toBe("moved");
  });

  it("no-ops while the spreadsheet is editing (split-view secondary editor via global flag)", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", "top");
    doc.setCellValue("Sheet1", "A3", "moved");

    const focus = vi.fn();
    const app = {
      isEditing: () => false,
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      getSelectionRanges: () => [
        { startRow: 1, endRow: 2, startCol: 0, endCol: 4 },
      ],
      getActiveCell: () => ({ row: 1, col: 0 }),
      getGridLimits: () => ({ maxRows: 10, maxCols: 5 }),
      focus,
    };

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      const handled = executeCellsStructuralRibbonCommand(app as any, "home.cells.insert.insertSheetRows");
      expect(handled).toBe(true);
      expect(focus).not.toHaveBeenCalled();

      // No structural edits should be applied while any editor is active.
      expect(doc.getCell("Sheet1", "A1").value).toBe("top");
      expect(doc.getCell("Sheet1", "A3").value).toBe("moved");
      expect(doc.getCell("Sheet1", "A5").value).toBe(null);
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }
  });

  it("shows a read-only toast when structural edits are blocked", () => {
    const focus = vi.fn();
    const app = {
      isEditing: () => false,
      isReadOnly: () => true,
      focus,
    };

    const handled = executeCellsStructuralRibbonCommand(app as any, "home.cells.insert.insertSheetRows");
    expect(handled).toBe(true);
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("insert rows");
    expect(focus).toHaveBeenCalled();
  });
});
