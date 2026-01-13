import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { createAddSheetCommand, createDeleteActiveSheetCommand } from "../sheetCommands";
import { WorkbookSheetStore } from "../workbookSheetStore";

describe("ribbon sheet commands", () => {
  it("inserts and deletes sheets via the ribbon handlers (active sheet + formula rewrite)", async () => {
    const doc = new DocumentController();

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "Sheet1";

    // Ensure the referenced sheet exists in the document model (DocumentController creates lazily).
    doc.getCell("Sheet2", { row: 0, col: 0 });
    doc.setCellFormula("Sheet1", { row: 0, col: 0 }, "=Sheet2!A1");

    const app = {
      getCurrentSheetId: () => activeSheetId,
      activateSheet: (sheetId: string) => {
        activeSheetId = sheetId;
      },
      getDocument: () => doc,
      getCollabSession: () => null,
    };

    const restoreFocusAfterSheetNavigation = vi.fn();
    const showToast = vi.fn();

    const handleInsertSheet = createAddSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
    });

    const handleDeleteSheet = createDeleteActiveSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
      confirm: async () => true,
    });

    // Insert Sheet.
    await handleInsertSheet();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet3", "Sheet2"]);
    expect(activeSheetId).toBe("Sheet3");

    // Delete active Sheet2 (Excel-like).
    activeSheetId = "Sheet2";
    await handleDeleteSheet();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet3"]);
    expect(activeSheetId).toBe("Sheet1");

    // Formula rewrite: direct references to deleted Sheet2 become #REF!.
    expect(doc.getCell("Sheet1", { row: 0, col: 0 }).formula).toBe("=#REF!");

    expect(showToast).not.toHaveBeenCalled();
  });
});

