import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../../collab/permissionGuards";
import { createAddSheetCommand, createDeleteActiveSheetCommand } from "../sheetCommands";
import { CollabWorkbookSheetStore } from "../collabWorkbookSheetStore";
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

  it("does not delete the last sheet (surfaces error)", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([{ id: "Sheet1", name: "Sheet1", visibility: "visible" }]);
    let activeSheetId = "Sheet1";

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

    const handleDeleteSheet = createDeleteActiveSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
      confirm: async () => true,
    });

    await handleDeleteSheet();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1"]);
    expect(activeSheetId).toBe("Sheet1");
    expect(showToast).toHaveBeenCalledWith("Cannot delete the last sheet", "error");
  });

  it("respects confirmation cancel (no-op)", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "Sheet1";

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

    const confirm = vi.fn(async () => false);
    const handleDeleteSheet = createDeleteActiveSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
      confirm,
    });

    await handleDeleteSheet();

    expect(confirm).toHaveBeenCalledTimes(1);
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);
    expect(activeSheetId).toBe("Sheet1");
    expect(showToast).not.toHaveBeenCalled();
  });

  it("blocks sheet insert/delete in read-only collab sessions", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "Sheet1";

    const session = { isReadOnly: () => true, sheets: { toArray: () => [] } } as any;

    const app = {
      getCurrentSheetId: () => activeSheetId,
      activateSheet: (sheetId: string) => {
        activeSheetId = sheetId;
      },
      getDocument: () => doc,
      getCollabSession: () => session,
    };

    const restoreFocusAfterSheetNavigation = vi.fn();
    const showToast = vi.fn();

    const handleInsertSheet = createAddSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
    });

    const confirm = vi.fn(async () => true);
    const handleDeleteSheet = createDeleteActiveSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
      confirm,
    });

    await handleInsertSheet();
    await handleDeleteSheet();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);
    expect(activeSheetId).toBe("Sheet1");

    // Insert shows toast via `tryInsertCollabSheet`.
    // Delete short-circuits before confirmation.
    expect(showToast).toHaveBeenCalledWith(READ_ONLY_SHEET_MUTATION_MESSAGE, "error");
    expect(confirm).not.toHaveBeenCalled();
  });

  it("inserts a sheet in collab sessions by mutating session.sheets (and activates it)", async () => {
    type CollabSheet = { id: string; name: string; visibility: "visible" | "hidden" | "veryHidden" };
    class MockSheets {
      constructor(private readonly items: CollabSheet[]) {}
      get length(): number {
        return this.items.length;
      }
      toArray(): CollabSheet[] {
        return this.items.slice();
      }
      get(i: number): CollabSheet | undefined {
        return this.items[i];
      }
      insert(index: number, values: CollabSheet[]): void {
        this.items.splice(index, 0, ...values);
      }
      delete(index: number, count: number): void {
        this.items.splice(index, count);
      }
    }

    const doc = new DocumentController();
    const session = {
      isReadOnly: () => false,
      sheets: new MockSheets([
        { id: "sheet_a", name: "Sheet1", visibility: "visible" },
        { id: "sheet_b", name: "Sheet2", visibility: "visible" },
      ]),
      transactLocal: (fn: () => void) => fn(),
    } as any;

    const store = new WorkbookSheetStore([
      { id: "sheet_a", name: "Sheet1", visibility: "visible" },
      { id: "sheet_b", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "sheet_a";
    const app = {
      getCurrentSheetId: () => activeSheetId,
      activateSheet: (sheetId: string) => {
        activeSheetId = sheetId;
      },
      getDocument: () => doc,
      getCollabSession: () => session,
    };

    const restoreFocusAfterSheetNavigation = vi.fn();
    const showToast = vi.fn();

    const handleInsertSheet = createAddSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
    });

    await handleInsertSheet();

    expect(session.sheets.toArray().map((s: CollabSheet) => s.name)).toEqual(["Sheet1", "Sheet3", "Sheet2"]);
    expect(activeSheetId).toBe(session.sheets.toArray()[1]!.id);
    expect(doc.getSheetIds()).toContain(activeSheetId);
    expect(showToast).not.toHaveBeenCalled();
  });

  it("deletes the active sheet in collab sessions via CollabWorkbookSheetStore (and rewrites formulas)", async () => {
    type CollabSheet = { id: string; name: string; visibility: "visible" | "hidden" | "veryHidden" };
    class MockSheets {
      constructor(private readonly items: CollabSheet[]) {}
      get length(): number {
        return this.items.length;
      }
      toArray(): CollabSheet[] {
        return this.items.slice();
      }
      get(i: number): CollabSheet | undefined {
        return this.items[i];
      }
      insert(index: number, values: CollabSheet[]): void {
        this.items.splice(index, 0, ...values);
      }
      delete(index: number, count: number): void {
        this.items.splice(index, count);
      }
    }

    const doc = new DocumentController();
    const session = {
      isReadOnly: () => false,
      sheets: new MockSheets([
        { id: "sheet_a", name: "Sheet1", visibility: "visible" },
        { id: "sheet_b", name: "Sheet2", visibility: "visible" },
        { id: "sheet_c", name: "Sheet3", visibility: "visible" },
      ]),
      transactLocal: (fn: () => void) => fn(),
    } as any;

    const store = new CollabWorkbookSheetStore(
      session,
      [
        { id: "sheet_a", name: "Sheet1", visibility: "visible" },
        { id: "sheet_b", name: "Sheet2", visibility: "visible" },
        { id: "sheet_c", name: "Sheet3", visibility: "visible" },
      ],
      { value: "" },
      { canEditWorkbook: () => true },
    );

    let activeSheetId = "sheet_b";

    doc.setCellFormula("sheet_a", { row: 0, col: 0 }, "=Sheet2!A1");

    const app = {
      getCurrentSheetId: () => activeSheetId,
      activateSheet: (sheetId: string) => {
        activeSheetId = sheetId;
      },
      getDocument: () => doc,
      getCollabSession: () => session,
    };

    const restoreFocusAfterSheetNavigation = vi.fn();
    const showToast = vi.fn();
    const confirm = vi.fn(async () => true);

    const handleDeleteSheet = createDeleteActiveSheetCommand({
      app,
      getWorkbookSheetStore: () => store,
      restoreFocusAfterSheetNavigation,
      showToast,
      confirm,
    });

    await handleDeleteSheet();

    expect(confirm).toHaveBeenCalledTimes(1);
    expect(store.listAll().map((s) => s.name)).toEqual(["Sheet1", "Sheet3"]);
    expect(session.sheets.toArray().map((s: CollabSheet) => s.name)).toEqual(["Sheet1", "Sheet3"]);
    expect(activeSheetId).toBe("sheet_a");

    expect(doc.getCell("sheet_a", { row: 0, col: 0 }).formula).toBe("=#REF!");
    expect(showToast).not.toHaveBeenCalled();
  });
});
