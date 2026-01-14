import { describe, expect, it, vi } from "vitest";

import { handleHomeCellsInsertDeleteCommand } from "../homeCellsCommands.js";

describe("Home → Cells dropdown commands", () => {
  it("no-ops while the spreadsheet is editing (split-view secondary editor via global flag)", async () => {
    const showQuickPick = vi.fn(async () => "shiftRight" as const);
    const showToast = vi.fn();
    const focus = vi.fn();

    const app = {
      isEditing: () => false,
      focus,
    } as any;

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      const handled = await handleHomeCellsInsertDeleteCommand({
        app,
        commandId: "home.cells.insert.insertCells",
        showQuickPick,
        showToast,
      });

      expect(handled).toBe(true);
      expect(showQuickPick).not.toHaveBeenCalled();
      expect(showToast).not.toHaveBeenCalled();
      expect(focus).not.toHaveBeenCalled();
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }
  });

  it("prompts for Insert Cells direction and calls app.insertCells", async () => {
    const showQuickPick = vi.fn(async () => "shiftRight" as const);
    const showToast = vi.fn();
    const insertCells = vi.fn(async () => {});
    const focus = vi.fn();

    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [{ startRow: 3, endRow: 2, startCol: 1, endCol: 0 }],
      insertCells,
      focus,
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.insert.insertCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(showQuickPick).toHaveBeenCalledTimes(1);
    const items = showQuickPick.mock.calls[0]?.[0] ?? [];
    expect(items.map((i: any) => i.label)).toEqual(["Shift cells right", "Shift cells down", "Entire row", "Entire column"]);

    expect(insertCells).toHaveBeenCalledWith({ startRow: 2, endRow: 3, startCol: 0, endCol: 1 }, "right");
    expect(focus).toHaveBeenCalled();
  });

  it("prompts for Delete Cells direction and calls app.deleteCells", async () => {
    const showQuickPick = vi.fn(async () => "shiftUp" as const);
    const showToast = vi.fn();
    const deleteCells = vi.fn(async () => {});
    const focus = vi.fn();

    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      deleteCells,
      focus,
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.delete.deleteCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(showQuickPick).toHaveBeenCalledTimes(1);
    const items = showQuickPick.mock.calls[0]?.[0] ?? [];
    expect(items.map((i: any) => i.label)).toEqual(["Shift cells left", "Shift cells up", "Entire row", "Entire column"]);

    expect(deleteCells).toHaveBeenCalledWith({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, "up");
    expect(focus).toHaveBeenCalled();
  });

  it("blocks multi-range selections", async () => {
    const showQuickPick = vi.fn(async () => "shiftRight" as const);
    const showToast = vi.fn();

    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [
        { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        { startRow: 1, endRow: 1, startCol: 1, endCol: 1 },
      ],
      insertCells: vi.fn(async () => {}),
      getActiveCell: () => ({ row: 0, col: 0 }),
      focus: vi.fn(),
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.insert.insertCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(showQuickPick).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalled();
  });

  it("defaults to active cell when selection ranges are empty", async () => {
    const showQuickPick = vi.fn(async () => "shiftDown" as const);
    const showToast = vi.fn();
    const insertCells = vi.fn(async () => {});
    const focus = vi.fn();

    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [],
      getActiveCell: () => ({ row: 5, col: 7 }),
      insertCells,
      focus,
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.insert.insertCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(showToast).not.toHaveBeenCalled();
    expect(insertCells).toHaveBeenCalledWith({ startRow: 5, endRow: 5, startCol: 7, endCol: 7 }, "down");
    expect(focus).toHaveBeenCalled();
  });

  it("routes Entire row insertion through DocumentController structural edits", async () => {
    const showQuickPick = vi.fn(async () => "entireRow" as const);
    const showToast = vi.fn();
    const focus = vi.fn();
    const doc = {
      insertRows: vi.fn(),
    };

    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [{ startRow: 10, endRow: 12, startCol: 3, endCol: 4 }],
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      focus,
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.insert.insertCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(doc.insertRows).toHaveBeenCalledWith("Sheet1", 10, 3, { label: "Insert Rows", source: "ribbon" });
    expect(focus).toHaveBeenCalled();
  });

  it("blocks commands in read-only mode without prompting", async () => {
    const showQuickPick = vi.fn(async () => "shiftRight" as const);
    const showToast = vi.fn();
    const focus = vi.fn();

    const app = {
      isEditing: () => false,
      isReadOnly: () => true,
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      getActiveCell: () => ({ row: 0, col: 0 }),
      focus,
    } as any;

    const handled = await handleHomeCellsInsertDeleteCommand({
      app,
      commandId: "home.cells.insert.insertCells",
      showQuickPick,
      showToast,
    });

    expect(handled).toBe(true);
    expect(showQuickPick).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalled();
    expect(focus).toHaveBeenCalled();
  });
});
