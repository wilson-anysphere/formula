import { describe, expect, it, vi } from "vitest";

import { handleHomeCellsInsertDeleteCommand } from "../homeCellsCommands.js";

describe("Home → Cells dropdown commands", () => {
  it("prompts for Insert Cells direction and calls app.insertCells", async () => {
    const showQuickPick = vi.fn(async () => "right" as const);
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
    expect(items.map((i: any) => i.label)).toEqual(["Shift cells right", "Shift cells down"]);

    expect(insertCells).toHaveBeenCalledWith({ startRow: 2, endRow: 3, startCol: 0, endCol: 1 }, "right");
    expect(focus).toHaveBeenCalled();
  });

  it("prompts for Delete Cells direction and calls app.deleteCells", async () => {
    const showQuickPick = vi.fn(async () => "up" as const);
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
    expect(items.map((i: any) => i.label)).toEqual(["Shift cells left", "Shift cells up"]);

    expect(deleteCells).toHaveBeenCalledWith({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, "up");
    expect(focus).toHaveBeenCalled();
  });

  it("blocks multi-range selections", async () => {
    const showQuickPick = vi.fn(async () => "right" as const);
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
    const showQuickPick = vi.fn(async () => "down" as const);
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
});
