import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";

const mocks = vi.hoisted(() => ({
  openFormatCellsDialog: vi.fn(),
}));

vi.mock("../openFormatCellsDialog.js", () => ({
  openFormatCellsDialog: mocks.openFormatCellsDialog,
}));

describe("format.openFormatCells", () => {
  it("executes the Format Cells dialog opener", async () => {
    const { createOpenFormatCells } = await import("../openFormatCellsCommand.js");

    const commandRegistry = new CommandRegistry();
    const host = {
      isEditing: () => false,
      getDocument: () => ({}),
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [],
      getGridLimits: () => ({ maxRows: 1000, maxCols: 1000 }),
      focusGrid: () => {},
    };

    const openFormatCells = createOpenFormatCells(host as any);
    commandRegistry.registerBuiltinCommand("format.openFormatCells", "Format Cells…", openFormatCells);

    await commandRegistry.executeCommand("format.openFormatCells");

    expect(mocks.openFormatCellsDialog).toHaveBeenCalledTimes(1);
    expect(mocks.openFormatCellsDialog).toHaveBeenCalledWith(host);
  });
});
