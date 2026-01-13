import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createRibbonActions } from "../../ribbon/ribbonCommandRouter";

describe("Organize Sheets ribbon wiring", () => {
  it("routes the ribbon command id to openOrganizeSheets()", async () => {
    const openOrganizeSheets = vi.fn();

    const ribbonActions = createRibbonActions({
      app: {
        focus: vi.fn(),
        getDocument: () => ({} as any),
        getCurrentSheetId: () => "Sheet1",
        getActiveCell: () => ({ row: 0, col: 0 }),
        getGridLimits: () => ({ maxRows: 1000, maxCols: 1000 }),
        getSelectionRanges: () => [],
        isReadOnly: () => false,
        getShowFormulas: () => false,
      } as any,
      commandRegistry: new CommandRegistry(),
      isSpreadsheetEditing: () => false,
      showToast: vi.fn(),
      showQuickPick: async () => null,
      showInputBox: async () => null,
      notify: vi.fn(),
      showDesktopOnlyToast: vi.fn(),
      getTauriBackend: () => null,
      handleSaveAs: async () => {},
      handleExportDelimitedText: vi.fn(),
      openOrganizeSheets,
      handleAddSheet: async () => {},
      handleDeleteActiveSheet: async () => {},
      openCustomSortDialog: vi.fn(),
      toggleAutoFilter: vi.fn(),
      clearAutoFilter: vi.fn(),
      reapplyAutoFilter: vi.fn(),
      applyAutoFilterFromSelection: async () => false,
      scheduleRibbonSelectionFormatStateUpdate: vi.fn(),
      applyFormattingToSelection: vi.fn(),
      getActiveCellNumberFormat: () => null,
      getEnsureExtensionsLoadedRef: () => null,
      getSyncContributedCommandsRef: () => null,
    });

    ribbonActions.onCommand?.("home.cells.format.organizeSheets");
    // `createRibbonActionsFromCommands` runs handlers in an async wrapper.
    await new Promise<void>((resolve) => setImmediate(resolve));
    expect(openOrganizeSheets).toHaveBeenCalledTimes(1);
  });
});
