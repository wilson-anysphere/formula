import { describe, expect, it, vi, afterEach } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createRibbonActions } from "../ribbonCommandRouter";

async function flushMicrotasks(): Promise<void> {
  // Allow any `queueMicrotask` / async IIFE work to run.
  // Nested async boundaries (CommandRegistry -> createRibbonActionsFromCommands) can
  // require multiple microtask turns in vitest, so flush a small batch.
  for (let i = 0; i < 8; i += 1) {
    await Promise.resolve();
  }
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("createRibbonActions edit-mode guards", () => {
  it("no-ops unknown commands that are disabled while editing (no unimplemented toast)", async () => {
    const commandRegistry = new CommandRegistry();
    const showToast = vi.fn();

    const app = {
      focus: vi.fn(),
      // Minimal surface area to satisfy `createRibbonActions` helpers.
      getGridLimits: () => ({ maxRows: 100, maxCols: 100 }),
      getSelectionRanges: () => [],
      getDocument: () => ({}),
      getCurrentSheetId: () => "Sheet1",
      isReadOnly: () => false,
    } as any;

    const actions = createRibbonActions({
      app,
      commandRegistry,
      isSpreadsheetEditing: () => true,
      showToast,
      showQuickPick: async () => null,
      showInputBox: async () => null,
      openOrganizeSheets: vi.fn(),
      handleAddSheet: vi.fn(async () => {}),
      handleDeleteActiveSheet: vi.fn(async () => {}),
      openCustomSortDialog: vi.fn(),
      toggleAutoFilter: vi.fn(),
      clearAutoFilter: vi.fn(),
      reapplyAutoFilter: vi.fn(),
      applyFormattingToSelection: vi.fn(),
      getActiveCellNumberFormat: () => null,
      getEnsureExtensionsLoadedRef: () => null,
      getSyncContributedCommandsRef: () => null,
    });

    // `clipboard.copy` is disabled while editing, but in this harness it is *not* registered,
    // so it would previously fall through to an "unimplemented" ribbon toast.
    actions.onCommand?.("clipboard.copy");
    await flushMicrotasks();

    expect(showToast).not.toHaveBeenCalled();
    expect(app.focus).not.toHaveBeenCalled();
  });
});

