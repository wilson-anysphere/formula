import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { createDefaultLayout, getPanelPlacement, openPanel, closePanel, setSplitDirection as setSplitDirectionState } from "../layout/layoutState.js";
import { PanelIds, panelRegistry } from "../panels/panelRegistry.js";

import { registerBuiltinCommands } from "./registerBuiltinCommands.js";

describe("registerBuiltinCommands: panel toggles", () => {
  const createHarness = () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;
    registerBuiltinCommands({ commandRegistry, app: {} as any, layoutController });
    return { commandRegistry, layoutController };
  };

  it("toggles required panel toggle commands open/closed", async () => {
    const { commandRegistry, layoutController } = createHarness();

    const cases: Array<{ commandId: string; panelId: string; expectedSide: "left" | "right" | "bottom" }> = [
      { commandId: "view.togglePanel.aiChat", panelId: PanelIds.AI_CHAT, expectedSide: "right" },
      { commandId: "view.togglePanel.aiAudit", panelId: PanelIds.AI_AUDIT, expectedSide: "right" },
      { commandId: "view.togglePanel.extensions", panelId: PanelIds.EXTENSIONS, expectedSide: "left" },
      { commandId: "view.togglePanel.macros", panelId: PanelIds.MACROS, expectedSide: "right" },
      { commandId: "view.togglePanel.dataQueries", panelId: PanelIds.DATA_QUERIES, expectedSide: "right" },
      { commandId: "view.togglePanel.scriptEditor", panelId: PanelIds.SCRIPT_EDITOR, expectedSide: "bottom" },
      { commandId: "view.togglePanel.python", panelId: PanelIds.PYTHON, expectedSide: "bottom" },
    ];

    for (const { commandId, panelId, expectedSide } of cases) {
      expect(getPanelPlacement(layoutController.layout, panelId).kind).toBe("closed");

      await commandRegistry.executeCommand(commandId);
      expect(getPanelPlacement(layoutController.layout, panelId)).toEqual({ kind: "docked", side: expectedSide });

      await commandRegistry.executeCommand(commandId);
      expect(getPanelPlacement(layoutController.layout, panelId).kind).toBe("closed");
    }
  });

  it("toggles Version History panel open/closed", async () => {
    const { commandRegistry, layoutController } = createHarness();

    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY).kind).toBe("closed");

    await commandRegistry.executeCommand("view.togglePanel.versionHistory");
    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY)).toEqual({ kind: "docked", side: "right" });

    await commandRegistry.executeCommand("view.togglePanel.versionHistory");
    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY).kind).toBe("closed");
  });

  it("toggles Branch Manager panel open/closed", async () => {
    const { commandRegistry, layoutController } = createHarness();

    expect(getPanelPlacement(layoutController.layout, PanelIds.BRANCH_MANAGER).kind).toBe("closed");

    await commandRegistry.executeCommand("view.togglePanel.branchManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.BRANCH_MANAGER)).toEqual({ kind: "docked", side: "right" });

    await commandRegistry.executeCommand("view.togglePanel.branchManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.BRANCH_MANAGER).kind).toBe("closed");
  });

  it("invokes ensureExtensionsLoaded when toggling Extensions panel", async () => {
    const commandRegistry = new CommandRegistry();

    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const ensureExtensionsLoaded = vi.fn(async () => {});
    const onExtensionsLoaded = vi.fn();

    registerBuiltinCommands({
      commandRegistry,
      app: {} as any,
      layoutController,
      ensureExtensionsLoaded,
      onExtensionsLoaded,
    });

    await commandRegistry.executeCommand("view.togglePanel.extensions");
    expect(getPanelPlacement(layoutController.layout, PanelIds.EXTENSIONS)).toEqual({ kind: "docked", side: "left" });
    expect(ensureExtensionsLoaded).toHaveBeenCalledTimes(1);

    // `view.togglePanel.extensions` schedules `onExtensionsLoaded` via a promise continuation.
    await Promise.resolve();
    await Promise.resolve();
    expect(onExtensionsLoaded).toHaveBeenCalledTimes(1);
  });
});

describe("registerBuiltinCommands: Home tab core commands", () => {
  it("registers clipboard, formatting, and find/replace/go-to commands", () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(_panelId: string) {},
      closePanel(_panelId: string) {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app: {} as any, layoutController });

    const required = [
      "clipboard.cut",
      "clipboard.copy",
      "clipboard.paste",
      "clipboard.pasteSpecial",
      "format.toggleBold",
      "format.toggleItalic",
      "format.toggleUnderline",
      "format.toggleWrapText",
      "format.fontSize.increase",
      "format.fontSize.decrease",
      "format.fontColor",
      "format.fillColor",
      "format.numberFormat.currency",
      "format.numberFormat.percent",
      "format.numberFormat.date",
      "edit.find",
      "edit.replace",
      "navigation.goTo",
    ];

    for (const commandId of required) {
      expect(commandRegistry.getCommand(commandId)).toBeTruthy();
    }
  });
});

describe("registerBuiltinCommands: sheet navigation", () => {
  it("uses DocumentController.getVisibleSheetIds when UI sheet-store order is not provided", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    let current = "Sheet1";
    const activated: string[] = [];

    const doc = {
      getSheetIds: () => ["Sheet1", "Sheet2", "Sheet3"],
      // Sheet2 is hidden, so visible order should skip it.
      getVisibleSheetIds: () => ["Sheet1", "Sheet3"],
    };

    const app = {
      getDocument: () => doc,
      getCurrentSheetId: () => current,
      isEditing: () => false,
      activateSheet: (id: string) => {
        current = id;
        activated.push(id);
      },
      focusAfterSheetNavigation: () => {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("workbook.nextSheet");
    expect(current).toBe("Sheet3");

    // Wrap around.
    await commandRegistry.executeCommand("workbook.nextSheet");
    expect(current).toBe("Sheet1");

    // Wrap around backwards too.
    await commandRegistry.executeCommand("workbook.previousSheet");
    expect(current).toBe("Sheet3");

    expect(activated).toEqual(["Sheet3", "Sheet1", "Sheet3"]);
  });

  it("blocks sheet navigation while editing non-formula content", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    let current = "Sheet1";
    const activated: string[] = [];

    const doc = {
      getVisibleSheetIds: () => ["Sheet1", "Sheet2"],
    };

    const app = {
      getDocument: () => doc,
      getCurrentSheetId: () => current,
      isEditing: () => true,
      isFormulaBarFormulaEditing: () => false,
      activateSheet: (id: string) => {
        current = id;
        activated.push(id);
      },
      focusAfterSheetNavigation: () => {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("workbook.nextSheet");
    expect(current).toBe("Sheet1");
    expect(activated).toEqual([]);
  });

  it("allows sheet navigation while the formula bar is actively editing a formula", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    let current = "Sheet1";
    const activated: string[] = [];

    const doc = {
      getVisibleSheetIds: () => ["Sheet1", "Sheet2"],
    };

    const app = {
      getDocument: () => doc,
      getCurrentSheetId: () => current,
      isEditing: () => true,
      isFormulaBarFormulaEditing: () => true,
      activateSheet: (id: string) => {
        current = id;
        activated.push(id);
      },
      focusAfterSheetNavigation: () => {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("workbook.nextSheet");
    expect(current).toBe("Sheet2");
    expect(activated).toEqual(["Sheet2"]);
  });

  it("jumps to the first visible sheet when the current sheet is not visible", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    let current = "Sheet2";
    const activated: string[] = [];
    const focusAfterSheetNavigation = vi.fn();

    const doc = {
      // Sheet2 is hidden.
      getVisibleSheetIds: () => ["Sheet1", "Sheet3"],
    };

    const app = {
      getDocument: () => doc,
      getCurrentSheetId: () => current,
      isEditing: () => false,
      activateSheet: (id: string) => {
        current = id;
        activated.push(id);
      },
      focusAfterSheetNavigation: () => {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController, focusAfterSheetNavigation });

    await commandRegistry.executeCommand("workbook.nextSheet");
    expect(current).toBe("Sheet1");
    expect(focusAfterSheetNavigation).toHaveBeenCalledTimes(1);

    // Previous sheet should behave the same (deterministic "jump to first visible").
    current = "Sheet2";
    await commandRegistry.executeCommand("workbook.previousSheet");
    expect(current).toBe("Sheet1");
    expect(focusAfterSheetNavigation).toHaveBeenCalledTimes(2);

    expect(activated).toEqual(["Sheet1", "Sheet1"]);
  });
});

describe("registerBuiltinCommands: view toggles", () => {
  it("view.togglePerformanceStats toggles current state when next is omitted", async () => {
    const commandRegistry = new CommandRegistry();

    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    let enabled = false;
    const setGridPerfStatsEnabled = vi.fn((value: boolean) => {
      enabled = value;
    });

    const app = {
      getGridPerfStats: () => ({ enabled }),
      setGridPerfStatsEnabled,
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("view.togglePerformanceStats");
    expect(setGridPerfStatsEnabled).toHaveBeenLastCalledWith(true);

    await commandRegistry.executeCommand("view.togglePerformanceStats");
    expect(setGridPerfStatsEnabled).toHaveBeenLastCalledWith(false);
  });

  it("view.toggleSplitView matches ribbon semantics (default vertical 0.5, toggles off to none)", async () => {
    const commandRegistry = new CommandRegistry();

    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
      setSplitDirection: vi.fn(),
    } as any;
    layoutController.setSplitDirection.mockImplementation((direction: "none" | "vertical" | "horizontal", ratio?: number) => {
      layoutController.layout = setSplitDirectionState(layoutController.layout, direction, ratio);
    });

    const app = {
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    // Toggle on when currently none -> default to vertical 0.5.
    expect(layoutController.layout.splitView.direction).toBe("none");
    await commandRegistry.executeCommand("view.toggleSplitView");
    expect(layoutController.setSplitDirection).toHaveBeenLastCalledWith("vertical", 0.5);
    expect(layoutController.layout.splitView.direction).toBe("vertical");
    expect(app.focus).toHaveBeenCalledTimes(1);

    // Toggle off -> set to none.
    await commandRegistry.executeCommand("view.toggleSplitView");
    expect(layoutController.setSplitDirection).toHaveBeenLastCalledWith("none");
    expect(layoutController.layout.splitView.direction).toBe("none");
    expect(app.focus).toHaveBeenCalledTimes(2);

    // Turning on when already split should not change direction.
    layoutController.layout = setSplitDirectionState(layoutController.layout, "horizontal", 0.6);
    (layoutController.setSplitDirection as any).mockClear();
    (app.focus as any).mockClear();

    await commandRegistry.executeCommand("view.toggleSplitView", true);
    expect(layoutController.setSplitDirection).not.toHaveBeenCalled();
    expect(app.focus).toHaveBeenCalledTimes(1);
  });
});

describe("registerBuiltinCommands: core editing/view/audit commands", () => {
  it("registers required commands and respects edit-state guards", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      undo: vi.fn(),
      redo: vi.fn(),
      isEditing: vi.fn(() => false),
      setShowFormulas: vi.fn(),
      toggleShowFormulas: vi.fn(),
      toggleAuditingPrecedents: vi.fn(),
      toggleAuditingDependents: vi.fn(),
      selectCurrentRegion: vi.fn(),
      openCellEditorAtActiveCell: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    for (const id of [
      "edit.undo",
      "edit.redo",
      "view.toggleShowFormulas",
      "audit.togglePrecedents",
      "audit.toggleDependents",
      "edit.editCell",
      "edit.selectCurrentRegion",
    ]) {
      expect(commandRegistry.getCommand(id)).toBeDefined();
    }

    await commandRegistry.executeCommand("edit.undo");
    await commandRegistry.executeCommand("edit.redo");
    expect(app.undo).toHaveBeenCalledTimes(1);
    expect(app.redo).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("view.toggleShowFormulas");
    await commandRegistry.executeCommand("view.toggleShowFormulas", true);
    await commandRegistry.executeCommand("audit.togglePrecedents");
    await commandRegistry.executeCommand("audit.toggleDependents");
    await commandRegistry.executeCommand("edit.editCell");
    await commandRegistry.executeCommand("edit.selectCurrentRegion");
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.setShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.setShowFormulas).toHaveBeenCalledWith(true);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);
    expect(app.openCellEditorAtActiveCell).toHaveBeenCalledTimes(1);
    expect(app.selectCurrentRegion).toHaveBeenCalledTimes(1);

    // When editing, these commands should no-op (Excel-like behavior).
    app.isEditing.mockReturnValue(true);
    await commandRegistry.executeCommand("view.toggleShowFormulas");
    await commandRegistry.executeCommand("view.toggleShowFormulas", false);
    await commandRegistry.executeCommand("audit.togglePrecedents");
    await commandRegistry.executeCommand("audit.toggleDependents");
    await commandRegistry.executeCommand("edit.editCell");
    await commandRegistry.executeCommand("edit.selectCurrentRegion");
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.setShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);
    expect(app.openCellEditorAtActiveCell).toHaveBeenCalledTimes(1);
    expect(app.selectCurrentRegion).toHaveBeenCalledTimes(1);

    // Sanity check: Edit Cell is keyword-searchable by its Excel shortcut.
    expect(commandRegistry.getCommand("edit.editCell")?.keywords).toEqual(expect.arrayContaining(["f2"]));
  });

  it("executes formulas.formulaAuditing.tracePrecedents by clearing auditing then toggling precedents", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const calls: string[] = [];
    const app = {
      isEditing: vi.fn(() => false),
      clearAuditing: vi.fn(() => calls.push("clearAuditing")),
      toggleAuditingPrecedents: vi.fn(() => calls.push("toggleAuditingPrecedents")),
      focus: vi.fn(() => calls.push("focus")),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    // Coverage: ensure the ribbon command id is registered.
    expect(commandRegistry.getCommand("formulas.formulaAuditing.tracePrecedents")).toBeDefined();

    await commandRegistry.executeCommand("formulas.formulaAuditing.tracePrecedents");
    expect(calls).toEqual(["clearAuditing", "toggleAuditingPrecedents", "focus"]);
  });

  it("uses document.execCommand for undo/redo when a text input is focused", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      undo: vi.fn(),
      redo: vi.fn(),
      isEditing: vi.fn(() => false),
      toggleShowFormulas: vi.fn(),
      toggleAuditingPrecedents: vi.fn(),
      toggleAuditingDependents: vi.fn(),
    } as any;

    const execCommand = vi.fn(() => true);
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = {
      activeElement: { tagName: "INPUT", isContentEditable: false },
      execCommand,
    };

    try {
      registerBuiltinCommands({ commandRegistry, app, layoutController });
      await commandRegistry.executeCommand("edit.undo");
      await commandRegistry.executeCommand("edit.redo");
      await commandRegistry.executeCommand("view.toggleShowFormulas");
      await commandRegistry.executeCommand("audit.togglePrecedents");
      await commandRegistry.executeCommand("audit.toggleDependents");
      await commandRegistry.executeCommand("audit.tracePrecedents");
      await commandRegistry.executeCommand("audit.traceDependents");
      await commandRegistry.executeCommand("audit.traceBoth");
      await commandRegistry.executeCommand("audit.clearAuditing");
      await commandRegistry.executeCommand("audit.toggleTransitive");
    } finally {
      if (prevDocument === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        delete (globalThis as any).document;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).document = prevDocument;
      }
    }

    expect(execCommand).toHaveBeenCalledWith("undo", false);
    expect(execCommand).toHaveBeenCalledWith("redo", false);
    expect(app.undo).not.toHaveBeenCalled();
    expect(app.redo).not.toHaveBeenCalled();
    expect(app.toggleShowFormulas).not.toHaveBeenCalled();
    expect(app.toggleAuditingPrecedents).not.toHaveBeenCalled();
    expect(app.toggleAuditingDependents).not.toHaveBeenCalled();
  });

  it("uses document.execCommand for undo/redo while the formula bar is editing (range selection mode)", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      undo: vi.fn(),
      redo: vi.fn(),
      // In this test, focus is on the grid (not an input), but the formula bar is still editing.
      isFormulaBarEditing: vi.fn(() => true),
      focusFormulaBar: vi.fn(),
    } as any;

    const execCommand = vi.fn(() => true);
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = {
      activeElement: { tagName: "DIV", isContentEditable: false },
      execCommand,
    };

    try {
      registerBuiltinCommands({ commandRegistry, app, layoutController });
      await commandRegistry.executeCommand("edit.undo");
      await commandRegistry.executeCommand("edit.redo");
    } finally {
      if (prevDocument === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        delete (globalThis as any).document;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).document = prevDocument;
      }
    }

    expect(app.focusFormulaBar).toHaveBeenCalledTimes(2);
    expect(execCommand).toHaveBeenCalledWith("undo", false);
    expect(execCommand).toHaveBeenCalledWith("redo", false);
    expect(app.undo).not.toHaveBeenCalled();
    expect(app.redo).not.toHaveBeenCalled();
  });

  it("uses document.execCommand for clipboard commands while the formula bar is editing (range selection mode)", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      copyToClipboard: vi.fn(),
      cutToClipboard: vi.fn(),
      pasteFromClipboard: vi.fn(),
      // Focus is on the grid (not an input), but the formula bar is still editing.
      isFormulaBarEditing: vi.fn(() => true),
      focusFormulaBar: vi.fn(),
    } as any;

    const execCommand = vi.fn(() => true);
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = {
      activeElement: { tagName: "DIV", isContentEditable: false },
      execCommand,
    };

    try {
      registerBuiltinCommands({ commandRegistry, app, layoutController });
      await commandRegistry.executeCommand("clipboard.copy");
      await commandRegistry.executeCommand("clipboard.cut");
      await commandRegistry.executeCommand("clipboard.paste");
    } finally {
      if (prevDocument === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        delete (globalThis as any).document;
      } else {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (globalThis as any).document = prevDocument;
      }
    }

    expect(app.focusFormulaBar).toHaveBeenCalledTimes(3);
    expect(execCommand).toHaveBeenCalledWith("copy", false);
    expect(execCommand).toHaveBeenCalledWith("cut", false);
    expect(execCommand).toHaveBeenCalledWith("paste", false);
    expect(app.copyToClipboard).not.toHaveBeenCalled();
    expect(app.cutToClipboard).not.toHaveBeenCalled();
    expect(app.pasteFromClipboard).not.toHaveBeenCalled();
  });

  it("executes view.zoom.zoom100 by setting zoom=1 (100%)", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      supportsZoom: vi.fn(() => true),
      setZoom: vi.fn(),
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });
    await commandRegistry.executeCommand("view.zoom.zoom100");

    expect(app.setZoom).toHaveBeenCalledWith(1);
    expect(app.focus).toHaveBeenCalledTimes(1);
  });

  it("executes view.zoom.zoomToSelection by invoking SpreadsheetApp.zoomToSelection", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      supportsZoom: vi.fn(() => true),
      zoomToSelection: vi.fn(),
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });
    await commandRegistry.executeCommand("view.zoom.zoomToSelection");

    expect(app.zoomToSelection).toHaveBeenCalledTimes(1);
    expect(app.focus).toHaveBeenCalledTimes(1);
  });
});

describe("registerBuiltinCommands: theme preference commands", () => {
  it("registers theme commands that update ThemeController and refresh ribbon UI state", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = { focus: vi.fn() } as any;

    const themeController = { setThemePreference: vi.fn() } as any;
    const refreshRibbonUiState = vi.fn();

    registerBuiltinCommands({
      commandRegistry,
      app,
      layoutController,
      themeController,
      refreshRibbonUiState,
    });

    await commandRegistry.executeCommand("view.appearance.theme.dark");
    expect(themeController.setThemePreference).toHaveBeenCalledWith("dark");
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(1);
    expect(app.focus).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("view.appearance.theme.highContrast");
    expect(themeController.setThemePreference).toHaveBeenCalledWith("high-contrast");
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(2);
    expect(app.focus).toHaveBeenCalledTimes(2);

    // Commands should be discoverable in the command palette under View.
    expect(commandRegistry.getCommand("view.appearance.theme.dark")).toMatchObject({
      commandId: "view.appearance.theme.dark",
      category: "View",
    });
    expect(commandRegistry.getCommand("view.appearance.theme.dark")?.keywords).toEqual(
      expect.arrayContaining(["theme", "dark"]),
    );
  });
});
