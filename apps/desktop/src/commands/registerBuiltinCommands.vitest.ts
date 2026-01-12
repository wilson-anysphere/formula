import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { createDefaultLayout, getPanelPlacement, openPanel, closePanel } from "../layout/layoutState.js";
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
    await commandRegistry.executeCommand("audit.togglePrecedents");
    await commandRegistry.executeCommand("audit.toggleDependents");
    await commandRegistry.executeCommand("edit.editCell");
    await commandRegistry.executeCommand("edit.selectCurrentRegion");
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);
    expect(app.openCellEditorAtActiveCell).toHaveBeenCalledTimes(1);
    expect(app.selectCurrentRegion).toHaveBeenCalledTimes(1);

    // When editing, these commands should no-op (Excel-like behavior).
    app.isEditing.mockReturnValue(true);
    await commandRegistry.executeCommand("view.toggleShowFormulas");
    await commandRegistry.executeCommand("audit.togglePrecedents");
    await commandRegistry.executeCommand("audit.toggleDependents");
    await commandRegistry.executeCommand("edit.editCell");
    await commandRegistry.executeCommand("edit.selectCurrentRegion");
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);
    expect(app.openCellEditorAtActiveCell).toHaveBeenCalledTimes(1);
    expect(app.selectCurrentRegion).toHaveBeenCalledTimes(1);

    // Sanity check: Edit Cell is keyword-searchable by its Excel shortcut.
    expect(commandRegistry.getCommand("edit.editCell")?.keywords).toEqual(expect.arrayContaining(["f2"]));
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
});
