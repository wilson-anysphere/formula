import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import {
  createDefaultLayout,
  getPanelPlacement,
  openPanel,
  closePanel,
  floatPanel,
  setDockCollapsed,
  setFloatingPanelMinimized,
  setSplitDirection as setSplitDirectionState,
} from "../layout/layoutState.js";
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
      setFloatingPanelMinimized(panelId: string, minimized: boolean) {
        this.layout = setFloatingPanelMinimized(this.layout, panelId, minimized);
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
      { commandId: "view.togglePanel.solver", panelId: PanelIds.SOLVER, expectedSide: "right" },
      { commandId: "view.togglePanel.scenarioManager", panelId: PanelIds.SCENARIO_MANAGER, expectedSide: "left" },
      { commandId: "view.togglePanel.monteCarlo", panelId: PanelIds.MONTE_CARLO, expectedSide: "left" },
      { commandId: "view.togglePanel.scriptEditor", panelId: PanelIds.SCRIPT_EDITOR, expectedSide: "bottom" },
      { commandId: "view.togglePanel.python", panelId: PanelIds.PYTHON, expectedSide: "bottom" },
      { commandId: "view.togglePanel.vbaMigrate", panelId: PanelIds.VBA_MIGRATE, expectedSide: "right" },
    ];

    for (const { commandId, panelId, expectedSide } of cases) {
      expect(getPanelPlacement(layoutController.layout, panelId).kind).toBe("closed");

      await commandRegistry.executeCommand(commandId);
      expect(getPanelPlacement(layoutController.layout, panelId)).toEqual({ kind: "docked", side: expectedSide });

      await commandRegistry.executeCommand(commandId);
      expect(getPanelPlacement(layoutController.layout, panelId).kind).toBe("closed");
    }
  });

  it("restores minimized floating panels when toggling a panel open", async () => {
    const { commandRegistry, layoutController } = createHarness();

    // Open the Data Queries panel once so it has a placement.
    await commandRegistry.executeCommand("view.togglePanel.dataQueries");
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).not.toBe("closed");

    // Force it into floating+minimized state.
    layoutController.layout = floatPanel(layoutController.layout, PanelIds.DATA_QUERIES, { x: 10, y: 10, width: 300, height: 200 });
    layoutController.layout = setFloatingPanelMinimized(layoutController.layout, PanelIds.DATA_QUERIES, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(true);

    // Toggling should restore (unminimize) instead of closing.
    await commandRegistry.executeCommand("view.togglePanel.dataQueries");
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(false);
  });

  it("restores collapsed docks when toggling a panel open", async () => {
    const { commandRegistry, layoutController } = createHarness();

    await commandRegistry.executeCommand("view.togglePanel.dataQueries");
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES)).toEqual({ kind: "docked", side: "right" });

    layoutController.layout = setDockCollapsed(layoutController.layout, "right", true);
    expect(layoutController.layout.docks.right.collapsed).toBe(true);

    // Toggling should restore (uncollapse) instead of closing.
    await commandRegistry.executeCommand("view.togglePanel.dataQueries");
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES)).toEqual({ kind: "docked", side: "right" });
    expect(layoutController.layout.docks.right.collapsed).toBe(false);
  });

  it("restores collapsed docks when executing open-panel commands", async () => {
    const { commandRegistry, layoutController } = createHarness();

    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.scenarioManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER)).toEqual({ kind: "docked", side: "left" });

    layoutController.layout = setDockCollapsed(layoutController.layout, "left", true);
    expect(layoutController.layout.docks.left.collapsed).toBe(true);

    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.scenarioManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER)).toEqual({ kind: "docked", side: "left" });
    expect(layoutController.layout.docks.left.collapsed).toBe(false);
  });

  it("activates an already-open docked panel when executing open-panel commands", async () => {
    const { commandRegistry, layoutController } = createHarness();

    // Open the Selection Pane first.
    await commandRegistry.executeCommand("pageLayout.arrange.selectionPane");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SELECTION_PANE)).toEqual({ kind: "docked", side: "right" });
    expect(layoutController.layout.docks.right.active).toBe(PanelIds.SELECTION_PANE);

    // Open another panel in the same dock, making Selection Pane inactive.
    await commandRegistry.executeCommand("view.togglePanel.aiChat");
    expect(getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT)).toEqual({ kind: "docked", side: "right" });
    expect(layoutController.layout.docks.right.active).toBe(PanelIds.AI_CHAT);

    // Re-executing the open-panel command should activate Selection Pane (not no-op).
    await commandRegistry.executeCommand("pageLayout.arrange.selectionPane");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SELECTION_PANE)).toEqual({ kind: "docked", side: "right" });
    expect(layoutController.layout.docks.right.active).toBe(PanelIds.SELECTION_PANE);
  });

  it("restores minimized floating Pivot Builder panel when inserting a pivot table", async () => {
    const { commandRegistry, layoutController } = createHarness();

    // Open once so the panel exists in the layout.
    await commandRegistry.executeCommand("view.insertPivotTable");
    expect(getPanelPlacement(layoutController.layout, PanelIds.PIVOT_BUILDER).kind).not.toBe("closed");

    // Force it into floating+minimized state.
    layoutController.layout = floatPanel(layoutController.layout, PanelIds.PIVOT_BUILDER, { x: 10, y: 10, width: 300, height: 200 });
    layoutController.layout = setFloatingPanelMinimized(layoutController.layout, PanelIds.PIVOT_BUILDER, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.PIVOT_BUILDER).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.PIVOT_BUILDER]?.minimized).toBe(true);

    // Inserting a pivot table should restore (unminimize) the panel.
    await commandRegistry.executeCommand("view.insertPivotTable");
    expect(getPanelPlacement(layoutController.layout, PanelIds.PIVOT_BUILDER).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.PIVOT_BUILDER]?.minimized).toBe(false);
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

  it("does not invoke ensureExtensionsLoaded when toggling Marketplace panel (keep extension host lazy)", async () => {
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

    await commandRegistry.executeCommand("view.togglePanel.marketplace");
    expect(getPanelPlacement(layoutController.layout, PanelIds.MARKETPLACE)).toEqual({ kind: "docked", side: "right" });
    expect(ensureExtensionsLoaded).toHaveBeenCalledTimes(0);

    await commandRegistry.executeCommand("view.togglePanel.marketplace");
    expect(getPanelPlacement(layoutController.layout, PanelIds.MARKETPLACE).kind).toBe("closed");
    expect(ensureExtensionsLoaded).toHaveBeenCalledTimes(0);

    // `view.togglePanel.marketplace` should not schedule onExtensionsLoaded either.
    await Promise.resolve();
    await Promise.resolve();
    expect(onExtensionsLoaded).toHaveBeenCalledTimes(0);
  });
});

describe("registerBuiltinCommands: What-If Analysis + Solver ribbon commands", () => {
  const createHarness = (opts: { openGoalSeekDialog?: () => void } = {}) => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
      setFloatingPanelMinimized(panelId: string, minimized: boolean) {
        this.layout = setFloatingPanelMinimized(this.layout, panelId, minimized);
      },
    } as any;

    registerBuiltinCommands({
      commandRegistry,
      app: {} as any,
      layoutController,
      openGoalSeekDialog: opts.openGoalSeekDialog ?? null,
    });

    return { commandRegistry, layoutController };
  };

  it("opens Scenario Manager / Monte Carlo / Solver panels and dispatches Goal Seek via host callback", async () => {
    const openGoalSeekDialog = vi.fn();
    const { commandRegistry, layoutController } = createHarness({ openGoalSeekDialog });

    expect(commandRegistry.getCommand("data.forecast.whatIfAnalysis.scenarioManager")).toBeTruthy();
    expect(commandRegistry.getCommand("data.forecast.whatIfAnalysis.monteCarlo")).toBeTruthy();
    expect(commandRegistry.getCommand("data.forecast.whatIfAnalysis.goalSeek")).toBeTruthy();
    expect(commandRegistry.getCommand("formulas.solutions.solver")).toBeTruthy();

    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER).kind).toBe("closed");
    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.scenarioManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER)).toEqual({ kind: "docked", side: "left" });

    expect(getPanelPlacement(layoutController.layout, PanelIds.MONTE_CARLO).kind).toBe("closed");
    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.monteCarlo");
    expect(getPanelPlacement(layoutController.layout, PanelIds.MONTE_CARLO)).toEqual({ kind: "docked", side: "left" });

    expect(getPanelPlacement(layoutController.layout, PanelIds.SOLVER).kind).toBe("closed");
    await commandRegistry.executeCommand("formulas.solutions.solver");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SOLVER)).toEqual({ kind: "docked", side: "right" });

    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.goalSeek");
    expect(openGoalSeekDialog).toHaveBeenCalledTimes(1);
  });

  it("restores minimized floating panels when invoking the What-If commands", async () => {
    const { commandRegistry, layoutController } = createHarness({ openGoalSeekDialog: vi.fn() });

    // Open Scenario Manager once so it has a placement.
    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.scenarioManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER).kind).not.toBe("closed");

    // Force it into floating+minimized state.
    layoutController.layout = floatPanel(layoutController.layout, PanelIds.SCENARIO_MANAGER, { x: 10, y: 10, width: 300, height: 200 });
    layoutController.layout = setFloatingPanelMinimized(layoutController.layout, PanelIds.SCENARIO_MANAGER, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.SCENARIO_MANAGER]?.minimized).toBe(true);

    // Invoking the command should restore (unminimize) instead of doing nothing.
    await commandRegistry.executeCommand("data.forecast.whatIfAnalysis.scenarioManager");
    expect(getPanelPlacement(layoutController.layout, PanelIds.SCENARIO_MANAGER).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.SCENARIO_MANAGER]?.minimized).toBe(false);
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

  it("hides clipboard.pasteSpecial.all from the command palette (alias of Paste)", () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(_panelId: string) {},
      closePanel(_panelId: string) {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app: {} as any, layoutController });

    expect(commandRegistry.getCommand("clipboard.pasteSpecial.all")).toBeTruthy();
    expect(commandRegistry.getCommand("clipboard.pasteSpecial.all")?.when).toBe("false");
  });
});

describe("registerBuiltinCommands: read-only formatting defaults", () => {
  const createLayoutController = () =>
    ({
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(_panelId: string) {},
      closePanel(_panelId: string) {},
    }) as any;

  it("blocks formatting commands in read-only mode for non-band selections", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutController();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    };

    const app = {
      isEditing: () => false,
      isReadOnly: () => true,
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("format.toggleBold", true);
    expect(doc.setRangeFormat).not.toHaveBeenCalled();
  });

  it("allows formatting commands in read-only mode when selection is a full row/column band", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutController();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    };

    const app = {
      isEditing: () => false,
      isReadOnly: () => true,
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      // Full column A within the current grid limits (10k rows).
      getSelectionRanges: () => [{ startRow: 0, endRow: 9_999, startCol: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("format.toggleBold", true);
    expect(doc.setRangeFormat).toHaveBeenCalled();
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

  it("view.toggleSplitView no-ops when the desktop shell reports editing via the global edit flag", async () => {
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

    const app = {
      isEditing: () => false,
      focus: vi.fn(),
    } as any;

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      registerBuiltinCommands({ commandRegistry, app, layoutController });
      await commandRegistry.executeCommand("view.toggleSplitView");
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(layoutController.setSplitDirection).not.toHaveBeenCalled();
    expect(app.focus).not.toHaveBeenCalled();
  });
});

describe("registerBuiltinCommands: core editing/view/audit commands", () => {
  it("no-ops view/audit commands when the desktop shell reports editing via the global edit flag", async () => {
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
      isEditing: vi.fn(() => false),
      setShowFormulas: vi.fn(),
      toggleShowFormulas: vi.fn(),
      clearAuditing: vi.fn(),
      toggleAuditingPrecedents: vi.fn(),
      toggleAuditingDependents: vi.fn(),
      toggleAuditingTransitive: vi.fn(),
      focus: vi.fn(),
    } as any;

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      registerBuiltinCommands({ commandRegistry, app, layoutController });

      await commandRegistry.executeCommand("view.toggleShowFormulas");
      await commandRegistry.executeCommand("view.toggleShowFormulas", true);
      await commandRegistry.executeCommand("audit.togglePrecedents");
      await commandRegistry.executeCommand("audit.toggleDependents");
      await commandRegistry.executeCommand("audit.tracePrecedents");
      await commandRegistry.executeCommand("audit.traceDependents");
      await commandRegistry.executeCommand("audit.traceBoth");
      await commandRegistry.executeCommand("audit.clearAuditing");
      await commandRegistry.executeCommand("audit.toggleTransitive");
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(app.setShowFormulas).not.toHaveBeenCalled();
    expect(app.toggleShowFormulas).not.toHaveBeenCalled();
    expect(app.clearAuditing).not.toHaveBeenCalled();
    expect(app.toggleAuditingPrecedents).not.toHaveBeenCalled();
    expect(app.toggleAuditingDependents).not.toHaveBeenCalled();
    expect(app.toggleAuditingTransitive).not.toHaveBeenCalled();
    expect(app.focus).not.toHaveBeenCalled();
  });

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
      focus: vi.fn(),
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
    // This ribbon id is an alias for the canonical audit command; keep it hidden from the command palette
    // to avoid duplicate entries.
    expect(commandRegistry.getCommand("formulas.formulaAuditing.tracePrecedents")?.when).toBe("false");

    await commandRegistry.executeCommand("formulas.formulaAuditing.tracePrecedents");
    expect(calls).toEqual(["clearAuditing", "toggleAuditingPrecedents", "focus"]);
  });

  it("executes formulas.formulaAuditing.traceDependents by clearing auditing then toggling dependents", async () => {
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
      toggleAuditingDependents: vi.fn(() => calls.push("toggleAuditingDependents")),
      focus: vi.fn(() => calls.push("focus")),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    // Coverage: ensure the ribbon command id is registered.
    expect(commandRegistry.getCommand("formulas.formulaAuditing.traceDependents")).toBeDefined();
    // This ribbon id is an alias for the canonical audit command; keep it hidden from the command palette
    // to avoid duplicate entries.
    expect(commandRegistry.getCommand("formulas.formulaAuditing.traceDependents")?.when).toBe("false");

    await commandRegistry.executeCommand("formulas.formulaAuditing.traceDependents");
    expect(calls).toEqual(["clearAuditing", "toggleAuditingDependents", "focus"]);
  });

  it("executes audit.toggleTransitive", async () => {
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
      isEditing: vi.fn(() => false),
      toggleAuditingTransitive: vi.fn(),
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("audit.toggleTransitive");
    expect(app.toggleAuditingTransitive).toHaveBeenCalledTimes(1);
    expect(app.focus).toHaveBeenCalledTimes(1);
  });

  it("executes view.split* commands", async () => {
    const commandRegistry = new CommandRegistry();
    const setSplitDirection = vi.fn();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
      setSplitDirection,
    } as any;

    const app = {
      focus: vi.fn(),
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    await commandRegistry.executeCommand("view.splitVertical");
    await commandRegistry.executeCommand("view.splitHorizontal");
    await commandRegistry.executeCommand("view.splitNone");

    expect(setSplitDirection).toHaveBeenNthCalledWith(1, "vertical", 0.5);
    expect(setSplitDirection).toHaveBeenNthCalledWith(2, "horizontal", 0.5);
    expect(setSplitDirection).toHaveBeenNthCalledWith(3, "none", 0.5);
    expect(app.focus).toHaveBeenCalledTimes(3);
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
  it("registers canonical theme commands and hides ribbon aliases from the command palette", async () => {
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

    // Canonical commands (used by keybindings + command palette).
    await commandRegistry.executeCommand("view.theme.dark");
    expect(themeController.setThemePreference).toHaveBeenNthCalledWith(1, "dark");
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(1);
    expect(app.focus).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("view.theme.highContrast");
    expect(themeController.setThemePreference).toHaveBeenNthCalledWith(2, "high-contrast");
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(2);
    expect(app.focus).toHaveBeenCalledTimes(2);

    // Ribbon ids are still registered for schema compatibility, but hidden from the palette.
    await commandRegistry.executeCommand("view.appearance.theme.dark");
    expect(themeController.setThemePreference).toHaveBeenNthCalledWith(3, "dark");
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(3);
    expect(app.focus).toHaveBeenCalledTimes(3);

    // Canonical commands should be visible in the command palette.
    expect(commandRegistry.getCommand("view.theme.dark")).toMatchObject({
      commandId: "view.theme.dark",
      category: "View",
    });
    expect(commandRegistry.getCommand("view.theme.dark")?.keywords).toEqual(
      expect.arrayContaining(["theme", "dark"]),
    );

    expect(commandRegistry.getCommand("view.theme.dark")?.when).toBeNull();
    expect(commandRegistry.getCommand("view.appearance.theme.dark")?.when).toBe("false");
  });
});
