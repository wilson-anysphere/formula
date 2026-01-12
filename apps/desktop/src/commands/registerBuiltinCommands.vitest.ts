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
    expect(getPanelPlacement(layoutController.layout, PanelIds.BRANCH_MANAGER)).toEqual({ kind: "docked", side: "left" });

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
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    for (const id of [
      "edit.undo",
      "edit.redo",
      "view.toggleShowFormulas",
      "audit.togglePrecedents",
      "audit.toggleDependents",
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
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);

    // When editing, these commands should no-op (Excel-like behavior).
    app.isEditing.mockReturnValue(true);
    await commandRegistry.executeCommand("view.toggleShowFormulas");
    await commandRegistry.executeCommand("audit.togglePrecedents");
    await commandRegistry.executeCommand("audit.toggleDependents");
    expect(app.toggleShowFormulas).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingPrecedents).toHaveBeenCalledTimes(1);
    expect(app.toggleAuditingDependents).toHaveBeenCalledTimes(1);
  });
});
