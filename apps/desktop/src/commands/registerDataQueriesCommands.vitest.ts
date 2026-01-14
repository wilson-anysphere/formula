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
} from "../layout/layoutState.js";
import { PanelIds, panelRegistry } from "../panels/panelRegistry.js";

import { DATA_QUERIES_RIBBON_COMMANDS, registerDataQueriesCommands } from "./registerDataQueriesCommands.js";

describe("registerDataQueriesCommands", () => {
  const createLayoutHarness = () => {
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
    return layoutController;
  };

  it("registers Data → Queries & Connections ribbon commands", () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: () => {},
    });

    for (const id of Object.values(DATA_QUERIES_RIBBON_COMMANDS)) {
      expect(commandRegistry.getCommand(id)).toBeDefined();
    }
  });

  it("toggles the Data Queries panel when executing the ribbon toggle command", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: () => {},
    });

    expect(commandRegistry.getCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections)).toBeDefined();
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("closed");

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).not.toBe("closed");

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, false);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("closed");
  });

  it("restores minimized floating panels when opening the Data Queries panel", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    // Place the panel in floating+minimized state.
    layoutController.layout = floatPanel(layoutController.layout, PanelIds.DATA_QUERIES, { x: 10, y: 10, width: 300, height: 200 });
    layoutController.layout = setFloatingPanelMinimized(layoutController.layout, PanelIds.DATA_QUERIES, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("floating");
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(true);

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: () => {},
    });

    // Explicit open (ribbon toggle pressed = true).
    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, true);
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(false);

    // Command palette toggle should also treat minimized panels as "closed" and restore them.
    layoutController.layout = setFloatingPanelMinimized(layoutController.layout, PanelIds.DATA_QUERIES, true);
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(true);
    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections);
    expect(layoutController.layout.floating?.[PanelIds.DATA_QUERIES]?.minimized).toBe(false);
  });

  it("restores collapsed docks when toggling the Data Queries panel via command palette", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: () => {},
    });

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES)).toEqual({ kind: "docked", side: "right" });

    layoutController.layout = setDockCollapsed(layoutController.layout, "right", true);
    expect(layoutController.layout.docks.right.collapsed).toBe(true);

    // No explicit pressed state (command palette toggle) should treat collapsed docks as "closed"
    // and restore them.
    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES)).toEqual({ kind: "docked", side: "right" });
    expect(layoutController.layout.docks.right.collapsed).toBe(false);
  });

  it("syncs ribbon state + restores focus when the toggle command cannot run (missing layout controller)", async () => {
    const commandRegistry = new CommandRegistry();
    const refreshRibbonUiState = vi.fn();
    const focusAfterExecute = vi.fn();
    const showToast = vi.fn();

    registerDataQueriesCommands({
      commandRegistry,
      layoutController: null,
      getPowerQueryService: () => null,
      showToast,
      notify: () => {},
      refreshRibbonUiState,
      focusAfterExecute,
    });

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, true);
    expect(showToast).toHaveBeenCalled();
    expect(refreshRibbonUiState).toHaveBeenCalledTimes(1);
    expect(focusAfterExecute).toHaveBeenCalledTimes(1);
  });

  it("invokes powerQueryService.refreshAll() when executing refresh commands", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    const focusAfterExecute = vi.fn();
    const refreshAll = vi.fn(() => ({ promise: Promise.resolve() }));
    const service = {
      ready: Promise.resolve(),
      getQueries: () => [{ id: "q1" }],
      refreshAll,
    };

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => service as any,
      showToast: () => {},
      notify: () => {},
      focusAfterExecute,
    });

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.refreshAll);
    expect(focusAfterExecute).toHaveBeenCalledTimes(1);
    // Refresh is kicked off in an async continuation; flush microtasks.
    await Promise.resolve();
    await Promise.resolve();
    expect(refreshAll).toHaveBeenCalledTimes(1);
  });

  it("does not execute refresh commands while the spreadsheet is editing (split-view secondary editor via global flag)", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

    const focusAfterExecute = vi.fn();
    const refreshAll = vi.fn(() => ({ promise: Promise.resolve() }));
    const service = {
      ready: Promise.resolve(),
      getQueries: () => [{ id: "q1" }],
      refreshAll,
    };

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => service as any,
      showToast: () => {},
      notify: () => {},
      focusAfterExecute,
    });

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.refreshAll);
      // Flush microtasks in case a refresh job was queued.
      await Promise.resolve();
      await Promise.resolve();
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(refreshAll).not.toHaveBeenCalled();
    expect(focusAfterExecute).not.toHaveBeenCalled();
  });
});
