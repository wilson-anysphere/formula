import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { createDefaultLayout, getPanelPlacement, openPanel, closePanel } from "../layout/layoutState.js";
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

    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("closed");

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, true);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).not.toBe("closed");

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections, false);
    expect(getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES).kind).toBe("closed");
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
});
