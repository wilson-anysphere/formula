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

  it("invokes powerQueryService.refreshAll() when executing refresh commands", async () => {
    const commandRegistry = new CommandRegistry();
    const layoutController = createLayoutHarness();

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
    });

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.refreshAll);
    // Refresh is kicked off in an async continuation; flush microtasks.
    await Promise.resolve();
    await Promise.resolve();
    expect(refreshAll).toHaveBeenCalledTimes(1);
  });
});

