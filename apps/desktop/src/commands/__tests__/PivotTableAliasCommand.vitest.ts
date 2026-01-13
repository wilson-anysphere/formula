// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createDefaultLayout, openPanel as openDockPanel } from "../../layout/layoutState.js";
import { panelRegistry, PanelIds } from "../../panels/panelRegistry.js";
import { registerBuiltinCommands } from "../registerBuiltinCommands.js";

describe("PivotTable ribbon alias commands", () => {
  it("routes insert.tables.pivotTable.fromTableRange through view.insertPivotTable", async () => {
    const commandRegistry = new CommandRegistry();
    const openPanelSpy = vi.fn();

    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        openPanelSpy(panelId);
        this.layout = openDockPanel(this.layout, panelId, { panelRegistry });
      },
      setFloatingPanelMinimized: () => {},
    } as any;

    registerBuiltinCommands({
      commandRegistry,
      app: {} as any,
      layoutController,
    });

    const selectionEventListener = vi.fn();
    window.addEventListener("pivot-builder:use-selection", selectionEventListener);

    try {
      await commandRegistry.executeCommand("insert.tables.pivotTable.fromTableRange");
    } finally {
      window.removeEventListener("pivot-builder:use-selection", selectionEventListener);
    }

    expect(openPanelSpy).toHaveBeenCalledTimes(1);
    expect(openPanelSpy).toHaveBeenCalledWith(PanelIds.PIVOT_BUILDER);
    expect(selectionEventListener).toHaveBeenCalledTimes(1);
  });
});
