// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { PanelIds } from "../../panels/panelRegistry.js";
import { registerBuiltinCommands } from "../registerBuiltinCommands.js";

describe("PivotTable ribbon alias commands", () => {
  it("routes insert.tables.pivotTable.fromTableRange through view.insertPivotTable", async () => {
    const commandRegistry = new CommandRegistry();
    const openPanel = vi.fn();

    registerBuiltinCommands({
      commandRegistry,
      app: {} as any,
      layoutController: { openPanel } as any,
    });

    const selectionEventListener = vi.fn();
    window.addEventListener("pivot-builder:use-selection", selectionEventListener);

    try {
      await commandRegistry.executeCommand("insert.tables.pivotTable.fromTableRange");
    } finally {
      window.removeEventListener("pivot-builder:use-selection", selectionEventListener);
    }

    expect(openPanel).toHaveBeenCalledTimes(1);
    expect(openPanel).toHaveBeenCalledWith(PanelIds.PIVOT_BUILDER);
    expect(selectionEventListener).toHaveBeenCalledTimes(1);
  });
});

