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
