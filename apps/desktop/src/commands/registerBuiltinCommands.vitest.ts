import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { createDefaultLayout, getPanelPlacement, openPanel, closePanel } from "../layout/layoutState.js";
import { PanelIds, panelRegistry } from "../panels/panelRegistry.js";

import { registerBuiltinCommands } from "./registerBuiltinCommands.js";

describe("registerBuiltinCommands: panel toggles", () => {
  it("toggles AI Chat panel open/closed", async () => {
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

    expect(getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT).kind).toBe("closed");

    await commandRegistry.executeCommand("view.togglePanel.aiChat");
    expect(getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT)).toEqual({ kind: "docked", side: "right" });

    await commandRegistry.executeCommand("view.togglePanel.aiChat");
    expect(getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT).kind).toBe("closed");
  });

  it("toggles Version History panel open/closed", async () => {
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

    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY).kind).toBe("closed");

    await commandRegistry.executeCommand("view.togglePanel.versionHistory");
    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY)).toEqual({ kind: "docked", side: "right" });

    await commandRegistry.executeCommand("view.togglePanel.versionHistory");
    expect(getPanelPlacement(layoutController.layout, PanelIds.VERSION_HISTORY).kind).toBe("closed");
  });

  it("toggles Branch Manager panel open/closed", async () => {
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
