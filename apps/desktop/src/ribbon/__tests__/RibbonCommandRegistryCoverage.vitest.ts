import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createDefaultLayout } from "../../layout/layoutState.js";
import { registerBuiltinCommands } from "../../commands/registerBuiltinCommands.js";
import { defaultRibbonSchema } from "../ribbonSchema";

describe("Ribbon ↔ CommandRegistry: Home debug buttons", () => {
  it("uses canonical CommandRegistry ids for Auditing / Split view / Freeze", () => {
    const homeTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "home");
    expect(homeTab, "Expected Home tab to exist").toBeTruthy();
    if (!homeTab) return;

    const debugGroupIds = ["home.debug.auditing", "home.debug.split", "home.debug.freeze"];
    const debugButtons = homeTab.groups
      .filter((group) => debugGroupIds.includes(group.id))
      .flatMap((group) => group.buttons);

    expect(debugButtons.length, "Expected Home debug groups to include buttons").toBeGreaterThan(0);

    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel() {},
      closePanel() {},
      setSplitDirection() {},
    } as any;

    const app = {
      isEditing: () => false,
      focus: () => {},
      toggleAuditingPrecedents: () => {},
      toggleAuditingDependents: () => {},
      toggleAuditingTransitive: () => {},
      freezePanes: () => {},
      freezeTopRow: () => {},
      freezeFirstColumn: () => {},
      unfreezePanes: () => {},
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    for (const button of debugButtons) {
      expect(commandRegistry.getCommand(button.id), `Expected command '${button.id}' (from ribbon) to be registered`).toBeDefined();
    }
  });
});

