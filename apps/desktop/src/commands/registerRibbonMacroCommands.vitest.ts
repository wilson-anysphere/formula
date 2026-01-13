import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { PanelIds } from "../panels/panelRegistry.js";

import { RIBBON_MACRO_COMMAND_IDS, registerRibbonMacroCommands } from "./registerRibbonMacroCommands.js";

describe("registerRibbonMacroCommands", () => {
  it("registers expected command ids", () => {
    const commandRegistry = new CommandRegistry();

    registerRibbonMacroCommands({
      commandRegistry,
      handlers: {
        openPanel: () => {},
        focusScriptEditorPanel: () => {},
        focusVbaMigratePanel: () => {},
        setPendingMacrosPanelFocus: () => {},
        startMacroRecorder: () => {},
        stopMacroRecorder: () => {},
        isTauri: () => false,
      },
    });

    for (const commandId of RIBBON_MACRO_COMMAND_IDS) {
      expect(commandRegistry.getCommand(commandId)).toBeTruthy();
    }
  });

  it("wires View → Macros → Run to set focus + open the Macros panel", async () => {
    const commandRegistry = new CommandRegistry();

    const openedPanels: string[] = [];
    let pendingFocus: string | null = null;

    registerRibbonMacroCommands({
      commandRegistry,
      handlers: {
        openPanel: (panelId) => {
          openedPanels.push(panelId);
        },
        focusScriptEditorPanel: vi.fn(),
        focusVbaMigratePanel: vi.fn(),
        setPendingMacrosPanelFocus: (target) => {
          pendingFocus = target;
        },
        startMacroRecorder: vi.fn(),
        stopMacroRecorder: vi.fn(),
        isTauri: () => false,
      },
    });

    await commandRegistry.executeCommand("view.macros.viewMacros.run");

    expect(pendingFocus).toBe("runner-run");
    expect(openedPanels).toEqual([PanelIds.MACROS]);
  });
});

