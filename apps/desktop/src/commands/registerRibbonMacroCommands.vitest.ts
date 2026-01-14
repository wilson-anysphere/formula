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

    // Developer tab macro ids are registered for ribbon coverage, but hidden from the command palette
    // to avoid duplicate entries (View tab exposes the canonical palette-visible commands).
    for (const commandId of [
      "developer.code.macros.run",
      "developer.code.macros.edit",
      "developer.code.recordMacro",
      "developer.code.recordMacro.stop",
      "developer.code.useRelativeReferences",
    ] as const) {
      expect(commandRegistry.getCommand(commandId)?.when).toBe("false");
    }

    // Spot-check a few titles so the command palette matches the ribbon labels.
    expect(commandRegistry.getCommand("view.macros.recordMacro")?.title).toBe("Record Macro…");
    expect(commandRegistry.getCommand("developer.code.macroSecurity")?.title).toBe("Macro Security…");
    expect(commandRegistry.getCommand("developer.code.macroSecurity.trustCenter")?.title).toBe("Trust Center…");
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

  it("does not execute macro commands while the spreadsheet is editing (split-view secondary editor via global flag)", async () => {
    const commandRegistry = new CommandRegistry();

    const openPanel = vi.fn();
    const setPendingMacrosPanelFocus = vi.fn();

    registerRibbonMacroCommands({
      commandRegistry,
      handlers: {
        openPanel,
        focusScriptEditorPanel: vi.fn(),
        focusVbaMigratePanel: vi.fn(),
        setPendingMacrosPanelFocus,
        startMacroRecorder: vi.fn(),
        stopMacroRecorder: vi.fn(),
        isTauri: () => false,
      },
    });

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      await commandRegistry.executeCommand("view.macros.viewMacros.run");
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(openPanel).not.toHaveBeenCalled();
    expect(setPendingMacrosPanelFocus).not.toHaveBeenCalled();
  });
});
