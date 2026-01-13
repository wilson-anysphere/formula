import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";

describe("registerDesktopCommands", () => {
  it("registers expected desktop command ids", () => {
    const commandRegistry = new CommandRegistry();

    registerDesktopCommands({
      commandRegistry,
      app: {} as any,
      layoutController: { layout: {} as any, openPanel: vi.fn(), closePanel: vi.fn() } as any,
      applyFormattingToSelection: () => {},
      getActiveCellNumberFormat: () => null,
      openFormatCells: vi.fn(),
      showQuickPick: async () => null,
      findReplace: { openFind: vi.fn(), openReplace: vi.fn(), openGoTo: vi.fn() },
      workbenchFileHandlers: {
        newWorkbook: vi.fn(),
        openWorkbook: vi.fn(),
        saveWorkbook: vi.fn(),
        saveWorkbookAs: vi.fn(),
        setAutoSaveEnabled: vi.fn(),
        print: vi.fn(),
        printPreview: vi.fn(),
        closeWorkbook: vi.fn(),
        quit: vi.fn(),
      },
      openCommandPalette: vi.fn(),
    });

    // From registerBuiltinCommands(...)
    expect(commandRegistry.getCommand("clipboard.copy")).toBeTruthy();
    // From inline registrations moved out of main.ts
    expect(commandRegistry.getCommand("format.toggleBold")).toBeTruthy();
    expect(commandRegistry.getCommand("edit.find")).toBeTruthy();
    // From registerWorkbenchFileCommands(...)
    expect(commandRegistry.getCommand("workbench.saveWorkbook")).toBeTruthy();
  });
});
