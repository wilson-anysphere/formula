import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";

describe("registerDesktopCommands", () => {
  it("registers expected desktop command ids and wires representative handlers", async () => {
    const commandRegistry = new CommandRegistry();

    const openCommandPalette = vi.fn();
    const openFind = vi.fn();
    const openReplace = vi.fn();
    const openGoTo = vi.fn();
    const openFormatCells = vi.fn();
    const applyFormattingToSelection = vi.fn();

    const handlers = {
      newWorkbook: vi.fn(),
      openWorkbook: vi.fn(),
      saveWorkbook: vi.fn(),
      saveWorkbookAs: vi.fn(),
      setAutoSaveEnabled: vi.fn(),
      print: vi.fn(),
      printPreview: vi.fn(),
      closeWorkbook: vi.fn(),
      quit: vi.fn(),
    };

    registerDesktopCommands({
      commandRegistry,
      app: {} as any,
      layoutController: { layout: {} as any, openPanel: vi.fn(), closePanel: vi.fn() } as any,
      applyFormattingToSelection,
      getActiveCellNumberFormat: () => "0.00",
      openFormatCells,
      showQuickPick: async () => null,
      findReplace: { openFind, openReplace, openGoTo },
      workbenchFileHandlers: handlers,
      openCommandPalette,
    });

    // From registerBuiltinCommands(...)
    expect(commandRegistry.getCommand("clipboard.copy")).toBeTruthy();
    // From inline registrations moved out of main.ts
    expect(commandRegistry.getCommand("format.toggleBold")).toBeTruthy();
    expect(commandRegistry.getCommand("format.numberFormat.increaseDecimal")).toBeTruthy();
    expect(commandRegistry.getCommand("edit.find")).toBeTruthy();
    // From registerWorkbenchFileCommands(...)
    expect(commandRegistry.getCommand("workbench.saveWorkbook")).toBeTruthy();

    await commandRegistry.executeCommand("workbench.showCommandPalette");
    expect(openCommandPalette).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("edit.find");
    expect(openFind).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("format.openFormatCells");
    expect(openFormatCells).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("workbench.setAutoSaveEnabled", true);
    expect(handlers.setAutoSaveEnabled).toHaveBeenCalledWith(true);

    // Ensure we didn't accidentally override registerBuiltinCommands' richer formatting command
    // registrations (which include keywords and accept pressed-state args for ribbon toggles).
    expect(commandRegistry.getCommand("format.toggleBold")?.keywords).toEqual(expect.arrayContaining(["bold"]));

    await commandRegistry.executeCommand("format.toggleStrikethrough", true);
    expect(applyFormattingToSelection).toHaveBeenCalled();
  });
});
