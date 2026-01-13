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
    const pageLayoutHandlers = {
      openPageSetupDialog: vi.fn(),
      updatePageSetup: vi.fn(),
      setPrintArea: vi.fn(),
      clearPrintArea: vi.fn(),
      addToPrintArea: vi.fn(),
      exportPdf: vi.fn(),
    };

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
      getActiveCellIndentLevel: () => 0,
      openFormatCells,
      showQuickPick: async () => null,
      findReplace: { openFind, openReplace, openGoTo },
      pageLayoutHandlers,
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
    // From registerPageLayoutCommands (when enabled via registerDesktopCommands param)
    expect(commandRegistry.getCommand("pageLayout.pageSetup.pageSetupDialog")).toBeTruthy();

    await commandRegistry.executeCommand("workbench.showCommandPalette");
    expect(openCommandPalette).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("edit.find");
    expect(openFind).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("format.openFormatCells");
    expect(openFormatCells).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("workbench.setAutoSaveEnabled", true);
    expect(handlers.setAutoSaveEnabled).toHaveBeenCalledWith(true);

    await commandRegistry.executeCommand("pageLayout.pageSetup.pageSetupDialog");
    expect(pageLayoutHandlers.openPageSetupDialog).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("pageLayout.pageSetup.margins.normal");
    expect(pageLayoutHandlers.updatePageSetup).toHaveBeenCalledTimes(1);
    const patch = pageLayoutHandlers.updatePageSetup.mock.calls[0]?.[0];
    expect(typeof patch).toBe("function");

    await commandRegistry.executeCommand("pageLayout.printArea.setPrintArea");
    expect(pageLayoutHandlers.setPrintArea).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("pageLayout.export.exportPdf");
    expect(pageLayoutHandlers.exportPdf).toHaveBeenCalledTimes(1);

    // Ensure we didn't accidentally override registerBuiltinCommands' richer formatting command
    // registrations (which include keywords and accept pressed-state args for ribbon toggles).
    expect(commandRegistry.getCommand("format.toggleBold")?.keywords).toEqual(expect.arrayContaining(["bold"]));

    await commandRegistry.executeCommand("format.toggleStrikethrough", true);
    expect(applyFormattingToSelection).toHaveBeenCalled();
  });
});
