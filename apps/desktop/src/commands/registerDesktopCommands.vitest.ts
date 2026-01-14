import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";

describe("registerDesktopCommands", () => {
  it("registers expected desktop command ids and wires representative handlers", async () => {
    const commandRegistry = new CommandRegistry();

    const bringSelectedDrawingForward = vi.fn();
    const sendSelectedDrawingBackward = vi.fn();
    const focus = vi.fn();

    const openCommandPalette = vi.fn();
    const openFind = vi.fn();
    const openReplace = vi.fn();
    const openGoTo = vi.fn();
    const openFormatCells = vi.fn();
    const applyFormattingToSelection = vi.fn();

    const formatPainterArm = vi.fn();
    const formatPainterDisarm = vi.fn();
    const formatPainterIsArmed = vi.fn(() => false);

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
      app: { bringSelectedDrawingForward, sendSelectedDrawingBackward, focus } as any,
      layoutController: { layout: {} as any, openPanel: vi.fn(), closePanel: vi.fn() } as any,
      applyFormattingToSelection,
      getActiveCellNumberFormat: () => "0.00",
      getActiveCellIndentLevel: () => 0,
      openFormatCells,
      showQuickPick: async () => null,
      findReplace: { openFind, openReplace, openGoTo },
      formatPainter: {
        isArmed: formatPainterIsArmed,
        arm: formatPainterArm,
        disarm: formatPainterDisarm,
        onCancel: null,
      },
      pageLayoutHandlers,
      workbenchFileHandlers: handlers,
      openCommandPalette,
    });

    // From registerBuiltinCommands(...)
    expect(commandRegistry.getCommand("clipboard.copy")).toBeTruthy();
    // From inline registrations moved out of main.ts
    expect(commandRegistry.getCommand("format.toggleBold")).toBeTruthy();
    expect(commandRegistry.getCommand("format.numberFormat.increaseDecimal")).toBeTruthy();
    expect(commandRegistry.getCommand("format.clearFormats")).toBeTruthy();
    expect(commandRegistry.getCommand("format.toggleFormatPainter")).toBeTruthy();
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

    await commandRegistry.executeCommand("pageLayout.arrange.bringForward");
    expect(bringSelectedDrawingForward).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("pageLayout.arrange.sendBackward");
    expect(sendSelectedDrawingBackward).toHaveBeenCalledTimes(1);

    expect(focus).toHaveBeenCalledTimes(2);

    // Ensure we didn't accidentally override registerBuiltinCommands' richer formatting command
    // registrations (which include keywords and accept pressed-state args for ribbon toggles).
    expect(commandRegistry.getCommand("format.toggleBold")?.keywords).toEqual(expect.arrayContaining(["bold"]));

    await commandRegistry.executeCommand("format.toggleStrikethrough", true);
    expect(applyFormattingToSelection).toHaveBeenCalled();

    await commandRegistry.executeCommand("format.toggleFormatPainter");
    expect(formatPainterArm).toHaveBeenCalledTimes(1);
  });

  it("registers Data → Queries & Connections commands when dataQueriesHandlers are provided", async () => {
    const commandRegistry = new CommandRegistry();

    const focus = vi.fn();
    const showToast = vi.fn();
    const notify = vi.fn(async () => {});

    const refreshAll = vi.fn(() => ({ promise: Promise.resolve() }));
    const service = {
      ready: Promise.resolve(),
      getQueries: () => [{ id: "q1" }],
      refreshAll,
    };

    registerDesktopCommands({
      commandRegistry,
      app: { focus } as any,
      layoutController: null,
      applyFormattingToSelection: () => {},
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
      openFormatCells: () => {},
      showQuickPick: async () => null,
      findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
      workbenchFileHandlers: {
        newWorkbook: () => {},
        openWorkbook: () => {},
        saveWorkbook: () => {},
        saveWorkbookAs: () => {},
        setAutoSaveEnabled: () => {},
        print: () => {},
        printPreview: () => {},
        closeWorkbook: () => {},
        quit: () => {},
      },
      dataQueriesHandlers: {
        getPowerQueryService: () => service as any,
        showToast,
        notify,
        now: () => 0,
        focusAfterExecute: focus,
      },
    });

    expect(commandRegistry.getCommand("data.queriesConnections.refreshAll")).toBeTruthy();

    await commandRegistry.executeCommand("data.queriesConnections.refreshAll");
    expect(focus).toHaveBeenCalledTimes(1);
    // Refresh is kicked off in an async continuation; flush microtasks.
    await Promise.resolve();
    await Promise.resolve();
    expect(refreshAll).toHaveBeenCalledTimes(1);
    expect(showToast).not.toHaveBeenCalled();
    expect(notify).not.toHaveBeenCalled();
  });
});
