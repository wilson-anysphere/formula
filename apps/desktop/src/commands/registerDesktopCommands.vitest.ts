import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { registerEncryptionUiCommands } from "../collab/encryption-ui/registerEncryptionUiCommands.js";
import { t } from "../i18n/index.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";

describe("registerDesktopCommands", () => {
  it("does not register duplicate command-palette entries (same category + title)", () => {
    const commandRegistry = new CommandRegistry();

    registerDesktopCommands({
      commandRegistry,
      // Commands are registered eagerly but not executed in this test, so provide only
      // the minimal surface area required by registration-time code.
      app: { focus: vi.fn() } as any,
      layoutController: { layout: {} as any, openPanel: vi.fn(), closePanel: vi.fn() } as any,
      // Include optional handler surfaces so this test covers the full set of commands
      // that the real desktop app registers (and can catch duplicates introduced only
      // when optional feature wiring is enabled).
      themeController: { setThemePreference: vi.fn() } as any,
      refreshRibbonUiState: vi.fn(),
      applyFormattingToSelection: vi.fn(),
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
      openFormatCells: vi.fn(),
      showQuickPick: async () => null,
      findReplace: { openFind: vi.fn(), openReplace: vi.fn(), openGoTo: vi.fn() },
      dataQueriesHandlers: {
        getPowerQueryService: () =>
          ({
            ready: Promise.resolve(),
            getQueries: () => [],
            refreshAll: () => ({ promise: Promise.resolve() }),
          }) as any,
        showToast: vi.fn(),
        notify: vi.fn(async () => {}),
        now: () => 0,
        focusAfterExecute: vi.fn(),
      },
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
      pageLayoutHandlers: {
        openPageSetupDialog: vi.fn(),
        updatePageSetup: vi.fn(),
        setPrintArea: vi.fn(),
        clearPrintArea: vi.fn(),
        addToPrintArea: vi.fn(),
        exportPdf: vi.fn(),
      },
      formatPainter: {
        isArmed: () => false,
        arm: vi.fn(),
        disarm: vi.fn(),
        onCancel: null,
      },
      ribbonMacroHandlers: {
        openPanel: vi.fn(),
        focusScriptEditorPanel: vi.fn(),
        focusVbaMigratePanel: vi.fn(),
        setPendingMacrosPanelFocus: vi.fn(),
        startMacroRecorder: vi.fn(),
        stopMacroRecorder: vi.fn(),
        isTauri: () => false,
      },
    });

    // `main.ts` registers some additional desktop-only commands outside `registerDesktopCommands(...)`.
    // Ensure these don't introduce duplicate command palette entries either.
    registerEncryptionUiCommands({ commandRegistry, app: {} as any });
    // A few desktop-only commands are still registered inline in `main.ts`. Mirror their titles/categories
    // here so this regression test covers command palette duplication across the full desktop catalog.
    //
    // NOTE: Keep the command ids passed as variables (not string literals) so the node:test static
    // `commandRegistryBuiltinCommandDuplicates` suite doesn't treat these Vitest registrations as
    // production command registrations under `src/commands/*`.
    const mainOnlyCommands: Array<{ id: string; titleKey: string; categoryKey: string }> = [
      { id: "ui.openContextMenu", titleKey: "command.ui.openContextMenu", categoryKey: "commandCategory.view" },
      { id: "checkForUpdates", titleKey: "commandPalette.command.checkForUpdates", categoryKey: "commandCategory.help" },
      { id: "debugShowSystemNotification", titleKey: "command.debugShowSystemNotification", categoryKey: "commandCategory.debug" },
    ];
    for (const cmd of mainOnlyCommands) {
      commandRegistry.registerBuiltinCommand(cmd.id, t(cmd.titleKey), () => {}, { category: t(cmd.categoryKey) });
    }

    // Treat `when: "false"` as "hidden from context-aware surfaces (command palette)".
    const visible = commandRegistry.listCommands().filter((cmd) => cmd.when !== "false");

    const seen = new Map<string, string[]>();
    for (const cmd of visible) {
      const key = `${cmd.category ?? ""}::${cmd.title}`;
      const list = seen.get(key) ?? [];
      list.push(cmd.commandId);
      seen.set(key, list);
    }

    const duplicates = [...seen.entries()]
      .filter(([, ids]) => ids.length > 1)
      .map(([key, ids]) => ({ key, ids: [...ids].sort() }))
      .sort((a, b) => a.key.localeCompare(b.key));

    expect(duplicates).toEqual([]);
  });

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
      sheetStructureHandlers: {
        openOrganizeSheets: vi.fn(),
        insertSheet: vi.fn(),
        deleteActiveSheet: vi.fn(),
      },
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

  it("wires sheet-structure ribbon command ids when sheetStructureHandlers are provided", async () => {
    const commandRegistry = new CommandRegistry();

    const insertSheet = vi.fn(async () => {});
    const deleteActiveSheet = vi.fn(async () => {});
    const openOrganizeSheets = vi.fn(async () => {});

    registerDesktopCommands({
      commandRegistry,
      app: { isReadOnly: () => false } as any,
      layoutController: null,
      sheetStructureHandlers: { insertSheet, deleteActiveSheet, openOrganizeSheets },
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
    });

    await commandRegistry.executeCommand("home.cells.insert.insertSheet");
    expect(insertSheet).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("home.cells.delete.deleteSheet");
    expect(deleteActiveSheet).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand("home.cells.format.organizeSheets");
    expect(openOrganizeSheets).toHaveBeenCalledTimes(1);
  });

  it("blocks sheet-structure commands when collab session is read-only", async () => {
    const commandRegistry = new CommandRegistry();

    const focus = vi.fn();
    const insertSheet = vi.fn(async () => {});
    const deleteActiveSheet = vi.fn(async () => {});
    const openOrganizeSheets = vi.fn(async () => {});

    registerDesktopCommands({
      commandRegistry,
      app: { isReadOnly: () => true, focus } as any,
      layoutController: null,
      sheetStructureHandlers: { insertSheet, deleteActiveSheet, openOrganizeSheets },
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
    });

    await commandRegistry.executeCommand("home.cells.insert.insertSheet");
    await commandRegistry.executeCommand("home.cells.delete.deleteSheet");
    await commandRegistry.executeCommand("home.cells.format.organizeSheets");

    expect(insertSheet).not.toHaveBeenCalled();
    expect(deleteActiveSheet).not.toHaveBeenCalled();
    expect(openOrganizeSheets).not.toHaveBeenCalled();
    expect(focus).toHaveBeenCalledTimes(3);
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
      sheetStructureHandlers: {
        openOrganizeSheets: () => {},
        insertSheet: () => {},
        deleteActiveSheet: () => {},
      },
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
