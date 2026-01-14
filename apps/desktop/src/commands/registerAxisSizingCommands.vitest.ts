import { describe, expect, it, vi } from "vitest";

vi.mock("../ribbon/axisSizing.js", async () => {
  const actual = await vi.importActual<typeof import("../ribbon/axisSizing.js")>("../ribbon/axisSizing.js");
  return {
    ...actual,
    promptAndApplyAxisSizing: vi.fn(async () => {}),
  };
});

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { promptAndApplyAxisSizing } from "../ribbon/axisSizing.js";
import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbon/ribbonCommandRegistryDisabling.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";

describe("axis sizing ribbon commands", () => {
  it("registers Home → Cells → Format sizing commands and keeps them enabled in ribbon auto-disable logic", async () => {
    const commandRegistry = new CommandRegistry();

    // Baseline: these ribbon menu items should be disabled when unregistered.
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    expect(baselineDisabledById["home.cells.format.rowHeight"]).toBe(true);
    expect(baselineDisabledById["home.cells.format.columnWidth"]).toBe(true);

    const app = {} as any;
    const isEditing = vi.fn(() => false);

    registerDesktopCommands({
      commandRegistry,
      app,
      layoutController: { layout: {} as any, openPanel: vi.fn(), closePanel: vi.fn() } as any,
      isEditing,
      applyFormattingToSelection: vi.fn(),
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
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
    });

    expect(commandRegistry.getCommand("home.cells.format.rowHeight")).toBeTruthy();
    expect(commandRegistry.getCommand("home.cells.format.columnWidth")).toBeTruthy();

    await commandRegistry.executeCommand("home.cells.format.rowHeight");
    expect(vi.mocked(promptAndApplyAxisSizing)).toHaveBeenCalledWith(app, "rowHeight", { isEditing });

    await commandRegistry.executeCommand("home.cells.format.columnWidth");
    expect(vi.mocked(promptAndApplyAxisSizing)).toHaveBeenCalledWith(app, "colWidth", { isEditing });

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    expect(disabledById["home.cells.format.rowHeight"]).toBeUndefined();
    expect(disabledById["home.cells.format.columnWidth"]).toBeUndefined();
  });
});

