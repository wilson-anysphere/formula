import { describe, expect, it, vi } from "vitest";

vi.mock("../sort-filter/sortSelection.js", () => ({
  sortSelection: vi.fn(),
}));

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbon/ribbonCommandRegistryDisabling.js";
import { sortSelection } from "../sort-filter/sortSelection.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";
import { SORT_FILTER_RIBBON_COMMANDS } from "./registerSortFilterCommands.js";

describe("Sort & Filter ribbon commands", () => {
  it("registers sort A→Z / Z→A commands and wires execution through sortSelection", async () => {
    const commandRegistry = new CommandRegistry();
    const app = {} as any;

    // Baseline: ids are disabled by default when unregistered.
    const baselineDisabled = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    for (const id of Object.values(SORT_FILTER_RIBBON_COMMANDS)) {
      expect(baselineDisabled[id]).toBe(true);
    }

    registerDesktopCommands({
      commandRegistry,
      app,
      layoutController: null,
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

    // Ensure commands are registered (prevents ribbon auto-disable).
    for (const id of Object.values(SORT_FILTER_RIBBON_COMMANDS)) {
      expect(commandRegistry.getCommand(id)).toBeTruthy();
    }

    // And computeRibbonDisabledByIdFromCommandRegistry keeps them enabled.
    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    for (const id of Object.values(SORT_FILTER_RIBBON_COMMANDS)) {
      expect(disabledById[id]).toBeUndefined();
    }

    const sortSelectionMock = vi.mocked(sortSelection);

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.homeSortAtoZ);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "ascending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.homeSortZtoA);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "descending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.dataSortAtoZ);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "ascending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.dataSortZtoA);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "descending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.dataDropdownSortAtoZ);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "ascending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.dataDropdownSortZtoA);
    expect(sortSelectionMock).toHaveBeenLastCalledWith(app, { order: "descending" });
  });
});
