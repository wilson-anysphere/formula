import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../sort-filter/sortSelection.js", () => ({
  sortSelection: vi.fn(),
}));

vi.mock("../sort-filter/openCustomSortDialog.js", () => ({
  openCustomSortDialog: vi.fn(),
}));

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { openCustomSortDialog } from "../sort-filter/openCustomSortDialog.js";
import { sortSelection } from "../sort-filter/sortSelection.js";

import { registerSortFilterCommands, SORT_FILTER_RIBBON_COMMANDS } from "./registerSortFilterCommands.js";

describe("registerSortFilterCommands", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("wires Sort A to Z / Sort Z to A commands when not editing", async () => {
    const commandRegistry = new CommandRegistry();
    const app = {} as any;

    registerSortFilterCommands({ commandRegistry, app, isEditing: () => false });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortAtoZ);
    expect(vi.mocked(sortSelection)).toHaveBeenLastCalledWith(app, { order: "ascending" });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortZtoA);
    expect(vi.mocked(sortSelection)).toHaveBeenLastCalledWith(app, { order: "descending" });
  });

  it("does not execute sort commands while editing (uses provided isEditing)", async () => {
    const commandRegistry = new CommandRegistry();
    const app = {} as any;

    registerSortFilterCommands({ commandRegistry, app, isEditing: () => true });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortAtoZ);
    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortZtoA);

    expect(sortSelection).not.toHaveBeenCalled();
  });

  it("does not execute sort commands while editing (split-view secondary editor via global flag)", async () => {
    const commandRegistry = new CommandRegistry();
    const app = {} as any;

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      registerSortFilterCommands({ commandRegistry, app });
      await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortAtoZ);
      await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.sortZtoA);
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(sortSelection).not.toHaveBeenCalled();
  });

  it("does not open the custom sort dialog while editing", async () => {
    const commandRegistry = new CommandRegistry();
    const app = {} as any;

    registerSortFilterCommands({ commandRegistry, app, isEditing: () => true });

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.homeCustomSort);
    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.dataCustomSort);

    expect(openCustomSortDialog).not.toHaveBeenCalled();
  });

  it("passes the provided isEditing predicate to the custom sort dialog host", async () => {
    const commandRegistry = new CommandRegistry();
    // `openCustomSortDialog` is mocked, so these methods are never invoked, but keep them present
    // to reflect the actual host contract.
    const app = {
      getDocument: () => ({}),
      getCurrentSheetId: () => "Sheet1",
      getSelectionRanges: () => [],
      getCellComputedValueForSheet: () => null,
      focus: () => {},
    } as any;

    const isEditing = vi.fn(() => false);
    registerSortFilterCommands({ commandRegistry, app, isEditing });

    // Home ribbon id is an alias; keep it registered but hidden from the command palette.
    expect(commandRegistry.getCommand(SORT_FILTER_RIBBON_COMMANDS.homeCustomSort)?.when).toBe("false");
    // Data tab id is treated as canonical and should remain visible.
    expect(commandRegistry.getCommand(SORT_FILTER_RIBBON_COMMANDS.dataCustomSort)?.when).toBeNull();

    await commandRegistry.executeCommand(SORT_FILTER_RIBBON_COMMANDS.homeCustomSort);

    expect(openCustomSortDialog).toHaveBeenCalledTimes(1);
    expect(isEditing).toHaveBeenCalled();

    const host = vi.mocked(openCustomSortDialog).mock.calls[0]?.[0] as any;
    expect(typeof host?.isEditing).toBe("function");
    expect(host.isEditing()).toBe(false);
    expect(isEditing).toHaveBeenCalledTimes(2);
  });
});
