import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { registerFormatFontDropdownCommands } from "../../commands/registerFormatFontDropdownCommands.js";
import { resolveHomeEditingClearCommandTarget } from "../homeEditingClearCommandRouting.js";

describe("resolveHomeEditingClearCommandTarget", () => {
  it.each([
    { commandId: "home.editing.clear.clearAll", target: "format.clearAll" },
    { commandId: "home.editing.clear.clearFormats", target: "format.clearFormats" },
    { commandId: "home.editing.clear.clearContents", target: "format.clearContents" },
  ])("maps $commandId → $target", (tc) => {
    expect(resolveHomeEditingClearCommandTarget(tc.commandId)).toBe(tc.target);
  });

  it("returns null for unimplemented clear variants", () => {
    expect(resolveHomeEditingClearCommandTarget("home.editing.clear.clearComments")).toBeNull();
    expect(resolveHomeEditingClearCommandTarget("home.editing.clear.clearHyperlinks")).toBeNull();
  });

  it("routes to the underlying DocumentController mutations via format.clear* commands", async () => {
    const registry = new CommandRegistry();

    const clearRange = vi.fn();
    const setRangeFormat = vi.fn().mockReturnValue(true);
    const doc = { clearRange, setRangeFormat } as any;

    const range = { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } };

    const applyFormattingToSelection = vi.fn((label: string, fn: any, options?: any) => {
      fn(doc, "Sheet1", [range]);
      return { label, options };
    });

    registerFormatFontDropdownCommands({
      commandRegistry: registry,
      category: "Format",
      applyFormattingToSelection: applyFormattingToSelection as any,
    });

    // Clear formats → setRangeFormat(..., null)
    clearRange.mockClear();
    setRangeFormat.mockClear();
    await registry.executeCommand(resolveHomeEditingClearCommandTarget("home.editing.clear.clearFormats")!);
    expect(clearRange).not.toHaveBeenCalled();
    expect(setRangeFormat).toHaveBeenCalledWith("Sheet1", range, null, { label: "Clear formats" });

    // Clear contents → clearRange(...)
    clearRange.mockClear();
    setRangeFormat.mockClear();
    await registry.executeCommand(resolveHomeEditingClearCommandTarget("home.editing.clear.clearContents")!);
    expect(clearRange).toHaveBeenCalledWith("Sheet1", range, { label: "Clear contents" });
    expect(setRangeFormat).not.toHaveBeenCalled();

    // Clear all → clearRange(...) + setRangeFormat(..., null)
    clearRange.mockClear();
    setRangeFormat.mockClear();
    await registry.executeCommand(resolveHomeEditingClearCommandTarget("home.editing.clear.clearAll")!);
    expect(clearRange).toHaveBeenCalledWith("Sheet1", range, { label: "Clear all" });
    expect(setRangeFormat).toHaveBeenCalledWith("Sheet1", range, null, { label: "Clear all" });
  });
});

