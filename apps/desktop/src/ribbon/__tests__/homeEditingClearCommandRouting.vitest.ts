import { describe, expect, it, vi } from "vitest";

import { registerFormatFontDropdownCommands } from "../../commands/registerFormatFontDropdownCommands";
import { CommandRegistry } from "../../extensions/commandRegistry";
import { resolveHomeEditingClearCommandTarget } from "../homeEditingClearCommandRouting";

describe("Home → Editing → Clear command routing", () => {
  it("routes legacy Home editing clear ids to the built-in clear implementations", async () => {
    const registry = new CommandRegistry();

    const clearRange = vi.fn();
    const setRangeFormat = vi.fn().mockReturnValue(true);
    const doc = { clearRange, setRangeFormat } as any;

    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

    const applyFormattingToSelection = vi.fn((label: string, fn: any, options?: any) => {
      fn(doc, "Sheet1", [range]);
      return { label, options };
    });

    registerFormatFontDropdownCommands({
      commandRegistry: registry,
      category: "Format",
      applyFormattingToSelection: applyFormattingToSelection as any,
    });

    // Clear formats routes to `doc.setRangeFormat(..., null)`.
    const clearFormatsTarget = resolveHomeEditingClearCommandTarget("home.editing.clear.clearFormats");
    expect(clearFormatsTarget).toBe("format.clearFormats");
    await registry.executeCommand(clearFormatsTarget!);
    expect(setRangeFormat).toHaveBeenCalledWith("Sheet1", range, null, { label: "Clear formats" });
    expect(clearRange).not.toHaveBeenCalled();

    clearRange.mockClear();
    setRangeFormat.mockClear();

    // Clear contents routes to `doc.clearRange`.
    const clearContentsTarget = resolveHomeEditingClearCommandTarget("home.editing.clear.clearContents");
    expect(clearContentsTarget).toBe("format.clearContents");
    await registry.executeCommand(clearContentsTarget!);
    expect(clearRange).toHaveBeenCalledWith("Sheet1", range, { label: "Clear contents" });
    expect(setRangeFormat).not.toHaveBeenCalled();

    clearRange.mockClear();
    setRangeFormat.mockClear();

    // Clear all routes to both `doc.clearRange` and `doc.setRangeFormat(..., null)`.
    const clearAllTarget = resolveHomeEditingClearCommandTarget("home.editing.clear.clearAll");
    expect(clearAllTarget).toBe("format.clearAll");
    await registry.executeCommand(clearAllTarget!);
    expect(clearRange).toHaveBeenCalledWith("Sheet1", range, { label: "Clear all" });
    expect(setRangeFormat).toHaveBeenCalledWith("Sheet1", range, null, { label: "Clear all" });

    // Unimplemented variants should not resolve to builtin targets.
    expect(resolveHomeEditingClearCommandTarget("home.editing.clear.clearComments")).toBeNull();
    expect(resolveHomeEditingClearCommandTarget("home.editing.clear.clearHyperlinks")).toBeNull();
  });
});

