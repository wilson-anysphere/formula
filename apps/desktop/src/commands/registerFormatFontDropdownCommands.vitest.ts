import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { registerFormatFontDropdownCommands } from "./registerFormatFontDropdownCommands.js";

describe("registerFormatFontDropdownCommands", () => {
  it("executes a border command via CommandRegistry (format.borders.top)", async () => {
    const registry = new CommandRegistry();

    const setRangeFormat = vi.fn().mockReturnValue(true);
    const doc = { setRangeFormat } as any;

    const range = { start: { row: 2, col: 3 }, end: { row: 4, col: 5 } };

    const applyFormattingToSelection = vi.fn((label: string, fn: any, options?: any) => {
      fn(doc, "Sheet1", [range]);
      return { label, options };
    });

    registerFormatFontDropdownCommands({
      commandRegistry: registry,
      category: "Format",
      applyFormattingToSelection: applyFormattingToSelection as any,
    });

    await registry.executeCommand("format.borders.top");

    expect(applyFormattingToSelection).toHaveBeenCalledWith("Borders", expect.any(Function), { forceBatch: true });
    expect(setRangeFormat).toHaveBeenCalledWith(
      "Sheet1",
      { start: { row: 2, col: 3 }, end: { row: 2, col: 5 } },
      { border: { top: { style: "thin", color: "#FF000000" } } },
      { label: "Borders" },
    );
  });

  it("executes a clear command via CommandRegistry (format.clearFormats)", async () => {
    const registry = new CommandRegistry();

    const setRangeFormat = vi.fn().mockReturnValue(true);
    const doc = { setRangeFormat, clearRange: vi.fn() } as any;

    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

    const applyFormattingToSelection = vi.fn((label: string, fn: any) => {
      fn(doc, "Sheet1", [range]);
      return label;
    });

    registerFormatFontDropdownCommands({
      commandRegistry: registry,
      category: "Format",
      applyFormattingToSelection: applyFormattingToSelection as any,
    });

    await registry.executeCommand("format.clearFormats");

    expect(setRangeFormat).toHaveBeenCalledWith("Sheet1", range, null, { label: "Clear formats" });
  });
});
