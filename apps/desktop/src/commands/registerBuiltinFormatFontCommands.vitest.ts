import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { registerBuiltinFormatFontCommands } from "./registerBuiltinFormatFontCommands.js";

describe("registerBuiltinFormatFontCommands", () => {
  it("executes font name + font size preset commands via doc.setRangeFormat patches", async () => {
    const commandRegistry = new CommandRegistry();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    } as any;

    const sheetId = "Sheet1";
    const ranges = [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }];

    const applyFormattingToSelection = vi.fn((_label, fn) => {
      fn(doc, sheetId, ranges);
    });

    registerBuiltinFormatFontCommands({
      commandRegistry,
      applyFormattingToSelection: applyFormattingToSelection as any,
    });

    await commandRegistry.executeCommand("format.fontName.calibri");
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { font: { name: "Calibri" } },
      { label: "Font" },
    );

    doc.setRangeFormat.mockClear();

    await commandRegistry.executeCommand("format.fontSize.12");
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { font: { size: 12 } },
      { label: "Font size" },
    );
  });
});
