import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { registerFormatAlignmentCommands } from "./registerFormatAlignmentCommands.js";

describe("registerFormatAlignmentCommands", () => {
  it("format.alignLeft applies a horizontal alignment patch via doc.setRangeFormat", async () => {
    const commandRegistry = new CommandRegistry();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    };

    const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

    registerFormatAlignmentCommands({
      commandRegistry,
      applyFormattingToSelection: (_label, fn) => {
        fn(doc, "sheet1", [range]);
      },
      activeCellIndentLevel: () => 0,
      openAlignmentDialog: () => {},
    });

    await commandRegistry.executeCommand("format.alignLeft");

    expect(doc.setRangeFormat).toHaveBeenCalledTimes(1);
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      "sheet1",
      range,
      { alignment: { horizontal: "left" } },
      { label: "Horizontal align" },
    );
  });
});

