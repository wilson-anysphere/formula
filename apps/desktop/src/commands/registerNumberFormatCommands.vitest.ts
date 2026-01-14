import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { registerNumberFormatCommands } from "./registerNumberFormatCommands.js";

describe("registerNumberFormatCommands", () => {
  it("executes number-format commands and applies the expected setRangeFormat patches", async () => {
    const commandRegistry = new CommandRegistry();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    };

    const sheetId = "Sheet1";
    const ranges = [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }];

    const applyFormattingToSelection = (
      _label: string,
      fn: (doc: any, sheetId: string, ranges: any[]) => void | boolean,
    ) => {
      fn(doc as any, sheetId, ranges);
    };

    registerNumberFormatCommands({
      commandRegistry,
      applyFormattingToSelection,
      getActiveCellNumberFormat: () => "0.00",
      t: (key) => key,
      category: "Format",
    });

    await commandRegistry.executeCommand("format.numberFormat.general");
    await commandRegistry.executeCommand("format.numberFormat.number");
    await commandRegistry.executeCommand("format.numberFormat.currency");
    await commandRegistry.executeCommand("format.numberFormat.longDate");
    await commandRegistry.executeCommand("format.numberFormat.time");
    await commandRegistry.executeCommand("format.numberFormat.increaseDecimal");

    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: null },
      { label: "Number format" },
    );
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: "0.00" },
      { label: "Number format" },
    );
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: "$#,##0.00" },
      { label: "Number format" },
    );
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: "yyyy-mm-dd" },
      { label: "Number format" },
    );
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: "h:mm:ss" },
      { label: "Number format" },
    );
    expect(doc.setRangeFormat).toHaveBeenCalledWith(
      sheetId,
      ranges[0],
      { numberFormat: "0.000" },
      { label: "Number format" },
    );
  });

  it("treats time-only formats as non-adjustable for decimal stepping commands", async () => {
    const commandRegistry = new CommandRegistry();

    const doc = {
      setRangeFormat: vi.fn(() => true),
    };

    const sheetId = "Sheet1";
    const ranges = [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }];

    const applyFormattingToSelection = (
      _label: string,
      fn: (doc: any, sheetId: string, ranges: any[]) => void | boolean,
    ) => {
      fn(doc as any, sheetId, ranges);
    };

    registerNumberFormatCommands({
      commandRegistry,
      applyFormattingToSelection,
      getActiveCellNumberFormat: () => "h:mm:ss",
      t: (key) => key,
      category: "Format",
    });

    await commandRegistry.executeCommand("format.numberFormat.increaseDecimal");
    await commandRegistry.executeCommand("format.numberFormat.decreaseDecimal");

    expect(doc.setRangeFormat).not.toHaveBeenCalled();
  });
});
