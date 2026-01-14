import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { promptAndApplyCustomNumberFormat } from "../promptCustomNumberFormat.js";

describe("promptAndApplyCustomNumberFormat (ribbon)", () => {
  it("does nothing while editing (does not prompt)", async () => {
    const showInputBox = vi.fn().mockResolvedValue("0.00");
    const applyFormattingToSelection = vi.fn();

    await promptAndApplyCustomNumberFormat({
      isEditing: () => true,
      showInputBox,
      getActiveCellNumberFormat: () => null,
      applyFormattingToSelection,
    });

    expect(showInputBox).not.toHaveBeenCalled();
    expect(applyFormattingToSelection).not.toHaveBeenCalled();
  });

  it("does not apply if editing starts while the prompt is open", async () => {
    let editing = false;
    const showInputBox = vi.fn().mockImplementation(async () => {
      editing = true;
      return "0.00";
    });
    const applyFormattingToSelection = vi.fn();

    await promptAndApplyCustomNumberFormat({
      isEditing: () => editing,
      showInputBox,
      getActiveCellNumberFormat: () => null,
      applyFormattingToSelection,
    });

    expect(showInputBox).toHaveBeenCalledTimes(1);
    expect(applyFormattingToSelection).not.toHaveBeenCalled();
  });

  it("applies the provided number format code", async () => {
    const doc = new DocumentController();
    const showInputBox = vi.fn().mockResolvedValue("0.00");
    const applyFormattingToSelection = vi.fn((_, fn) => {
      fn(doc, "Sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    });

    await promptAndApplyCustomNumberFormat({
      isEditing: () => false,
      showInputBox,
      getActiveCellNumberFormat: () => null,
      applyFormattingToSelection,
    });

    expect(showInputBox).toHaveBeenCalledTimes(1);
    expect(applyFormattingToSelection).toHaveBeenCalledTimes(1);
    expect(doc.getCellFormat("Sheet1", "A1").numberFormat).toBe("0.00");
  });

  it("clears numberFormat when the input is empty", async () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { numberFormat: "0.00" });

    const showInputBox = vi.fn().mockResolvedValue("");
    const applyFormattingToSelection = vi.fn((_, fn) => {
      fn(doc, "Sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    });

    await promptAndApplyCustomNumberFormat({
      isEditing: () => false,
      showInputBox,
      getActiveCellNumberFormat: () => doc.getCellFormat("Sheet1", "A1").numberFormat ?? null,
      applyFormattingToSelection,
    });

    // Seeded with the current active cell's format.
    expect(showInputBox).toHaveBeenCalledWith(expect.objectContaining({ value: "0.00" }));
    expect(doc.getCellFormat("Sheet1", "A1").numberFormat).toBeNull();
  });

  it("preserves whitespace in the provided number format code", async () => {
    const doc = new DocumentController();
    const showInputBox = vi.fn().mockResolvedValue("0.00 ");
    const applyFormattingToSelection = vi.fn((_, fn) => {
      fn(doc, "Sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    });

    await promptAndApplyCustomNumberFormat({
      isEditing: () => false,
      showInputBox,
      getActiveCellNumberFormat: () => null,
      applyFormattingToSelection,
    });

    expect(doc.getCellFormat("Sheet1", "A1").numberFormat).toBe("0.00 ");
  });
});
