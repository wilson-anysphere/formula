import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { handleRibbonCommand, type RibbonCommandHandlerContext } from "../commandHandlers.js";

function createCtx(
  doc: DocumentController,
  options: {
    selection?: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>;
    isEditing?: boolean;
  } = {},
): RibbonCommandHandlerContext {
  return {
    app: {
      getDocument: () => doc,
      getCurrentSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => options.selection ?? [],
      focus: () => {
        // no-op for tests
      },
    },
    isEditing: () => Boolean(options.isEditing),
    applyFormattingToSelection: (_label, fn) => {
      fn(doc, "Sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    },
  };
}

describe("handleRibbonCommand", () => {
  it("returns true for implemented formatting commands", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    expect(handleRibbonCommand(ctx, "format.toggleBold")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.fillColor.yellow")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.numberFormat.number")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.numberFormat.percent")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.numberFormat.longDate")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.fontSize.increase")).toBe(true);
  });

  it("toggles bold", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.toggleBold");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.font?.bold).toBe(true);
  });

  it("applies fill color", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.fillColor.yellow");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.fill?.fgColor).toBe("#FFFFFF00");
  });

  it("delegates fillColor.moreColors to the builtin picker command", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    const executeCommand = vi.fn();
    ctx.executeCommand = executeCommand;

    expect(handleRibbonCommand(ctx, "format.fillColor.moreColors")).toBe(true);
    expect(executeCommand).toHaveBeenCalledWith("format.fillColor");
  });

  it("delegates fontColor.moreColors to the builtin picker command", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    const executeCommand = vi.fn();
    ctx.executeCommand = executeCommand;

    expect(handleRibbonCommand(ctx, "format.fontColor.moreColors")).toBe(true);
    expect(executeCommand).toHaveBeenCalledWith("format.fontColor");
  });

  it("applies number format", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.number");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("0.00");
  });

  it("applies long date format", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.longDate");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("yyyy-mm-dd");
  });

  it("applies accounting symbol formats", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.accounting.eur");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("€#,##0.00");

    handleRibbonCommand(ctx, "format.numberFormat.accounting.jpy");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("¥#,##0.00");
  });

  it("does not adjust decimals for time-only formats", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.time");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("h:mm:ss");

    handleRibbonCommand(ctx, "format.numberFormat.increaseDecimal");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("h:mm:ss");
  });

  it("preserves scientific notation when stepping decimals", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.scientific");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("0.00E+00");

    handleRibbonCommand(ctx, "format.numberFormat.increaseDecimal");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("0.000E+00");
  });

  it("adjusts classic fraction formats when stepping decimals", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.fraction");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("# ?/?");

    handleRibbonCommand(ctx, "format.numberFormat.increaseDecimal");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("# ??/??");

    handleRibbonCommand(ctx, "format.numberFormat.decreaseDecimal");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("# ?/?");
  });

  it("does not convert text formats when stepping decimals", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.text");
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 }).numberFormat).toBe("@");

    handleRibbonCommand(ctx, "format.numberFormat.increaseDecimal");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("@");
  });

  it("adjusts decimals using snake_case number_format as the source format", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { number_format: "0%" });
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.increaseDecimal");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("0.0%");
  });

  it("increases/decreases font size", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.fontSize.increase");
    let style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.font?.size).toBe(12);

    handleRibbonCommand(ctx, "format.fontSize.decrease");
    style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.font?.size).toBe(11);
  });

  it("returns false for unknown commands", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    expect(handleRibbonCommand(ctx, "home.font.nonexistentCommand")).toBe(false);
  });

  it("merges and unmerges cells", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc, { selection: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }] });

    expect(handleRibbonCommand(ctx, "home.alignment.mergeCenter.mergeCells")).toBe(true);
    expect(doc.getMergedRanges("Sheet1")).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]);

    expect(handleRibbonCommand(ctx, "home.alignment.mergeCenter.unmergeCells")).toBe(true);
    expect(doc.getMergedRanges("Sheet1")).toEqual([]);
  });

  it("routes sort commands through ctx.sortSelection", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    const sortSelection = vi.fn();
    ctx.sortSelection = sortSelection;

    expect(handleRibbonCommand(ctx, "data.sortFilter.sortAtoZ")).toBe(true);
    expect(sortSelection).toHaveBeenCalledWith({ order: "ascending" });

    sortSelection.mockClear();
    expect(handleRibbonCommand(ctx, "data.sortFilter.sortZtoA")).toBe(true);
    expect(sortSelection).toHaveBeenCalledWith({ order: "descending" });
  });

  it("routes custom sort commands through ctx.openCustomSort", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    const openCustomSort = vi.fn();
    ctx.openCustomSort = openCustomSort;

    expect(handleRibbonCommand(ctx, "home.editing.sortFilter.customSort")).toBe(true);
    expect(openCustomSort).toHaveBeenCalledWith("home.editing.sortFilter.customSort");

    openCustomSort.mockClear();
    expect(handleRibbonCommand(ctx, "data.sortFilter.sort.customSort")).toBe(true);
    expect(openCustomSort).toHaveBeenCalledWith("data.sortFilter.sort.customSort");
  });

  it("routes custom number format through ctx.promptCustomNumberFormat", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);
    const promptCustomNumberFormat = vi.fn();
    ctx.promptCustomNumberFormat = promptCustomNumberFormat;

    expect(handleRibbonCommand(ctx, "home.number.moreFormats.custom")).toBe(true);
    expect(promptCustomNumberFormat).toHaveBeenCalledTimes(1);
  });

  it("routes filter commands through ctx auto-filter hooks", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    ctx.toggleAutoFilter = vi.fn();
    ctx.clearAutoFilter = vi.fn();
    ctx.reapplyAutoFilter = vi.fn();

    expect(handleRibbonCommand(ctx, "data.sortFilter.filter")).toBe(true);
    expect(ctx.toggleAutoFilter).toHaveBeenCalledTimes(1);

    expect(handleRibbonCommand(ctx, "data.sortFilter.clear")).toBe(true);
    expect(ctx.clearAutoFilter).toHaveBeenCalledTimes(1);

    expect(handleRibbonCommand(ctx, "data.sortFilter.reapply")).toBe(true);
    expect(ctx.reapplyAutoFilter).toHaveBeenCalledTimes(1);
  });

  it("no-ops sort/filter and picker commands while editing", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc, { isEditing: true });
    ctx.sortSelection = vi.fn();
    ctx.openCustomSort = vi.fn();
    ctx.toggleAutoFilter = vi.fn();
    ctx.clearAutoFilter = vi.fn();
    ctx.reapplyAutoFilter = vi.fn();
    ctx.promptCustomNumberFormat = vi.fn();
    ctx.executeCommand = vi.fn();
    ctx.openFormatCells = vi.fn();

    expect(handleRibbonCommand(ctx, "data.sortFilter.sortAtoZ")).toBe(true);
    expect(handleRibbonCommand(ctx, "data.sortFilter.sortZtoA")).toBe(true);
    expect(handleRibbonCommand(ctx, "data.sortFilter.sort.customSort")).toBe(true);
    expect(handleRibbonCommand(ctx, "data.sortFilter.filter")).toBe(true);
    expect(handleRibbonCommand(ctx, "data.sortFilter.clear")).toBe(true);
    expect(handleRibbonCommand(ctx, "data.sortFilter.reapply")).toBe(true);

    expect(handleRibbonCommand(ctx, "home.number.moreFormats.custom")).toBe(true);
    expect(handleRibbonCommand(ctx, "format.openFormatCells")).toBe(true);

    expect(ctx.sortSelection).not.toHaveBeenCalled();
    expect(ctx.openCustomSort).not.toHaveBeenCalled();
    expect(ctx.toggleAutoFilter).not.toHaveBeenCalled();
    expect(ctx.clearAutoFilter).not.toHaveBeenCalled();
    expect(ctx.reapplyAutoFilter).not.toHaveBeenCalled();
    expect(ctx.promptCustomNumberFormat).not.toHaveBeenCalled();
    expect(ctx.executeCommand).not.toHaveBeenCalled();
    expect(ctx.openFormatCells).not.toHaveBeenCalled();
  });

  it("treats dropdown trigger ids as no-op fallbacks and delegates where appropriate", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    const executeCommand = vi.fn();
    ctx.executeCommand = executeCommand;

    // Dropdown triggers with menu items should not execute formatting directly when invoked.
    expect(handleRibbonCommand(ctx, "home.font.fontName")).toBe(true);
    expect(doc.getCellFormat("Sheet1", { row: 0, col: 0 })).toEqual({});
    expect(executeCommand).not.toHaveBeenCalled();

    expect(handleRibbonCommand(ctx, "home.font.clearFormatting")).toBe(true);
    expect(executeCommand).not.toHaveBeenCalled();

    expect(handleRibbonCommand(ctx, "home.alignment.orientation")).toBe(true);
    expect(executeCommand).not.toHaveBeenCalled();

    expect(handleRibbonCommand(ctx, "home.number.numberFormat")).toBe(true);
    expect(handleRibbonCommand(ctx, "home.number.moreFormats")).toBe(true);
    expect(executeCommand).not.toHaveBeenCalled();

    // Some legacy trigger ids route to canonical commands for opening pickers.
    expect(handleRibbonCommand(ctx, "home.font.fillColor")).toBe(true);
    expect(executeCommand).toHaveBeenLastCalledWith("format.fillColor");

    expect(handleRibbonCommand(ctx, "home.font.fontColor")).toBe(true);
    expect(executeCommand).toHaveBeenLastCalledWith("format.fontColor");

    expect(handleRibbonCommand(ctx, "home.font.fontSize")).toBe(true);
    expect(executeCommand).toHaveBeenLastCalledWith("format.fontSize.set");
  });
});
