import { describe, expect, it } from "vitest";

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

  it("applies number format", () => {
    const doc = new DocumentController();
    const ctx = createCtx(doc);

    handleRibbonCommand(ctx, "format.numberFormat.number");

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
    expect(style.numberFormat).toBe("0.00");
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
});
