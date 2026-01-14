import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../extensions/ui.js", () => ({
  showInputBox: vi.fn(),
  showToast: vi.fn(),
}));

import { DocumentController } from "../../document/documentController.js";
import { promptAndApplyAxisSizing } from "../axisSizing.js";
import { showInputBox, showToast } from "../../extensions/ui.js";

describe("ribbon/axisSizing", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("enumerates selected rows and applies row heights via DocumentController sheet view deltas", async () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const setRowHeightSpy = vi.spyOn(doc, "setRowHeight");
    const beginBatchSpy = vi.spyOn(doc, "beginBatch");
    const endBatchSpy = vi.spyOn(doc, "endBatch");

    vi.mocked(showInputBox).mockResolvedValue("42");

    const focus = vi.fn();
    const app = {
      getSelectionRanges: () => [
        { startRow: 1, endRow: 3, startCol: 0, endCol: 0 },
        { startRow: 2, endRow: 4, startCol: 5, endCol: 5 },
      ],
      getCurrentSheetId: () => sheetId,
      getDocument: () => doc,
      focus,
      isEditing: () => false,
    };

    await promptAndApplyAxisSizing(app, "rowHeight");

    expect(beginBatchSpy).toHaveBeenCalledTimes(1);
    expect(endBatchSpy).toHaveBeenCalledTimes(1);

    expect(setRowHeightSpy).toHaveBeenCalledTimes(4);
    expect(setRowHeightSpy.mock.calls.map((call) => call[1])).toEqual([1, 2, 3, 4]);

    const view = doc.getSheetView(sheetId);
    expect(view.rowHeights).toEqual({ "1": 42, "2": 42, "3": 42, "4": 42 });
    expect(view.colWidths).toBeUndefined();

    expect(focus).toHaveBeenCalledTimes(1);
  });

  it("enumerates selected columns and applies column widths via DocumentController sheet view deltas", async () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const setColWidthSpy = vi.spyOn(doc, "setColWidth");
    vi.mocked(showInputBox).mockResolvedValue("120");

    const focus = vi.fn();
    const app = {
      getSelectionRanges: () => [
        { startRow: 0, endRow: 0, startCol: 0, endCol: 2 },
        { startRow: 5, endRow: 5, startCol: 2, endCol: 4 },
      ],
      getCurrentSheetId: () => sheetId,
      getDocument: () => doc,
      focus,
      isEditing: () => false,
    };

    await promptAndApplyAxisSizing(app, "colWidth");

    expect(setColWidthSpy).toHaveBeenCalledTimes(5);
    expect(setColWidthSpy.mock.calls.map((call) => call[1])).toEqual([0, 1, 2, 3, 4]);

    const view = doc.getSheetView(sheetId);
    expect(view.colWidths).toEqual({ "0": 120, "1": 120, "2": 120, "3": 120, "4": 120 });
    expect(view.rowHeights).toBeUndefined();

    expect(focus).toHaveBeenCalledTimes(1);
  });

  it("aborts before prompting when the selection implies >10k rows/cols", async () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const setRowHeightSpy = vi.spyOn(doc, "setRowHeight");

    const focus = vi.fn();
    const app = {
      getSelectionRanges: () => [{ startRow: 0, endRow: 10_000, startCol: 0, endCol: 0 }],
      getCurrentSheetId: () => sheetId,
      getDocument: () => doc,
      focus,
      isEditing: () => false,
    };

    await promptAndApplyAxisSizing(app, "rowHeight");

    expect(showInputBox).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledTimes(1);
    expect(vi.mocked(showToast).mock.calls[0]?.[0]).toContain("Selection too large");
    expect(setRowHeightSpy).not.toHaveBeenCalled();
    expect(focus).not.toHaveBeenCalled();
  });

  it("no-ops in read-only mode before prompting", async () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const setRowHeightSpy = vi.spyOn(doc, "setRowHeight");
    vi.mocked(showInputBox).mockResolvedValue("42");

    const focus = vi.fn();
    const app = {
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      getCurrentSheetId: () => sheetId,
      getDocument: () => doc,
      focus,
      isEditing: () => false,
      isReadOnly: () => true,
    };

    await promptAndApplyAxisSizing(app, "rowHeight");

    expect(showInputBox).not.toHaveBeenCalled();
    expect(setRowHeightSpy).not.toHaveBeenCalled();
    expect(focus).not.toHaveBeenCalled();
  });
});
