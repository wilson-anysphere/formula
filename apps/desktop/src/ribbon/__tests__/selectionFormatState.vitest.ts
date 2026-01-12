import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { computeSelectionFormatState } from "../selectionFormatState.js";

describe("computeSelectionFormatState", () => {
  it("returns default formatting for an unformatted single-cell selection", () => {
    const doc = new DocumentController();
    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state).toEqual({
      bold: false,
      italic: false,
      underline: false,
      wrapText: false,
      align: "left",
      numberFormat: null,
    });
  });

  it("detects bold/italic/underline/wrap on a single cell", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", {
      font: { bold: true, italic: true, underline: true },
      alignment: { wrapText: true },
    });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.bold).toBe(true);
    expect(state.italic).toBe(true);
    expect(state.underline).toBe(true);
    expect(state.wrapText).toBe(true);
  });

  it("reports mixed state for alignment and number formats", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { alignment: { horizontal: "center" }, numberFormat: "0%" });
    // B1 uses defaults (left alignment + general format).

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(state.align).toBe("mixed");
    expect(state.numberFormat).toBe("mixed");
  });

  it("samples large selections and still reports uniform formatting", () => {
    const doc = new DocumentController();
    // 20x20=400 cells > default maxInspectCells (256), so this exercises sampling mode.
    doc.setRangeFormat("Sheet1", "A1:T20", { font: { bold: true }, alignment: { horizontal: "center" } });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 19, endCol: 19 }]);
    expect(state.bold).toBe(true);
    expect(state.align).toBe("center");
  });
});

