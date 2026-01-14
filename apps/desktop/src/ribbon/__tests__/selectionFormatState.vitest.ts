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
      strikethrough: false,
      fontName: null,
      fontSize: null,
      fontVariantPosition: null,
      wrapText: false,
      align: "left",
      verticalAlign: "bottom",
      numberFormat: null,
    });
  });

  it("detects bold/italic/underline/wrap on a single cell", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", {
      font: { bold: true, italic: true, underline: true, strike: true },
      alignment: { wrapText: true },
    });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.bold).toBe(true);
    expect(state.italic).toBe(true);
    expect(state.underline).toBe(true);
    expect(state.strikethrough).toBe(true);
    expect(state.wrapText).toBe(true);
  });

  it("detects subscript/superscript via font.vertAlign", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { font: { vertAlign: "superscript" } });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.fontVariantPosition).toBe("superscript");

    doc.setRangeFormat("Sheet1", "B1", { font: { vertAlign: "subscript" } });
    const mixed = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(mixed.fontVariantPosition).toBe("mixed");
  });

  it("detects font name/size and reports mixed when selection differs", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { font: { name: "Arial", size: 14 } });

    const single = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(single.fontName).toBe("Arial");
    expect(single.fontSize).toBe(14);

    const mixed = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(mixed.fontName).toBe("mixed");
    expect(mixed.fontSize).toBe("mixed");
  });

  it("reports mixed state for alignment and number formats", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { alignment: { horizontal: "center" }, numberFormat: "0%" });
    // B1 uses defaults (left alignment + general format).

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(state.align).toBe("mixed");
    expect(state.numberFormat).toBe("mixed");
  });

  it("detects vertical alignment and reports mixed when selection differs", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { alignment: { vertical: "top" } });
    doc.setRangeFormat("Sheet1", "B1", { alignment: { vertical: "bottom" } });

    const single = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(single.verticalAlign).toBe("top");

    const mixed = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(mixed.verticalAlign).toBe("mixed");

    doc.setRangeFormat("Sheet1", "A1", { alignment: { vertical: "center" } });
    doc.setRangeFormat("Sheet1", "B1", { alignment: { vertical: "center" } });
    const centered = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
    expect(centered.verticalAlign).toBe("center");
  });

  it("samples large selections and still reports uniform formatting", () => {
    const doc = new DocumentController();
    // 20x20=400 cells > default maxInspectCells (256), so this exercises sampling mode.
    doc.setRangeFormat("Sheet1", "A1:T20", { font: { bold: true }, alignment: { horizontal: "center" } });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 19, endCol: 19 }]);
    expect(state.bold).toBe(true);
    expect(state.align).toBe("center");
  });

  it("uses effective (layered) formatting for full-column format overrides", () => {
    const doc = new DocumentController();
    // Full column range - this should be stored as a column style layer, leaving most
    // individual cells with `styleId === 0`.
    doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

    const state = computeSelectionFormatState(doc, "Sheet1", [
      { startRow: 0, startCol: 0, endRow: 1_048_575, endCol: 0 },
    ]);
    expect(state.bold).toBe(true);
  });

  it("uses effective formatting for sheet-level number formats", () => {
    const doc = new DocumentController();
    doc.setSheetFormat("Sheet1", { numberFormat: "0.00" });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.numberFormat).toBe("0.00");
  });

  it("uses effective formatting for row-level alignment overrides", () => {
    const doc = new DocumentController();
    doc.setRowFormat("Sheet1", 0, { alignment: { horizontal: "center" } });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.align).toBe("center");
  });

  it("reads snake_case formula-model keys for wrapText, alignment, number_format, and font size_100pt", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", {
      alignment: { wrap_text: true, horizontal_alignment: "center", vertical_alignment: "top" },
      font: { size_100pt: 1200 },
      number_format: "0%",
    });

    const state = computeSelectionFormatState(doc, "Sheet1", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state.wrapText).toBe(true);
    expect(state.align).toBe("center");
    expect(state.verticalAlign).toBe("top");
    expect(state.fontSize).toBe(12);
    expect(state.numberFormat).toBe("0%");
  });

  it("does not resurrect a deleted sheet when called with a stale sheet id", () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellValue("Sheet2", "A1", 2);
    doc.deleteSheet("Sheet2");

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();

    const state = computeSelectionFormatState(doc, "Sheet2", [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    expect(state).toEqual({
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontName: null,
      fontSize: null,
      fontVariantPosition: null,
      wrapText: false,
      align: "left",
      verticalAlign: "bottom",
      numberFormat: null,
    });

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();
  });
});
