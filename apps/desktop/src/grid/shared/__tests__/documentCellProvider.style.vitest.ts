import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { applyAllBorders, setFillColor, setHorizontalAlign, toggleBold, toggleItalic, toggleUnderline, toggleWrap } from "../../../formatting/toolbar.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider (shared grid) style mapping", () => {
  it("maps DocumentController style table entries into @formula/grid CellStyle", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Styled");

    toggleBold(doc, sheetId, "A1");
    toggleItalic(doc, sheetId, "A1");
    toggleUnderline(doc, sheetId, "A1");
    setFillColor(doc, sheetId, "A1", "#FFFFFF00");
    applyAllBorders(doc, sheetId, "A1", { style: "thin", color: "#FF000000" });
    setHorizontalAlign(doc, sheetId, "A1", "center");
    toggleWrap(doc, sheetId, "A1");

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect(cell?.value).toBe("Styled");

    const style = (cell as any).style as any;
    expect(style).toBeTruthy();

    expect(style.fill).toBe("#ffff00");
    expect(style.fontWeight).toBe("700");
    expect(style.fontStyle).toBe("italic");
    expect(style.underline).toBe(true);
    expect(style.textAlign).toBe("center");
    expect(style.wrapMode).toBe("word");

    expect(style.borders).toEqual({
      left: { width: 1, style: "solid", color: "#000000" },
      right: { width: 1, style: "solid", color: "#000000" },
      top: { width: 1, style: "solid", color: "#000000" },
      bottom: { width: 1, style: "solid", color: "#000000" }
    });
  });

  it("supports formula-model snake_case style patches (font.size_100pt, fill.fg_color, alignment.wrap_text)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Styled");
    doc.setRangeFormat(sheetId, "A1", {
      font: { size_100pt: 1200 },
      fill: { fg_color: "#FFFF0000" },
      alignment: { wrap_text: true }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();

    const style = (cell as any).style as any;
    expect(style).toBeTruthy();
    // 1200 => 12pt => 16px @96DPI.
    expect(style.fontSize).toBeCloseTo(16, 5);
    expect(style.fill).toBe("#ff0000");
    expect(style.wrapMode).toBe("word");
  });

  it("converts semi-transparent Excel ARGB fill colors into rgba() for canvas rendering", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Alpha");
    // #AARRGGBB with alpha < 0xFF should be converted to rgba(..., a).
    setFillColor(doc, sheetId, "A1", "#80FF0000");

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any).style?.fill).toBe("rgba(255,0,0,0.502)");
  });

  it("supports theme/indexed color reference objects when mapping fill + font colors", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Color refs");
    doc.setRangeFormat(sheetId, "A1", {
      fill: { fgColor: { theme: 4 } },
      font: { color: { indexed: 2 } }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any).style?.fill).toBe("#5b9bd5"); // Office 2013 theme accent1
    expect((cell as any).style?.color).toBe("#ff0000"); // Excel indexed color 2
  });

  it('supports snake_case alignment.vertical_alignment when mapping vertical alignment', () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Aligned");
    doc.setRangeFormat(sheetId, "A1", { alignment: { vertical_alignment: "center" } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any).style?.verticalAlign).toBe("middle");
  });

  it("allows UI patches to clear an imported snake_case alignment.vertical_alignment", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Aligned");
    // Imported formula-model style (snake_case).
    doc.setRangeFormat(sheetId, "A1", { alignment: { vertical_alignment: "top" }, font: { bold: true } });
    // User clears back to default/general.
    doc.setRangeFormat(sheetId, "A1", { alignment: { vertical: null } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any).style?.fontWeight).toBe("700");
    expect((cell as any).style?.verticalAlign).toBeUndefined();
  });

  it("allows modern fill:null patches to clear legacy flat backgroundColor fills", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "No fill");
    // Legacy/clipboard-ish payloads store background colors as flat keys.
    doc.setRangeFormat(sheetId, "A1", { backgroundColor: "#FFFFFF00", font: { bold: true } });
    // Modern UI "No fill" writes `fill: null`.
    doc.setRangeFormat(sheetId, "A1", { fill: null });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    // Bold is preserved...
    expect((cell as any).style?.fontWeight).toBe("700");
    // ...but fill should be cleared.
    expect((cell as any).style?.fill).toBeUndefined();
  });

  it("allows font.color:null patches to clear legacy flat font_color (automatic font color)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Auto font");
    // Legacy/clipboard-ish payloads store font colors as flat keys.
    doc.setRangeFormat(sheetId, "A1", { font_color: "#FF00FF00", bold: true });
    // The ribbon "automatic" font color option writes `font.color: null`.
    doc.setRangeFormat(sheetId, "A1", { font: { color: null } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any).style?.fontWeight).toBe("700");
    expect((cell as any).style?.color).toBeUndefined();
  });

  it("prefers camelCase overrides when both snake_case + camelCase fields exist (imported style then user edits)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Styled");
    // Imported formula-model style (snake_case)
    doc.setRangeFormat(sheetId, "A1", {
      font: { size_100pt: 1200 },
      alignment: { wrap_text: true }
    });
    // User edits via toolbar / dialogs (camelCase)
    doc.setRangeFormat(sheetId, "A1", {
      font: { size: 20 },
      alignment: { wrapText: false }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();

    const style = (cell as any).style as any;
    expect(style).toBeTruthy();
    // 20pt -> 26.666...px @96DPI.
    expect(style.fontSize).toBeCloseTo((20 * 96) / 72, 5);
    // wrapText:false should override imported wrap_text:true.
    expect(style.wrapMode).toBeUndefined();
  });

  it("allows UI patches to clear an imported snake_case alignment.horizontal_alignment", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Aligned");
    // Imported formula-model style (snake_case).
    doc.setRangeFormat(sheetId, "A1", { alignment: { horizontal_alignment: "right" }, font: { bold: true } });
    // User clears alignment back to "General" (null in UI patch representation).
    doc.setRangeFormat(sheetId, "A1", { alignment: { horizontal: null } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    // Bold is preserved...
    expect((cell as any).style?.fontWeight).toBe("700");
    // ...but the explicit null alignment should clear the imported horizontal_alignment.
    expect((cell as any).style?.textAlign).toBeUndefined();
    expect((cell as any).style?.horizontalAlign).toBeUndefined();
  });

  it("maps a variety of Excel border styles into grid border primitives", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Borders");
    doc.setRangeFormat(sheetId, "A1", {
      border: {
        left: { style: "medium", color: "#FF112233" },
        right: { style: "thick", color: "#FF112233" },
        top: { style: "dashed", color: "#FF112233" },
        bottom: { style: "double", color: "#FF112233" },
        diagonal: { style: "dotted", color: "#FF112233" },
        diagonalUp: true
      }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();

    const style = (cell as any).style as any;
    expect(style?.borders).toEqual({
      left: { width: 2, style: "solid", color: "#112233" },
      right: { width: 3, style: "solid", color: "#112233" },
      top: { width: 1, style: "dashed", color: "#112233" },
      bottom: { width: 3, style: "double", color: "#112233" }
    });
    expect(style?.diagonalBorders).toEqual({
      up: { width: 1, style: "dotted", color: "rgba(17,34,51,1)" }
    });
  });

  it("falls back to a theme-safe border color when no explicit border color is provided", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Borders");
    doc.setRangeFormat(sheetId, "A1", {
      border: {
        left: { style: "thin" },
        diagonal: { style: "thin" },
        diagonalDown: true
      }
    });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();

    const style = (cell as any).style as any;
    // In unit tests (no DOM), resolveCssVar falls back to the provided fallback "CanvasText".
    expect(style?.borders?.left?.color).toBe("CanvasText");
    expect(style?.diagonalBorders?.down?.color).toBe("CanvasText");
  });

  it("allows UI patches to clear an imported snake_case number_format", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", 0.5);
    // Imported formula-model style (snake_case).
    doc.setRangeFormat(sheetId, "A1", { number_format: "0%" });

    const importedProvider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const importedCell = importedProvider.getCell(1, 1);
    expect(importedCell).not.toBeNull();
    expect(importedCell?.value).toBe("50%");

    // User clears the number format (e.g. Format Cells → Number → General).
    doc.setRangeFormat(sheetId, "A1", { numberFormat: null });

    const clearedProvider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const clearedCell = clearedProvider.getCell(1, 1);
    expect(clearedCell).not.toBeNull();
    // If `numberFormat: null` is present, it should override the imported `number_format` string.
    expect(clearedCell?.value).toBe(0.5);
  });

  it("treats an imported 'General' number_format as equivalent to no number format", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", 0.1234567);
    // Imported formula-model style (snake_case) may explicitly encode "General".
    doc.setRangeFormat(sheetId, "A1", { number_format: "General" });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    // General should not force numeric values through the custom formatter; keep as number.
    expect(cell?.value).toBe(0.1234567);
  });

  it("maps alignment.indent into grid CellStyle.textIndentPx", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Indented");
    doc.setRangeFormat(sheetId, "A1", { alignment: { indent: 2 } });

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const cell = provider.getCell(1, 1);
    expect(cell).not.toBeNull();
    expect((cell as any)?.style?.textIndentPx).toBe(16);

    // Absent when indent is 0/undefined.
    doc.setCellValue(sheetId, "B1", "No indent");
    doc.setRangeFormat(sheetId, "B1", { font: { bold: true }, alignment: { indent: 0 } });
    const cellB1 = provider.getCell(1, 2);
    expect(cellB1).not.toBeNull();
    expect((cellB1 as any)?.style?.fontWeight).toBe("700");
    expect((cellB1 as any)?.style?.textIndentPx).toBeUndefined();

    doc.setCellValue(sheetId, "C1", "No indent 2");
    const cellC1 = provider.getCell(1, 3);
    expect(cellC1).not.toBeNull();
    expect((cellC1 as any)?.style?.textIndentPx).toBeUndefined();
  });

  it("respects layered sheet/row/col/cell formatting precedence", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "Styled");
    // Sheet -> col -> row -> cell precedence.
    doc.setSheetFormat(sheetId, { fill: { pattern: "solid", fgColor: "#FFFFFF00" } }); // yellow
    doc.setColFormat(sheetId, 0, { fill: { pattern: "solid", fgColor: "#FF00FF00" } }); // green
    doc.setRowFormat(sheetId, 0, { fill: { pattern: "solid", fgColor: "#FF0000FF" } }); // blue
    doc.setRangeFormat(sheetId, "A1", { fill: { pattern: "solid", fgColor: "#FFFF0000" } }); // red (cell)

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    const a1 = provider.getCell(1, 1);
    expect(a1).not.toBeNull();
    expect(a1?.value).toBe("Styled");
    expect((a1 as any).style?.fill).toBe("#ff0000");

    // B1: row formatting wins (no col/cell override).
    const b1 = provider.getCell(1, 2);
    expect(b1).not.toBeNull();
    expect((b1 as any).style?.fill).toBe("#0000ff");

    // A2: col formatting wins (no row/cell override).
    const a2 = provider.getCell(2, 1);
    expect(a2).not.toBeNull();
    expect((a2 as any).style?.fill).toBe("#00ff00");
  });

  it("respects range-run formatting precedence (sheet < col < row < range-run < cell)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // Establish base layers.
    doc.setSheetFormat(sheetId, { fill: { pattern: "solid", fgColor: "#FFFFFF00" } }); // yellow
    doc.setColFormat(sheetId, 0, { fill: { pattern: "solid", fgColor: "#FF00FF00" } }); // green (col A)
    doc.setRowFormat(sheetId, 0, { fill: { pattern: "solid", fgColor: "#FF0000FF" } }); // blue (row 1)

    // Apply a large rectangular range so DocumentController stores it as range runs (not per-cell styles).
    // A1:D20000 => 20,000 rows * 4 cols = 80,000 cells (> 50,000 threshold).
    doc.setRangeFormat(sheetId, "A1:D20000", { fill: { pattern: "solid", fgColor: "#FF800080" } }); // purple (range-run)

    // Assert the range-run layer is actually in use (not enumerated per-cell styles).
    // DocumentController.getCellFormatStyleIds returns:
    // [sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId, rangeRunStyleId].
    const b1Ids = doc.getCellFormatStyleIds(sheetId, "B1");
    expect(b1Ids[0]).not.toBe(0); // sheet
    expect(b1Ids[1]).not.toBe(0); // row 1 (0-based row 0)
    expect(b1Ids[2]).toBe(0); // no col formatting for col B
    expect(b1Ids[3]).toBe(0); // no explicit cell formatting
    expect(b1Ids[4]).not.toBe(0); // range-run formatting

    const a2Ids = doc.getCellFormatStyleIds(sheetId, "A2");
    expect(a2Ids[0]).not.toBe(0); // sheet
    expect(a2Ids[1]).toBe(0); // no row formatting for row 2 (0-based row 1)
    expect(a2Ids[2]).not.toBe(0); // col A formatting
    expect(a2Ids[3]).toBe(0); // no explicit cell formatting
    expect(a2Ids[4]).not.toBe(0); // range-run formatting

    // Cell formatting should still win over the range-run layer.
    doc.setRangeFormat(sheetId, "A1", { fill: { pattern: "solid", fgColor: "#FFFF0000" } }); // red (cell)

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => sheetId,
      headerRows: 1,
      headerCols: 1,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    // A1: cell wins.
    expect((provider.getCell(1, 1) as any)?.style?.fill).toBe("#ff0000");
    // B1: range-run wins over row.
    expect((provider.getCell(1, 2) as any)?.style?.fill).toBe("#800080");
    // A2: range-run wins over col.
    expect((provider.getCell(2, 1) as any)?.style?.fill).toBe("#800080");
    // E1: no range-run, so row wins over sheet.
    expect((provider.getCell(1, 5) as any)?.style?.fill).toBe("#0000ff");
  });
});
