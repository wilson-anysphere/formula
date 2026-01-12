import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { applyAllBorders, setFillColor, setHorizontalAlign, toggleBold, toggleItalic, toggleUnderline, toggleWrap } from "../../../formatting/toolbar.js";
import { DocumentCellProvider } from "../documentCellProvider";

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
});
