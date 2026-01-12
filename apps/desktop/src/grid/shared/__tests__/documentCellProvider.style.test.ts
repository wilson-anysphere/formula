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

    expect(style.fill).toBe("rgba(255,255,0,1)");
    expect(style.fontWeight).toBe("700");
    expect(style.fontStyle).toBe("italic");
    expect(style.underline).toBe(true);
    expect(style.textAlign).toBe("center");
    expect(style.wrapMode).toBe("word");

    expect(style.borders).toEqual({
      left: { width: 1, style: "solid", color: "rgba(0,0,0,1)" },
      right: { width: 1, style: "solid", color: "rgba(0,0,0,1)" },
      top: { width: 1, style: "solid", color: "rgba(0,0,0,1)" },
      bottom: { width: 1, style: "solid", color: "rgba(0,0,0,1)" }
    });
  });
});

