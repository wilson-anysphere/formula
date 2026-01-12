import { describe, expect, it } from "vitest";

import { DocumentController } from "../document/documentController.js";

import { applyExcelFillDown, applyExcelFillRight } from "./excelFill";

describe("excel fill shortcuts", () => {
  it("Fill Down copies the top row and shifts formulas", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellInput(sheetId, "A1", "=B1");
    doc.setCellValue(sheetId, "B1", 10);

    applyExcelFillDown({
      document: doc,
      sheetId,
      ranges: [{ startRow: 0, endRow: 2, startCol: 0, endCol: 1 }], // A1:B3
    });

    expect(doc.getCell(sheetId, "B2").value).toBe(10);
    expect(doc.getCell(sheetId, "B3").value).toBe(10);
    expect(doc.getCell(sheetId, "A2").formula).toBe("=B2");
    expect(doc.getCell(sheetId, "A3").formula).toBe("=B3");
    expect(doc.getStackDepths().undo).toBe(1);
    expect(doc.undoLabel).toBe("Fill Down");
  });

  it("Fill Right copies the left column and shifts formulas", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", 1);
    doc.setCellInput(sheetId, "A2", "=A1+1");

    applyExcelFillRight({
      document: doc,
      sheetId,
      ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }], // A1:C2
    });

    expect(doc.getCell(sheetId, "B1").value).toBe(1);
    expect(doc.getCell(sheetId, "C1").value).toBe(1);
    expect(doc.getCell(sheetId, "B2").formula).toBe("=B1+1");
    expect(doc.getCell(sheetId, "C2").formula).toBe("=C1+1");
    expect(doc.getStackDepths().undo).toBe(1);
    expect(doc.undoLabel).toBe("Fill Right");
  });
});

