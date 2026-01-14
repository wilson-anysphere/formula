import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider no-resurrection", () => {
  it("does not resurrect deleted sheets when rendering cells or merged ranges", () => {
    const doc = new DocumentController();

    doc.setCellValue("Sheet1", "A1", "one");
    doc.setCellValue("Sheet2", "A1", "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet2",
      headerRows: 0,
      headerCols: 0,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null
    });

    expect(provider.getCell(0, 0)).toBeNull();
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    expect(provider.getMergedRangeAt(0, 0)).toBeNull();
    expect(provider.getMergedRangesInRange({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 })).toEqual([]);
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
  });
});

