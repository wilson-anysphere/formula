import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider (shared grid) hyperlink styling", () => {
  it("applies a default underline + link color token for URL-like string values", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "https://example.com");

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
    expect(cell?.value).toBe("https://example.com");

    const style = (cell as any)?.style as any;
    expect(style?.underline).toBe(true);
    // In unit tests (non-DOM), `resolveCssVar()` falls back to a system color keyword.
    expect(style?.color).toBe("LinkText");
  });

  it("does not override an explicit font color while still adding default underline", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    doc.setCellValue(sheetId, "A1", "https://example.com");
    doc.setRangeFormat(sheetId, "A1", { font: { color: "#FFFF0000" } });

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
    const style = (cell as any)?.style as any;
    expect(style?.underline).toBe(true);
    expect(style?.color).toBe("#ff0000");
  });
});

