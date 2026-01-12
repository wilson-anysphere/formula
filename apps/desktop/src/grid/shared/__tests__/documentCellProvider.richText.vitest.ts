import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";
import type { CellRichText } from "@formula/grid/node";

describe("DocumentCellProvider (shared grid) rich text mapping", () => {
  it("preserves DocumentController rich text values on CellData.richText while keeping value as plain text", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const richText: CellRichText = {
      text: "Hello world",
      runs: [
        { start: 0, end: 5, style: { italic: true, underline: "single" } },
        { start: 5, end: 11, style: {} }
      ]
    };

    doc.setCellValue(sheetId, "A1", richText as any);

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
    expect(cell?.value).toBe("Hello world");
    expect((cell as any).richText).toEqual(richText);
  });
});
