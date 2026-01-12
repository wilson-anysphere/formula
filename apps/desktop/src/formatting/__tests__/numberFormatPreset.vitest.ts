import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { NUMBER_FORMATS, applyNumberFormatPreset } from "../toolbar.js";

describe("applyNumberFormatPreset", () => {
  it("applies the preset to every cell in the target range", () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", "A1:B2", [
      [1, 2],
      [3, 4],
    ]);

    applyNumberFormatPreset(doc, "Sheet1", { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }, "currency");

    for (const addr of ["A1", "B1", "A2", "B2"]) {
      const cell = doc.getCell("Sheet1", addr);
      const style = doc.styleTable.get(cell.styleId);
      expect(style.numberFormat).toBe(NUMBER_FORMATS.currency);
    }
  });
});

