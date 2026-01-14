import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { getEffectiveCellStyle } from "../getEffectiveCellStyle.js";

describe("getEffectiveCellStyle", () => {
  it("does not resurrect a deleted sheet when probing formatting", () => {
    const doc = new DocumentController();

    // Create two sheets, then delete one.
    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellValue("Sheet2", "A1", 2);
    doc.deleteSheet("Sheet2");

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();

    // A stale sheet id should be treated as missing. Importantly, this must not
    // materialize a new (phantom) sheet via DocumentController.getCellFormat/getCell.
    expect(getEffectiveCellStyle(doc, "Sheet2", "A1")).toEqual({});

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();
  });
});

