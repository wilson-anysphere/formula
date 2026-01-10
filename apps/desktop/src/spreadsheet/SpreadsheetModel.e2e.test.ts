import { describe, expect, it } from "vitest";

import { SpreadsheetModel } from "./SpreadsheetModel.js";

describe("SpreadsheetModel E2E", () => {
  it("type formula, click+drag range, commit formula, computed value updates", () => {
    const sheet = new SpreadsheetModel({ A1: 1, A2: 2 });

    sheet.selectCell("C1");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=SUM(", "=SUM(".length);

    sheet.beginRangeSelection("A1");
    sheet.updateRangeSelection("A2");
    sheet.endRangeSelection();

    expect(sheet.formulaBar.draft).toBe("=SUM(A1:A2");
    expect(sheet.formulaBar.hoveredReference()).toEqual(sheet.selection);
    sheet.typeInFormulaBar(sheet.formulaBar.draft + ")", sheet.formulaBar.draft.length + 1);
    sheet.commitFormulaBar();

    expect(sheet.getCellValue("C1")).toBe(3);
  });
});
