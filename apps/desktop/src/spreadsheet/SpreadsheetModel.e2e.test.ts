import { describe, expect, it } from "vitest";

import { SpreadsheetModel } from "./SpreadsheetModel.js";
import { getLocale, setLocale } from "../i18n/index.js";

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

  it("uses document.documentElement.lang for locale-aware evaluation when i18n locale is not wired", () => {
    const beforeLocale = getLocale();
    const beforeDocument = (globalThis as any).document;
    try {
      (globalThis as any).document = { documentElement: { lang: "de_DE.UTF-8" } };
      // Simulate "i18n locale not wired" (default en-US) while the host still sets `<html lang>`.
      setLocale("en-US");
      (globalThis as any).document.documentElement.lang = "de_DE.UTF-8";

      const sheet = new SpreadsheetModel();
      sheet.setCellInput("A1", "=SUMME(1;2)");
      expect(sheet.getCellValue("A1")).toBe(3);
    } finally {
      setLocale(beforeLocale);
      if (beforeDocument === undefined) {
        delete (globalThis as any).document;
      } else {
        (globalThis as any).document = beforeDocument;
      }
    }
  });
});
