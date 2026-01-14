import { describe, expect, it } from "vitest";

import { SpreadsheetModel } from "../../spreadsheet/SpreadsheetModel.js";
import { getLocale, setLocale } from "../../i18n/index.js";

describe("Formula bar tab completion", () => {
  it("suggests VLOOKUP( when typing a function prefix", async () => {
    const sheet = new SpreadsheetModel();

    sheet.selectCell("A1");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=VLO", 4);

    await sheet.flushTabCompletion();

    expect(sheet.formulaBar.aiSuggestion()).toBe("=VLOOKUP(");
    expect(sheet.formulaBar.aiGhostText()).toBe("OKUP(");

    expect(sheet.formulaBar.acceptAiSuggestion()).toBe(true);
    expect(sheet.formulaBar.draft).toBe("=VLOOKUP(");
  });

  it("suggests contiguous ranges for SUM when typing a column reference", async () => {
    const initial: Record<string, number> = {};
    for (let row = 1; row <= 10; row += 1) {
      initial[`A${row}`] = row;
    }

    const sheet = new SpreadsheetModel(initial);
    sheet.selectCell("B11");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=SUM(A", 6);

    await sheet.flushTabCompletion();

    expect(sheet.formulaBar.aiSuggestion()).toBe("=SUM(A1:A10)");
    expect(sheet.formulaBar.aiGhostText()).toBe("1:A10)");

    expect(sheet.formulaBar.acceptAiSuggestion()).toBe(true);
    expect(sheet.formulaBar.draft).toBe("=SUM(A1:A10)");
  });

  it("suggests contiguous ranges for localized function names (de-DE SUMME)", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
    try {
      const initial: Record<string, number> = {};
      for (let row = 1; row <= 10; row += 1) {
        initial[`A${row}`] = row;
      }

      const sheet = new SpreadsheetModel(initial);
      sheet.selectCell("B11");
      sheet.beginFormulaEdit();
      sheet.typeInFormulaBar("=SUMME(A", 8);

      await sheet.flushTabCompletion();

      expect(sheet.formulaBar.aiSuggestion()).toBe("=SUMME(A1:A10)");
      expect(sheet.formulaBar.aiGhostText()).toBe("1:A10)");
    } finally {
      setLocale(prevLocale);
    }
  });

  it("suggests localized starter functions for bare '=' in de-DE", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
    try {
      const sheet = new SpreadsheetModel();

      sheet.selectCell("A1");
      sheet.beginFormulaEdit();
      sheet.typeInFormulaBar("=", 1);

      await sheet.flushTabCompletion();

      expect(sheet.formulaBar.aiSuggestion()).toBe("=SUMME(");
      expect(sheet.formulaBar.aiGhostText()).toBe("SUMME(");
    } finally {
      setLocale(prevLocale);
    }
  });

  it("suggests localized function-name completion in de-DE (SU -> SUMME)", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
    try {
      const sheet = new SpreadsheetModel();

      sheet.selectCell("A1");
      sheet.beginFormulaEdit();
      sheet.typeInFormulaBar("=SU", 3);

      await sheet.flushTabCompletion();

      expect(sheet.formulaBar.aiSuggestion()).toBe("=SUMME(");
      expect(sheet.formulaBar.aiGhostText()).toBe("MME(");

      expect(sheet.formulaBar.acceptAiSuggestion()).toBe(true);
      expect(sheet.formulaBar.draft).toBe("=SUMME(");
    } finally {
      setLocale(prevLocale);
    }
  });

  it("treats blank-valued formulas as non-empty when suggesting ranges", async () => {
    const initial: Record<string, string> = {};
    for (let row = 1; row <= 10; row += 1) {
      // Formula that evaluates to empty string.
      initial[`A${row}`] = '=""';
    }

    const sheet = new SpreadsheetModel(initial);
    sheet.selectCell("B11");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=SUM(A", 6);

    await sheet.flushTabCompletion();

    expect(sheet.formulaBar.aiSuggestion()).toBe("=SUM(A1:A10)");
    expect(sheet.formulaBar.aiGhostText()).toBe("1:A10)");
  });

  it("suggests TODAY() for zero-arg functions", async () => {
    const sheet = new SpreadsheetModel();

    sheet.selectCell("A1");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=TOD", 4);

    await sheet.flushTabCompletion();

    expect(sheet.formulaBar.aiSuggestion()).toBe("=TODAY()");
    expect(sheet.formulaBar.acceptAiSuggestion()).toBe(true);
    expect(sheet.formulaBar.draft).toBe("=TODAY()");
  });

  it("suggests _xlfn.XLOOKUP( when typing an _xlfn. function prefix", async () => {
    const sheet = new SpreadsheetModel();

    sheet.selectCell("A1");
    sheet.beginFormulaEdit();
    sheet.typeInFormulaBar("=_xlfn.XLO", 10);

    await sheet.flushTabCompletion();

    expect(sheet.formulaBar.aiSuggestion()).toBe("=_xlfn.XLOOKUP(");
    expect(sheet.formulaBar.acceptAiSuggestion()).toBe(true);
    expect(sheet.formulaBar.draft).toBe("=_xlfn.XLOOKUP(");
  });
});
