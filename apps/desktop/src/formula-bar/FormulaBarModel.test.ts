import { describe, expect, it } from "vitest";

import { FormulaBarModel } from "./FormulaBarModel.js";
import { parseA1Range } from "../spreadsheet/a1.js";
import { FORMULA_REFERENCE_PALETTE } from "@formula/spreadsheet-frontend";

describe("FormulaBarModel", () => {
  it("inserts and updates range selections while editing", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "C1", input: "=SUM(", value: null });
    model.beginEdit();

    model.updateDraft("=SUM(", 5, 5);
    model.beginRangeSelection(parseA1Range("A1:A2")!);
    expect(model.draft).toBe("=SUM(A1:A2");
    expect(model.hoveredReference()).toEqual(parseA1Range("A1:A2"));

    model.updateRangeSelection(parseA1Range("A1:A3")!);
    expect(model.draft).toBe("=SUM(A1:A3");
    expect(model.hoveredReference()).toEqual(parseA1Range("A1:A3"));

    model.endRangeSelection();
    model.updateDraft(model.draft + ")", model.draft.length + 1, model.draft.length + 1);
    expect(model.draft).toBe("=SUM(A1:A3)");
  });

  it("replaces the active reference when selecting a new range", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "C1", input: "=A1+B1", value: null });
    model.beginEdit();

    // Place the caret within the first reference (A1).
    model.updateDraft("=A1+B1", 2, 2);
    model.beginRangeSelection(parseA1Range("D1")!);
    expect(model.draft).toBe("=D1+B1");
  });

  it("formats sheet-qualified references when selecting a range on another sheet", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "C1", input: "=", value: null });
    model.beginEdit();
    model.updateDraft("=", 1, 1);

    model.beginRangeSelection(parseA1Range("B2")!, "Sheet 2");
    expect(model.draft).toBe("='Sheet 2'!B2");
    expect(model.hoveredReference()).toEqual(parseA1Range("B2"));

    model.updateRangeSelection(parseA1Range("B2:C3")!, "O'Hare");
    expect(model.draft).toBe("='O''Hare'!B2:C3");
    expect(model.hoveredReference()).toEqual(parseA1Range("B2:C3"));

    model.updateRangeSelection(parseA1Range("B2:C3")!, "TRUE");
    expect(model.draft).toBe("='TRUE'!B2:C3");

    model.updateRangeSelection(parseA1Range("B2:C3")!, "A1");
    expect(model.draft).toBe("='A1'!B2:C3");
  });

  it("accepts AI suggestions as an insertion at the caret", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "A1", input: "=SU", value: null });
    model.beginEdit();
    model.updateDraft("=SU", 3, 3);

    model.setAiSuggestion("=SUM");
    expect(model.acceptAiSuggestion()).toBe(true);
    expect(model.draft).toBe("=SUM");
    expect(model.cursorStart).toBe(4);
    expect(model.cursorEnd).toBe(4);
  });

  it("parses sheet-qualified references for hover previews", () => {
    const model = new FormulaBarModel();
    expect(model.resolveReferenceText("A1:B2")).toEqual(parseA1Range("A1:B2"));
    expect(model.resolveReferenceText("Sheet2!A1:B2")).toEqual(parseA1Range("A1:B2"));
    expect(model.resolveReferenceText("'My Sheet'!A1")).toEqual(parseA1Range("A1"));
  });

  it("treats sheet-qualified ranges as their A1 portion when hovering by cursor", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "C1", input: "=Sheet2!A1:B2", value: null });
    model.beginEdit();
    expect(model.hoveredReference()).toEqual(parseA1Range("A1:B2"));
  });

  it("setHoveredReference parses sheet-qualified ranges", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "A1", input: "", value: null });
    model.setHoveredReference("Sheet2!A1:B2");
    expect(model.hoveredReference()).toEqual(parseA1Range("A1:B2"));
  });

  it("resolves structured table references for hover previews when tables are configured", () => {
    const model = new FormulaBarModel();
    model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 0,
          sheetName: "Sheet1",
        },
      ],
    });

    const formula = "=SUM(Table1[Amount])";
    model.setActiveCell({ address: "A1", input: formula, value: null });
    model.beginEdit();

    // Place caret inside the structured ref token.
    const caret = formula.indexOf("Amount") + 1;
    model.updateDraft(formula, caret, caret);

    expect(model.hoveredReferenceText()).toBe("Table1[Amount]");
    expect(model.hoveredReference()).toEqual(parseA1Range("A2:A3"));

    // `setHoveredReference()` also uses the structured-ref resolver.
    model.setHoveredReference("Table1[Amount]");
    expect(model.hoveredReference()).toEqual(parseA1Range("A2:A3"));
  });

  it("resolves nested structured refs (e.g. [[#All],[Column]]) when tables are configured", () => {
    const model = new FormulaBarModel();
    model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 0,
          sheetName: "Sheet1",
        },
      ],
    });

    model.setActiveCell({ address: "A1", input: "", value: null });
    model.setHoveredReference("Table1[[#All],[Amount]]");
    // #All includes the header row; with a 3-row table this is A1:A3.
    expect(model.hoveredReference()).toEqual(parseA1Range("A1:A3"));
  });

  it("includes named ranges in reference highlights when a resolver is provided", () => {
    const model = new FormulaBarModel();
    model.setNameResolver((name) =>
      name === "SalesData" ? { startRow: 0, startCol: 0, endRow: 9, endCol: 0, sheet: "Sheet1" } : null
    );
    model.setActiveCell({ address: "A1", input: "=SUM(SalesData)", value: null });
    model.beginEdit();

    const caretInside = "=SUM(".length + 1;
    model.updateDraft("=SUM(SalesData)", caretInside, caretInside);

    expect(model.referenceHighlights()).toEqual([
      {
        range: { startRow: 0, startCol: 0, endRow: 9, endCol: 0, sheet: "Sheet1" },
        color: FORMULA_REFERENCE_PALETTE[0],
        text: "SalesData",
        index: 0,
        active: true,
      },
    ]);
  });

  it("returns explanations for AI-related error codes", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "A1", input: "", value: "#GETTING_DATA" });
    expect(model.errorExplanation()?.title).toBe("Loading");

    model.setActiveCell({ address: "A1", input: "", value: "#DLP!" });
    expect(model.errorExplanation()?.title).toBe("Blocked by data loss prevention");

    model.setActiveCell({ address: "A1", input: "", value: "#AI!" });
    expect(model.errorExplanation()?.title).toBe("AI error");
  });

  it("returns explanations for newer Excel error codes", () => {
    const model = new FormulaBarModel();
    model.setActiveCell({ address: "A1", input: "", value: "#CALC!" });
    expect(model.errorExplanation()?.title).toBe("Calculation error");

    model.setActiveCell({ address: "A1", input: "", value: "#CONNECT!" });
    expect(model.errorExplanation()?.title).toBe("Connection error");

    model.setActiveCell({ address: "A1", input: "", value: "#FIELD!" });
    expect(model.errorExplanation()?.title).toBe("Invalid field");

    model.setActiveCell({ address: "A1", input: "", value: "#BLOCKED!" });
    expect(model.errorExplanation()?.title).toBe("Blocked");

    model.setActiveCell({ address: "A1", input: "", value: "#UNKNOWN!" });
    expect(model.errorExplanation()?.title).toBe("Unknown error");

    model.setActiveCell({ address: "A1", input: "", value: "#NULL!" });
    expect(model.errorExplanation()?.title).toBe("Null intersection");
  });
});
