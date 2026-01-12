import { describe, expect, it } from "vitest";

import { FormulaBarModel } from "./FormulaBarModel.js";
import { parseA1Range } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";

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
    expect(parseSheetQualifiedA1Range("A1:B2")).toEqual(parseA1Range("A1:B2"));
    expect(parseSheetQualifiedA1Range("Sheet2!A1:B2")).toEqual(parseA1Range("A1:B2"));
    expect(parseSheetQualifiedA1Range("'My Sheet'!A1")).toEqual(parseA1Range("A1"));
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
});
