import { describe, expect, it } from "vitest";

import { FormulaBarModel } from "./FormulaBarModel.js";
import { parseA1Range } from "../spreadsheet/a1.js";

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
});
