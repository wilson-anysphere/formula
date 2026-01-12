import { describe, expect, it } from "vitest";

import { assignFormulaReferenceColors, extractFormulaReferences, FORMULA_REFERENCE_PALETTE } from "../src/formulaReferences";

describe("extractFormulaReferences", () => {
  it("extracts simple A1 references with stable indices", () => {
    const { references, activeIndex } = extractFormulaReferences("=A1+B1", 0, 0);
    expect(activeIndex).toBe(null);
    expect(references).toEqual([
      {
        text: "A1",
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, sheet: undefined },
        index: 0,
        start: 1,
        end: 3
      },
      {
        text: "B1",
        range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1, sheet: undefined },
        index: 1,
        start: 4,
        end: 6
      }
    ]);
  });

  it("parses sheet-qualified ranges", () => {
    const { references } = extractFormulaReferences("=SUM('My Sheet'!$A$1:$B$2)", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("'My Sheet'!$A$1:$B$2");
    expect(references[0]?.range).toEqual({ sheet: "My Sheet", startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  });

  it("parses sheet-qualified refs with escaped apostrophes", () => {
    const { references } = extractFormulaReferences("=SUM('O''Brien'!A1)", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("'O''Brien'!A1");
    expect(references[0]?.range).toEqual({ sheet: "O'Brien", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });

  it("parses unquoted Unicode sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=résumé!A1+数据!B2", 0, 0);
    expect(references).toHaveLength(2);
    expect(references[0]?.text).toBe("résumé!A1");
    expect(references[0]?.range).toEqual({ sheet: "résumé", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(references[1]?.text).toBe("数据!B2");
    expect(references[1]?.range).toEqual({ sheet: "数据", startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
  });

  it("does not treat invalid unquoted sheet names with spaces as sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=My Sheet!A1", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("A1");
    expect(references[0]?.range).toEqual({ sheet: undefined, startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });

  it("detects the active reference at the caret (including token end)", () => {
    // =A1+B1, caret after final "1" should count as being in B1.
    const input = "=A1+B1";
    const { activeIndex } = extractFormulaReferences(input, input.length, input.length);
    expect(activeIndex).toBe(1);
  });
});

describe("assignFormulaReferenceColors", () => {
  it("assigns palette colors by index on first pass", () => {
    const { references } = extractFormulaReferences("=A1+B1", 0, 0);
    const { colored } = assignFormulaReferenceColors(references, null);
    expect(colored.map((r) => r.color)).toEqual([FORMULA_REFERENCE_PALETTE[0], FORMULA_REFERENCE_PALETTE[1]]);
  });

  it("reuses the same color for repeated references within a formula", () => {
    const { references } = extractFormulaReferences("=A1+A1", 0, 0);
    const { colored } = assignFormulaReferenceColors(references, null);
    expect(colored).toHaveLength(2);
    expect(colored[0]?.color).toBe(FORMULA_REFERENCE_PALETTE[0]);
    expect(colored[1]?.color).toBe(FORMULA_REFERENCE_PALETTE[0]);
  });

  it("reuses colors for the same reference text across edits", () => {
    const first = extractFormulaReferences("=A1+B1", 0, 0).references;
    const { colored: coloredFirst, nextByText } = assignFormulaReferenceColors(first, null);

    const second = extractFormulaReferences("=B1+A1", 0, 0).references;
    const { colored: coloredSecond } = assignFormulaReferenceColors(second, nextByText);

    expect(coloredFirst.map((r) => [r.text, r.color])).toEqual([
      ["A1", FORMULA_REFERENCE_PALETTE[0]],
      ["B1", FORMULA_REFERENCE_PALETTE[1]]
    ]);
    expect(coloredSecond.map((r) => [r.text, r.color])).toEqual([
      ["B1", FORMULA_REFERENCE_PALETTE[1]],
      ["A1", FORMULA_REFERENCE_PALETTE[0]]
    ]);
  });

  it("preserves existing reference colors when a new reference is inserted earlier", () => {
    const initialRefs = extractFormulaReferences("=A1+B1", 0, 0).references;
    const { nextByText } = assignFormulaReferenceColors(initialRefs, null);

    const editedRefs = extractFormulaReferences("=C1+A1+B1", 0, 0).references;
    const { colored } = assignFormulaReferenceColors(editedRefs, nextByText);

    expect(colored.map((r) => [r.text, r.color])).toEqual([
      ["C1", FORMULA_REFERENCE_PALETTE[2]],
      ["A1", FORMULA_REFERENCE_PALETTE[0]],
      ["B1", FORMULA_REFERENCE_PALETTE[1]]
    ]);
  });
});
