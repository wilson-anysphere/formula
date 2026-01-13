import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { sortRangeRowsInDocument } from "../sortSelection.js";

describe("sortRangeRowsInDocument", () => {
  it("sorts rows in a selection and moves style/formula state with the row", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // A1:B3
    doc.setRangeValues(sheetId, { row: 0, col: 0 }, [
      [
        { value: 2, styleId: 101 },
        { value: "b", styleId: 102 },
      ],
      [
        { value: 10, styleId: 201 },
        { formula: "=1+1", styleId: 202 },
      ],
      [
        { value: 1, styleId: 301 },
        { value: "a", styleId: 302 },
      ],
    ]);

    const result = sortRangeRowsInDocument(
      doc,
      sheetId,
      { startRow: 0, endRow: 2, startCol: 0, endCol: 1 },
      { row: 0, col: 0 },
      { order: "ascending" },
    );
    expect(result.applied).toBe(true);

    // Rows should sort by column A: [1, 2, 10]
    expect(doc.getCell(sheetId, { row: 0, col: 0 })).toMatchObject({ value: 1, styleId: 301 });
    expect(doc.getCell(sheetId, { row: 0, col: 1 })).toMatchObject({ value: "a", styleId: 302 });

    expect(doc.getCell(sheetId, { row: 1, col: 0 })).toMatchObject({ value: 2, styleId: 101 });
    expect(doc.getCell(sheetId, { row: 1, col: 1 })).toMatchObject({ value: "b", styleId: 102 });

    expect(doc.getCell(sheetId, { row: 2, col: 0 })).toMatchObject({ value: 10, styleId: 201 });
    expect(doc.getCell(sheetId, { row: 2, col: 1 })).toMatchObject({ value: null, formula: "=1+1", styleId: 202 });
  });

  it("keeps blanks last (even for descending sort)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // A1:A4
    doc.setRangeValues(sheetId, { row: 0, col: 0 }, [
      [{ value: 3, styleId: 1 }],
      [{ value: null, styleId: 2 }],
      [{ value: 2, styleId: 3 }],
      [{ value: "", styleId: 4 }],
    ]);

    const result = sortRangeRowsInDocument(
      doc,
      sheetId,
      { startRow: 0, endRow: 3, startCol: 0, endCol: 0 },
      { row: 0, col: 0 },
      { order: "descending" },
    );
    expect(result.applied).toBe(true);

    // Descending numeric keys first: 3,2 then blanks.
    expect(doc.getCell(sheetId, { row: 0, col: 0 })).toMatchObject({ value: 3, styleId: 1 });
    expect(doc.getCell(sheetId, { row: 1, col: 0 })).toMatchObject({ value: 2, styleId: 3 });
    expect(doc.getCell(sheetId, { row: 2, col: 0 }).value).toBeNull();
    expect(doc.getCell(sheetId, { row: 3, col: 0 }).value).toBe("");
  });

  it("rejects selections larger than the safety threshold", () => {
    const doc = new DocumentController();

    const result = sortRangeRowsInDocument(
      doc,
      "Sheet1",
      { startRow: 0, endRow: 50_000, startCol: 0, endCol: 0 },
      { row: 0, col: 0 },
      { order: "ascending", maxCells: 50_000 },
    );

    expect(result).toMatchObject({ applied: false, reason: "tooLarge" });
    expect(doc.getStackDepths()).toEqual({ undo: 0, redo: 0 });
  });
});

