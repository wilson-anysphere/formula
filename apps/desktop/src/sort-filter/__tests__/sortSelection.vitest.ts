import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { applySortSpecToSelection, sortRangeRowsInDocument, sortSelection } from "../sortSelection.js";

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

  it("sorts formula cells using caller-provided computed values (instead of formula text)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";
    // A1:B3 (key column A contains formulas; column B is a payload to verify row movement)
    doc.setRangeValues(sheetId, { row: 0, col: 0 }, [
      [
        { formula: "=2", styleId: 1 },
        { value: "two", styleId: 2 },
      ],
      [
        { formula: "=10", styleId: 3 },
        { value: "ten", styleId: 4 },
      ],
      [
        { formula: "=1", styleId: 5 },
        { value: "one", styleId: 6 },
      ],
    ]);

    const computedByRow = new Map<number, number>([
      [0, 2],
      [1, 10],
      [2, 1],
    ]);

    const result = sortRangeRowsInDocument(
      doc,
      sheetId,
      { startRow: 0, endRow: 2, startCol: 0, endCol: 1 },
      { row: 0, col: 0 },
      {
        order: "ascending",
        getCellValue: ({ row }) => computedByRow.get(row) ?? null,
      },
    );
    expect(result.applied).toBe(true);

    // Computed sort keys: [2,10,1] => sorted order should be [1,2,10]
    expect(doc.getCell(sheetId, { row: 0, col: 0 })).toMatchObject({ value: null, formula: "=1", styleId: 5 });
    expect(doc.getCell(sheetId, { row: 0, col: 1 })).toMatchObject({ value: "one", styleId: 6 });

    expect(doc.getCell(sheetId, { row: 1, col: 0 })).toMatchObject({ value: null, formula: "=2", styleId: 1 });
    expect(doc.getCell(sheetId, { row: 1, col: 1 })).toMatchObject({ value: "two", styleId: 2 });

    expect(doc.getCell(sheetId, { row: 2, col: 0 })).toMatchObject({ value: null, formula: "=10", styleId: 3 });
    expect(doc.getCell(sheetId, { row: 2, col: 1 })).toMatchObject({ value: "ten", styleId: 4 });
  });

  it("falls back to comparing formula text when getCellValue is omitted (regression)", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";
    // A1:B3 (same fixture as above, but we intentionally omit `getCellValue`)
    doc.setRangeValues(sheetId, { row: 0, col: 0 }, [
      [
        { formula: "=2", styleId: 1 },
        { value: "two", styleId: 2 },
      ],
      [
        { formula: "=10", styleId: 3 },
        { value: "ten", styleId: 4 },
      ],
      [
        { formula: "=1", styleId: 5 },
        { value: "one", styleId: 6 },
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

    // Formula text sort keys: ["=2", "=10", "=1"] => ["=1", "=10", "=2"]
    expect(doc.getCell(sheetId, { row: 0, col: 0 })).toMatchObject({ value: null, formula: "=1", styleId: 5 });
    expect(doc.getCell(sheetId, { row: 0, col: 1 })).toMatchObject({ value: "one", styleId: 6 });

    expect(doc.getCell(sheetId, { row: 1, col: 0 })).toMatchObject({ value: null, formula: "=10", styleId: 3 });
    expect(doc.getCell(sheetId, { row: 1, col: 1 })).toMatchObject({ value: "ten", styleId: 4 });

    expect(doc.getCell(sheetId, { row: 2, col: 0 })).toMatchObject({ value: null, formula: "=2", styleId: 1 });
    expect(doc.getCell(sheetId, { row: 2, col: 1 })).toMatchObject({ value: "two", styleId: 2 });
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

  it("treats in-cell image payloads without alt text as blank sort keys", () => {
    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // A1:A3
    doc.setRangeValues(sheetId, { row: 0, col: 0 }, [
      [{ value: { type: "image", value: { imageId: "img-1" } }, styleId: 1 }],
      [{ value: "b", styleId: 2 }],
      [{ value: "a", styleId: 3 }],
    ]);

    const result = sortRangeRowsInDocument(
      doc,
      sheetId,
      { startRow: 0, endRow: 2, startCol: 0, endCol: 0 },
      { row: 0, col: 0 },
      { order: "ascending" },
    );
    expect(result.applied).toBe(true);

    // The image row should sort as blank (last), rather than being compared as "[object Object]".
    expect(doc.getCell(sheetId, { row: 0, col: 0 })).toMatchObject({ value: "a", styleId: 3 });
    expect(doc.getCell(sheetId, { row: 1, col: 0 })).toMatchObject({ value: "b", styleId: 2 });
    expect(doc.getCell(sheetId, { row: 2, col: 0 })).toMatchObject({
      value: { type: "image", value: { imageId: "img-1" } },
      styleId: 1,
    });
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

  it("does not resurrect deleted sheets when given a stale sheet id (no phantom creation)", () => {
    const doc = new DocumentController();

    // Ensure Sheet1 exists so deleting Sheet2 doesn't trip the last-sheet guard.
    doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    const result = sortRangeRowsInDocument(
      doc,
      "Sheet2",
      { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      { row: 0, col: 0 },
      { order: "ascending" },
    );

    expect(result.applied).toBe(false);
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
  });
});

describe("applySortSpecToSelection", () => {
  it("supports multi-key stable sorts and preserves the header row", () => {
    const doc = new DocumentController();

    doc.setRangeValues("Sheet1", "A1:C5", [
      ["Letter", "Number", "Id"],
      ["A", 1, "first"],
      ["A", 1, "second"],
      ["A", 0, "third"],
      ["B", 1, "fourth"],
    ]);

    const ok = applySortSpecToSelection({
      doc,
      sheetId: "Sheet1",
      selection: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
      spec: {
        hasHeader: true,
        keys: [
          { column: 0, order: "ascending" },
          { column: 1, order: "ascending" },
        ],
      },
      getCellValue: (cell) => {
        const state = doc.getCell("Sheet1", cell) as { value: any };
        return (state?.value ?? null) as any;
      },
    });

    expect(ok).toBe(true);

    const readRow = (row: number): any[] => {
      const out: any[] = [];
      for (let col = 0; col < 3; col += 1) {
        const state = doc.getCell("Sheet1", { row, col }) as { value: any };
        out.push(state?.value ?? null);
      }
      return out;
    };

    // Header row preserved.
    expect(readRow(0)).toEqual(["Letter", "Number", "Id"]);

    // Sorted rows:
    //   A,0,third
    //   A,1,first   (stable among equal keys)
    //   A,1,second
    //   B,1,fourth
    expect(readRow(1)).toEqual(["A", 0, "third"]);
    expect(readRow(2)).toEqual(["A", 1, "first"]);
    expect(readRow(3)).toEqual(["A", 1, "second"]);
    expect(readRow(4)).toEqual(["B", 1, "fourth"]);
  });

  it("does not resurrect deleted sheets when given a stale sheet id (no phantom creation)", () => {
    const doc = new DocumentController();

    // Ensure Sheet1 exists so deleting Sheet2 doesn't trip the last-sheet guard.
    doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    const ok = applySortSpecToSelection({
      doc,
      sheetId: "Sheet2",
      selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      spec: { hasHeader: false, keys: [{ column: 0, order: "ascending" }] },
      getCellValue: () => null,
    });

    expect(ok).toBe(false);
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
  });

  it("moves formulas + formatting (styleId) with rows when sorting", () => {
    const doc = new DocumentController();

    doc.setRangeValues("Sheet1", { row: 0, col: 0 }, [
      [
        { value: "Letter", styleId: 10 },
        { value: "Number", styleId: 11 },
        { value: "Payload", styleId: 12 },
      ],
      [
        { value: "A", styleId: 20 },
        { value: 2, styleId: 21 },
        { formula: "=1+1", styleId: 22 },
      ],
      [
        { value: "A", styleId: 30 },
        { value: 1, styleId: 31 },
        { value: "second", styleId: 32 },
      ],
      [
        { value: "B", styleId: 40 },
        { value: 0, styleId: 41 },
        { value: "third", styleId: 42 },
      ],
    ]);

    const ok = applySortSpecToSelection({
      doc,
      sheetId: "Sheet1",
      selection: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 },
      spec: {
        hasHeader: true,
        keys: [
          { column: 0, order: "ascending" },
          { column: 1, order: "ascending" },
        ],
      },
      getCellValue: (cell) => {
        const state = doc.getCell("Sheet1", cell) as { value: any };
        return (state?.value ?? null) as any;
      },
    });

    expect(ok).toBe(true);

    // Header row remains fixed.
    expect(doc.getCell("Sheet1", { row: 0, col: 0 })).toMatchObject({ value: "Letter", styleId: 10 });
    expect(doc.getCell("Sheet1", { row: 0, col: 1 })).toMatchObject({ value: "Number", styleId: 11 });
    expect(doc.getCell("Sheet1", { row: 0, col: 2 })).toMatchObject({ value: "Payload", styleId: 12 });

    // Sorted rows should be:
    // - A,1,second (styleId 30/31/32)
    // - A,2,=1+1   (styleId 20/21/22)
    // - B,0,third  (styleId 40/41/42)
    expect(doc.getCell("Sheet1", { row: 1, col: 0 })).toMatchObject({ value: "A", styleId: 30 });
    expect(doc.getCell("Sheet1", { row: 1, col: 1 })).toMatchObject({ value: 1, styleId: 31 });
    expect(doc.getCell("Sheet1", { row: 1, col: 2 })).toMatchObject({ value: "second", styleId: 32 });

    expect(doc.getCell("Sheet1", { row: 2, col: 0 })).toMatchObject({ value: "A", styleId: 20 });
    expect(doc.getCell("Sheet1", { row: 2, col: 1 })).toMatchObject({ value: 2, styleId: 21 });
    expect(doc.getCell("Sheet1", { row: 2, col: 2 })).toMatchObject({ value: null, formula: "=1+1", styleId: 22 });

    expect(doc.getCell("Sheet1", { row: 3, col: 0 })).toMatchObject({ value: "B", styleId: 40 });
    expect(doc.getCell("Sheet1", { row: 3, col: 1 })).toMatchObject({ value: 0, styleId: 41 });
    expect(doc.getCell("Sheet1", { row: 3, col: 2 })).toMatchObject({ value: "third", styleId: 42 });
  });
});

describe("sortSelection (UI wrapper)", () => {
  it("returns early in read-only mode (does not consult selection state)", () => {
    const app = {
      isReadOnly: () => true,
      // These should never be consulted when read-only.
      getSelectionRanges: () => {
        throw new Error("getSelectionRanges should not be called");
      },
      getActiveCell: () => {
        throw new Error("getActiveCell should not be called");
      },
      focus: () => {},
    } as any;

    expect(() => sortSelection(app, { order: "ascending" })).not.toThrow();
  });
});
