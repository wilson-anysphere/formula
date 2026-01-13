import { describe, expect, it } from "vitest";

import { RibbonAutoFilterStore, computeFilterHiddenRows, computeUniqueFilterValues } from "../ribbonAutoFilter";

describe("RibbonAutoFilterStore", () => {
  it("stores filters keyed by (sheetId, rangeA1) and can query by cell", () => {
    const store = new RibbonAutoFilterStore();

    expect(store.hasAny("Sheet1")).toBe(false);

    store.set("Sheet1", {
      rangeA1: "A1:B4",
      headerRows: 1,
      filterColumns: [{ colId: 0, values: ["x"] }],
    });

    expect(store.hasAny("Sheet1")).toBe(true);
    expect(store.get("Sheet1", "A1:B4")?.filterColumns).toEqual([{ colId: 0, values: ["x"] }]);

    // Inside range.
    expect(store.findByCell("Sheet1", { row: 2, col: 1 })?.rangeA1).toBe("A1:B4");
    // Outside range.
    expect(store.findByCell("Sheet1", { row: 10, col: 0 })).toBeUndefined();

    store.delete("Sheet1", "A1:B4");
    expect(store.hasAny("Sheet1")).toBe(false);
  });
});

describe("Ribbon AutoFilter row computation", () => {
  it("computes unique values for the active column (excluding header rows)", () => {
    const values = [
      ["Header"],
      ["x"],
      ["y"],
      ["x"],
    ];
    const getValue = (row: number, col: number) => values[row]?.[col] ?? "";

    expect(
      computeUniqueFilterValues({
        range: { startRow: 0, endRow: 3, startCol: 0, endCol: 0 },
        headerRows: 1,
        colId: 0,
        getValue,
      }),
    ).toEqual(["x", "y"]);
  });

  it("hides rows that do not match selected values", () => {
    const values = [
      ["Header"],
      ["x"],
      ["y"],
      ["x"],
    ];
    const getValue = (row: number, col: number) => values[row]?.[col] ?? "";

    expect(
      computeFilterHiddenRows({
        range: { startRow: 0, endRow: 3, startCol: 0, endCol: 0 },
        headerRows: 1,
        filterColumns: [{ colId: 0, values: ["x"] }],
        getValue,
      }),
    ).toEqual([2]);
  });

  it("treats empty selections as hiding all data rows (Excel-like)", () => {
    const values = [
      ["Header"],
      ["x"],
      ["y"],
      ["x"],
    ];
    const getValue = (row: number, col: number) => values[row]?.[col] ?? "";

    expect(
      computeFilterHiddenRows({
        range: { startRow: 0, endRow: 3, startCol: 0, endCol: 0 },
        headerRows: 1,
        filterColumns: [{ colId: 0, values: [] }],
        getValue,
      }),
    ).toEqual([1, 2, 3]);
  });

  it("combines multiple column filters with AND semantics", () => {
    const values = [
      ["H1", "H2"],
      ["x", "1"],
      ["x", "2"],
      ["y", "1"],
      ["x", "1"],
    ];
    const getValue = (row: number, col: number) => values[row]?.[col] ?? "";

    expect(
      computeFilterHiddenRows({
        range: { startRow: 0, endRow: 4, startCol: 0, endCol: 1 },
        headerRows: 1,
        filterColumns: [
          { colId: 0, values: ["x"] },
          { colId: 1, values: ["1"] },
        ],
        getValue,
      }),
    ).toEqual([2, 3]);
  });
});

