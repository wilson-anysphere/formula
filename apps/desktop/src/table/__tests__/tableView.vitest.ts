import { describe, expect, it } from "vitest";

import { distinctColumnValues, type TableViewRow } from "../tableView";

describe("distinctColumnValues", () => {
  it("sorts blank values last (Excel-like)", () => {
    const rows: TableViewRow[] = [
      { row: 0, values: [""] },
      { row: 1, values: ["b"] },
      { row: 2, values: ["a"] },
    ];

    expect(distinctColumnValues(rows, 0)).toEqual(["a", "b", ""]);
  });

  it("sorts numeric strings using numeric collation (Excel-like)", () => {
    const rows: TableViewRow[] = [
      { row: 0, values: ["2"] },
      { row: 1, values: ["10"] },
      { row: 2, values: ["1"] },
    ];

    expect(distinctColumnValues(rows, 0)).toEqual(["1", "2", "10"]);
  });

  it("includes null/undefined values as blanks", () => {
    const rows: TableViewRow[] = [
      { row: 0, values: [null] },
      { row: 1, values: ["x"] },
    ];

    expect(distinctColumnValues(rows, 0)).toEqual(["x", ""]);
  });
});
