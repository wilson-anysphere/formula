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
});
