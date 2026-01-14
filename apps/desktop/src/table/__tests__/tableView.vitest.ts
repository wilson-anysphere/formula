import { describe, expect, it } from "vitest";

import { distinctColumnValues } from "../tableView";

describe("distinctColumnValues", () => {
  it("sorts blank values last (Excel-like)", () => {
    const rows = [
      { row: 0, values: [""] },
      { row: 1, values: ["b"] },
      { row: 2, values: ["a"] },
    ];

    expect(distinctColumnValues(rows as any, 0)).toEqual(["a", "b", ""]);
  });
});

