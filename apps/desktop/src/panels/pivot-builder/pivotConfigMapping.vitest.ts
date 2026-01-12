import { describe, expect, it } from "vitest";

import type { PivotTableConfig } from "./types";
import { toRustPivotConfig } from "./pivotConfigMapping.js";

describe("toRustPivotConfig", () => {
  it("maps field arrays and grand totals using Rust serde casing", () => {
    const cfg: PivotTableConfig = {
      rowFields: [{ sourceField: "Category" }],
      columnFields: [{ sourceField: "Region" }],
      valueFields: [{ sourceField: "Amount", name: "Sum of Amount", aggregation: "sum" }],
      filterFields: [{ sourceField: "Year" }],
      layout: "tabular",
      subtotals: "none",
      grandTotals: { rows: false, columns: true },
    };

    expect(toRustPivotConfig(cfg)).toEqual({
      rowFields: [{ sourceField: "Category" }],
      columnFields: [{ sourceField: "Region" }],
      valueFields: [{ sourceField: "Amount", name: "Sum of Amount", aggregation: "sum" }],
      filterFields: [{ sourceField: "Year" }],
      calculatedFields: [],
      calculatedItems: [],
      layout: "tabular",
      subtotals: "none",
      grandTotals: { rows: false, columns: true },
    });
  });
});
