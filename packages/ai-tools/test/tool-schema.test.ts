import { describe, expect, it } from "vitest";
import { validateToolCall } from "../src/tool-schema.js";

describe("create_pivot_table aggregation normalization", () => {
  it("accepts common spellings and normalizes to Rust camelCase aggregation strings", () => {
    const call = validateToolCall({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:B2",
        rows: ["Region"],
        values: [
          { field: "Sales", aggregation: "stddevp" },
          { field: "Sales", aggregation: "VARP" },
          { field: "Sales", aggregation: "countnumbers" }
        ],
        destination: "Sheet1!D1"
      }
    });

    expect(call.name).toBe("create_pivot_table");
    const params = call.parameters as any;
    expect(params.values[0].aggregation).toBe("stdDevP");
    expect(params.values[1].aggregation).toBe("varP");
    expect(params.values[2].aggregation).toBe("countNumbers");
  });
});

describe("set_range schema (large inputs)", () => {
  it("validates a very tall values matrix without spread argument overflow", () => {
    // Regression test: `Math.max(...rows.map(r => r.length))` will throw in V8 once the
    // row count crosses the engine's argument limits.
    const rows = 130_000;
    const values: Array<Array<null>> = Array.from({ length: rows }, () => []);
    values[rows - 1] = [null];

    const call = validateToolCall({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1",
        values
      }
    });

    expect(call.name).toBe("set_range");
  });
});
