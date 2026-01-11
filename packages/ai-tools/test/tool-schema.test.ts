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

