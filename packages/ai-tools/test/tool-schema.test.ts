import { describe, expect, it } from "vitest";
import { TOOL_REGISTRY, validateToolCall } from "../src/tool-schema.js";

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

describe("tool JSON schema fidelity (matches Zod refinements)", () => {
  it("apply_formatting requires at least one formatting field (minProperties: 1)", () => {
    const formatSchema = (TOOL_REGISTRY.apply_formatting.jsonSchema as any).properties?.format;
    expect(formatSchema).toBeDefined();
    expect(formatSchema.minProperties).toBe(1);
  });

  it("filter_range requires value2 when operator is between (oneOf variant)", () => {
    const criteriaItemsSchema = (TOOL_REGISTRY.filter_range.jsonSchema as any).properties?.criteria?.items;
    expect(criteriaItemsSchema).toBeDefined();
    expect(criteriaItemsSchema.oneOf).toBeDefined();
    expect(Array.isArray(criteriaItemsSchema.oneOf)).toBe(true);

    const oneOf = criteriaItemsSchema.oneOf as any[];
    const betweenVariant = oneOf.find((variant) => variant?.properties?.operator?.enum?.includes?.("between"));
    expect(betweenVariant).toBeDefined();
    expect(betweenVariant.required).toContain("value2");

    const nonBetweenVariant = oneOf.find(
      (variant) => variant?.properties?.operator?.enum && !variant.properties.operator.enum.includes("between")
    );
    expect(nonBetweenVariant).toBeDefined();
    expect(nonBetweenVariant.required).toEqual(["column", "operator", "value"]);
  });

  it("includes minItems: 1 for create_pivot_table.rows and create_pivot_table.values (and requires destination)", () => {
    const schema = TOOL_REGISTRY.create_pivot_table.jsonSchema as any;
    expect(schema.properties.rows.minItems).toBe(1);
    expect(schema.properties.values.minItems).toBe(1);
    expect(schema.required).toContain("destination");
  });

  it("includes minItems: 1 for sort_range.sort_by", () => {
    const schema = TOOL_REGISTRY.sort_range.jsonSchema as any;
    expect(schema.properties.sort_by.minItems).toBe(1);
  });

  it("includes minItems: 1 for filter_range.criteria", () => {
    const schema = TOOL_REGISTRY.filter_range.jsonSchema as any;
    expect(schema.properties.criteria.minItems).toBe(1);
  });
});

describe("column label normalization", () => {
  it("accepts $-prefixed columns in sort_range and normalizes to uppercase letters", () => {
    const call = validateToolCall({
      name: "sort_range",
      parameters: {
        range: "Sheet1!A1:C10",
        sort_by: [{ column: "$b", order: "asc" }]
      }
    });

    expect(call.name).toBe("sort_range");
    const params = call.parameters as any;
    expect(params.sort_by[0].column).toBe("B");
  });

  it("accepts $-prefixed columns in filter_range and normalizes to uppercase letters", () => {
    const call = validateToolCall({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:C10",
        criteria: [{ column: "$b", operator: "equals", value: "x" }]
      }
    });

    expect(call.name).toBe("filter_range");
    const params = call.parameters as any;
    expect(params.criteria[0].column).toBe("B");
  });

  it("accepts $-prefixed columns in apply_formula_column and normalizes to uppercase letters", () => {
    const call = validateToolCall({
      name: "apply_formula_column",
      parameters: {
        column: "$b",
        formula_template: "=A{row}+1",
        start_row: 2
      }
    });

    expect(call.name).toBe("apply_formula_column");
    const params = call.parameters as any;
    expect(params.column).toBe("B");
  });
});
