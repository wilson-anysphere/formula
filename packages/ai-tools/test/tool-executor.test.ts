import { describe, expect, it } from "vitest";
import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell, parseA1Range } from "../src/spreadsheet/a1.js";

describe("ToolExecutor", () => {
  it("write_cell writes a scalar value", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 42 }
    });

    expect(result.ok).toBe(true);
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(42);
  });

  it("write_cell writes a formula when value starts with '='", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!B2", value: "=SUM(A1:A10)" }
    });

    const cell = workbook.getCell(parseA1Cell("Sheet1!B2"));
    expect(cell.value).toBeNull();
    expect(cell.formula).toBe("=SUM(A1:A10)");
  });

  it("set_range updates a rectangular range", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:B2",
        values: [
          [1, 2],
          [3, 4]
        ]
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("set_range");
    if (!result.ok || result.tool !== "set_range") throw new Error("Unexpected tool result");
    expect(result.data?.updated_cells).toBe(4);
    const range = parseA1Range("Sheet1!A1:B2");
    const values = workbook.readRange(range).map((row) => row.map((cell) => cell.value));
    expect(values).toEqual([
      [1, 2],
      [3, 4]
    ]);
  });

  it("apply_formula_column fills formulas down to the last used row when end_row = -1", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A3",
        values: [["Header"], [10], [20]]
      }
    });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: { column: "C", formula_template: "=A{row}*2", start_row: 2, end_row: -1 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("apply_formula_column");
    if (!result.ok || result.tool !== "apply_formula_column") throw new Error("Unexpected tool result");
    expect(result.data?.updated_cells).toBe(2);
    expect(workbook.getCell(parseA1Cell("Sheet1!C2")).formula).toBe("=A2*2");
    expect(workbook.getCell(parseA1Cell("Sheet1!C3")).formula).toBe("=A3*2");
  });

  it("accepts camelCase parameter aliases from docs examples", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 5 }
    });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: { column: "B", formulaTemplate: "=A{row}*10", startRow: 1, endRow: 1 }
    });

    expect(result.ok).toBe(true);
    expect(workbook.getCell(parseA1Cell("Sheet1!B1")).formula).toBe("=A1*10");
  });

  it("returns validation_error for invalid A1 references", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "NotACell", value: 1 }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("validation_error");
  });

  it("create_pivot_table writes a pivot output table", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:C5",
        values: [
          ["Region", "Product", "Sales"],
          ["East", "A", 100],
          ["East", "B", 150],
          ["West", "A", 200],
          ["West", "B", 250]
        ]
      }
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C5",
        rows: ["Region"],
        columns: ["Product"],
        values: [{ field: "Sales", aggregation: "sum" }],
        destination: "Sheet1!E1"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");

    const out = workbook
      .readRange(parseA1Range("Sheet1!E1:H4"))
      .map((row) => row.map((cell) => cell.value));

    expect(out).toEqual([
      ["Region", "A - Sum of Sales", "B - Sum of Sales", "Grand Total - Sum of Sales"],
      ["East", 100, 150, 250],
      ["West", 200, 250, 450],
      ["Grand Total", 300, 400, 700]
    ]);

    // Updating the source range should refresh the pivot output automatically.
    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!C2", value: 110 }
    });

    const refreshed = workbook
      .readRange(parseA1Range("Sheet1!E1:H4"))
      .map((row) => row.map((cell) => cell.value));

    expect(refreshed).toEqual([
      ["Region", "A - Sum of Sales", "B - Sum of Sales", "Grand Total - Sum of Sales"],
      ["East", 110, 150, 260],
      ["West", 200, 250, 450],
      ["Grand Total", 310, 400, 710]
    ]);
  });

  it("create_pivot_table supports variance/stddev aggregations", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:C5",
        values: [
          ["Region", "Product", "Sales"],
          ["East", "A", 100],
          ["East", "B", 150],
          ["West", "A", 200],
          ["West", "B", 250]
        ]
      }
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C5",
        rows: ["Region"],
        values: [
          { field: "Sales", aggregation: "varp" },
          { field: "Sales", aggregation: "stddevp" }
        ],
        destination: "Sheet1!E1"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");

    const out = workbook
      .readRange(parseA1Range("Sheet1!E1:G4"))
      .map((row) => row.map((cell) => cell.value));

    expect(out[0]).toEqual(["Region", "VarP of Sales", "StdDevP of Sales"]);
    expect(out[1]).toEqual(["East", 625, 25]);
    expect(out[2]).toEqual(["West", 625, 25]);

    // Grand total is based on all records; check it roughly matches expected values.
    expect(out[3]?.[0]).toBe("Grand Total");
    expect(out[3]?.[1]).toBeCloseTo(3125, 10);
    expect(out[3]?.[2]).toBeCloseTo(Math.sqrt(3125), 10);
  });
});
