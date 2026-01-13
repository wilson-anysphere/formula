import { describe, expect, it } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { parseA1Range } from "../src/spreadsheet/a1.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

describe("ToolExecutor numeric parsing (spreadsheet-formatted strings)", () => {
  it("compute_statistics parses numeric strings like '1,200' and '$300'", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A2",
        values: [["1,200"], ["$300"]],
      },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:A2", measures: ["sum", "mean"] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

    expect(result.data?.statistics.sum).toBe(1500);
    expect(result.data?.statistics.mean).toBe(750);
  });

  it("filter_range numeric operators match spreadsheet-formatted numeric strings", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A3",
        values: [["500"], ["1,200"], ["$300"]],
      },
    });

    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A3",
        criteria: [{ column: "A", operator: "greater", value: "1,000" }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");

    expect(result.data?.count).toBe(1);
    expect(result.data?.matching_rows).toEqual([2]);
  });

  it("sort_range orders spreadsheet-formatted numeric strings numerically", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:B3",
        values: [
          ["Name", "Amount"],
          ["A", "1,200"],
          ["B", "$300"],
        ],
      },
    });

    const result = await executor.execute({
      name: "sort_range",
      parameters: {
        range: "Sheet1!A1:B3",
        sort_by: [{ column: "B", order: "asc" }],
        has_header: true,
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("sort_range");
    if (!result.ok || result.tool !== "sort_range") throw new Error("Unexpected tool result");

    const values = workbook
      .readRange(parseA1Range("Sheet1!A1:B3"))
      .map((row) => row.map((cell) => cell.value));

    expect(values).toEqual([
      ["Name", "Amount"],
      ["B", "$300"],
      ["A", "1,200"],
    ]);
  });
});

