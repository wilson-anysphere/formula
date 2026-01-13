import { describe, expect, it } from "vitest";
import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

describe("filter_range between validation", () => {
  it("returns validation_error when operator 'between' is missing value2", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A5",
        criteria: [{ column: "A", operator: "between", value: 1 }]
      }
    });

    expect(result.ok).toBe(false);
    expect(result.tool).toBe("filter_range");
    expect(result.error?.code).toBe("validation_error");
    // Ensure the validation feedback is actionable.
    expect((result.error?.details as any)?.fieldErrors?.criteria?.join("\n") ?? "").toMatch(/value2/i);
  });

  it("accepts valid 'between' criteria and returns expected matches", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A5",
        values: [[1], [2], [3], [4], [5]]
      }
    });

    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A5",
        criteria: [{ column: "A", operator: "between", value: 2, value2: 4 }]
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");

    expect(result.data?.range).toBe("Sheet1!A1:A5");
    expect(result.data?.count).toBe(3);
    expect(result.data?.matching_rows).toEqual([2, 3, 4]);
  });
});

