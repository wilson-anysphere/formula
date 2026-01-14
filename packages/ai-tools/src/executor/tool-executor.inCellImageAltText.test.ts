import { describe, expect, it } from "vitest";

import { ToolExecutor } from "./tool-executor.ts";
import { parseA1Cell } from "../spreadsheet/a1.ts";
import { InMemoryWorkbook } from "../spreadsheet/in-memory-workbook.ts";

describe("ToolExecutor rich value normalization", () => {
  it("read_range trims in-cell image alt text", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), {
      value: { imageId: "img1", altText: "  Logo  " },
    });

    const executor = new ToolExecutor(workbook, {});
    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["Logo"]]);
  });
});

