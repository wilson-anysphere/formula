import { describe, expect, it, vi } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import type { SpreadsheetApi } from "../src/spreadsheet/api.js";

describe("detect_anomalies zscore/iqr", () => {
  it("zscore flags a clear outlier with the expected score", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A4",
        values: [[1], [1], [1], [100]],
      },
    });

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A4", method: "zscore", threshold: 1.4 },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "zscore") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toHaveLength(1);
    expect(result.data.anomalies[0]?.cell).toBe("Sheet1!A4");
    expect(result.data.anomalies[0]?.value).toBe(100);
    expect(result.data.anomalies[0]?.score).toBeCloseTo(1.5, 10);
  });

  it("iqr flags an outlier and only formats cell refs for anomalies", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A4",
        values: [[1], [1], [1], [100]],
      },
    });

    const formatSpy = vi.spyOn(executor as any, "formatCellForUser");

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A4", method: "iqr", threshold: 1.5 },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "iqr") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toHaveLength(1);
    expect(result.data.anomalies[0]?.cell).toBe("Sheet1!A4");
    expect(result.data.anomalies[0]?.value).toBe(100);

    // Regression: prior implementations formatted A1 cell refs for every numeric cell in the
    // range up-front. We only need to format refs for the returned anomalies.
    expect(formatSpy).toHaveBeenCalledTimes(1);
  });

  it("zscore does not format cell refs for large ranges when no anomalies are returned", async () => {
    const sharedOne = { value: 1 };
    const sharedTwo = { value: 2 };

    // 1000 rows x 200 columns = 200,000 cells (default `max_tool_range_cells`).
    const row = new Array(200).fill(sharedOne);
    row[0] = sharedTwo;
    const cells = new Array(1000).fill(row);

    const spreadsheet: SpreadsheetApi = {
      listSheets: () => ["Sheet1"],
      listNonEmptyCells: () => [],
      getCell: () => ({ value: null }),
      setCell: () => undefined,
      readRange: () => cells as any,
      writeRange: () => undefined,
      applyFormatting: () => 0,
      getLastUsedRow: () => 0,
      clone: () => spreadsheet,
    };

    const executor = new ToolExecutor(spreadsheet);
    const formatSpy = vi.spyOn(executor as any, "formatCellForUser");

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: {
        range: "Sheet1!A1:GR1000",
        method: "zscore",
        threshold: 100,
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "zscore") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toEqual([]);
    expect(formatSpy).toHaveBeenCalledTimes(0);
  });
});

