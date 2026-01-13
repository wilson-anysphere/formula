import { describe, expect, it } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

describe("ToolExecutor quantile helpers", () => {
  it("compute_statistics median/quartiles match the existing linear interpolation semantics", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A4",
        values: [[1], [2], [3], [4]],
      },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:A4", measures: ["median", "quartiles"] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

    // Existing quantile behavior (pos = (n-1)*q; linear interpolation between neighbors):
    // sorted=[1,2,3,4]
    // median/q2: pos=1.5 => 2 + 0.5*(3-2) = 2.5
    // q1: pos=0.75 => 1 + 0.75*(2-1) = 1.75
    // q3: pos=2.25 => 3 + 0.25*(4-3) = 3.25
    const stats = result.data?.statistics;
    expect(stats?.median).toBeCloseTo(2.5);
    expect(stats?.q1).toBeCloseTo(1.75);
    expect(stats?.q2).toBeCloseTo(2.5);
    expect(stats?.q3).toBeCloseTo(3.25);
  });

  it("detect_anomalies iqr uses quartiles consistently (no change in results)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A5",
        values: [[1], [2], [3], [4], [100]],
      },
    });

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A5", method: "iqr" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "iqr") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toEqual([{ cell: "Sheet1!A5", value: 100 }]);
  });
});

