import { describe, expect, it, vi } from "vitest";
import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

describe("detect_anomalies isolation_forest", () => {
  it("detects an obvious outlier and is deterministic", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A5",
        values: [[1], [1], [2], [2], [100]]
      }
    });

    const call = {
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A5", method: "isolation_forest" }
    } as const;

    const first = await executor.execute(call);
    const second = await executor.execute(call);

    expect(first.ok).toBe(true);
    expect(first.tool).toBe("detect_anomalies");
    if (!first.ok || first.tool !== "detect_anomalies") throw new Error("Unexpected tool result");

    expect(second.ok).toBe(true);
    expect(second.tool).toBe("detect_anomalies");
    if (!second.ok || second.tool !== "detect_anomalies") throw new Error("Unexpected tool result");

    expect(second.data).toEqual(first.data);

    if (!first.data || first.data.method !== "isolation_forest") throw new Error("Unexpected anomaly result");

    expect(first.data.anomalies.length).toBeGreaterThan(0);
    expect(first.data.anomalies[0]?.cell).toBe("Sheet1!A5");

    const outlier = first.data.anomalies.find((anomaly) => anomaly.cell === "Sheet1!A5");
    expect(outlier?.value).toBe(100);
    expect(outlier?.score).toBeGreaterThanOrEqual(0.65);
    expect(outlier?.score).toBeLessThanOrEqual(1);
  });

  it("treats threshold > 1 as top N anomalies", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A5",
        values: [[1], [1], [2], [2], [100]]
      }
    });

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A5", method: "isolation_forest", threshold: 2 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");

    if (!result.data || result.data.method !== "isolation_forest") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toHaveLength(2);
    expect(result.data.anomalies[0]?.cell).toBe("Sheet1!A5");
  });

  it("avoids Array.from({length: n}) allocations per tree", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const n = 2048;
    const values: number[][] = [];
    for (let i = 0; i < n; i++) values.push([i]);
    values[n - 1]![0] = 1_000_000;

    await executor.execute({
      name: "set_range",
      parameters: {
        range: `Sheet1!A1:A${n}`,
        values
      }
    });

    const fromSpy = vi.spyOn(Array, "from");
    try {
      const result = await executor.execute({
        name: "detect_anomalies",
        parameters: { range: `Sheet1!A1:A${n}`, method: "isolation_forest" }
      });
      expect(result.ok).toBe(true);

      const fullRangeAllocations = fromSpy.mock.calls.filter(([arg]) => {
        if (!arg || typeof arg !== "object") return false;
        if (Array.isArray(arg)) return false;
        return "length" in arg && (arg as any).length === n;
      });
      expect(fullRangeAllocations.length).toBeLessThan(5);
    } finally {
      fromSpy.mockRestore();
    }
  });
});
