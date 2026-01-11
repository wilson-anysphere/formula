import { describe, expect, it } from "vitest";
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
});
