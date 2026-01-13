import { describe, expect, it } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";

describe("ToolExecutor compute_statistics (streaming + distribution measures)", () => {
  it("computes correct results for a mixed dataset (numbers, numeric strings, formulas, non-numeric)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "2" });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A4"), { value: "nope" });
    workbook.setCell(parseA1Cell("Sheet1!A5"), { value: null });
    // Formula cells are ignored by compute_statistics.
    workbook.setCell(parseA1Cell("Sheet1!A6"), { value: 4, formula: "=2+2" });

    const executor = new ToolExecutor(workbook);
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: {
        range: "Sheet1!A1:A6",
        measures: ["count", "sum", "mean", "min", "max", "variance", "stdev", "median", "mode", "quartiles"],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

    const stats = result.data?.statistics ?? {};
    expect(stats.count).toBe(3);
    expect(stats.sum).toBe(5);
    expect(stats.mean).toBeCloseTo(5 / 3, 12);
    expect(stats.min).toBe(1);
    expect(stats.max).toBe(2);
    expect(stats.variance).toBeCloseTo(1 / 3, 12);
    expect(stats.stdev).toBeCloseTo(Math.sqrt(1 / 3), 12);
    expect(stats.median).toBe(2);
    expect(stats.mode).toBe(2);
    expect(stats.q1).toBeCloseTo(1.5, 12);
    expect(stats.q2).toBeCloseTo(2, 12);
    expect(stats.q3).toBeCloseTo(2, 12);
  });

  it("matches legacy semantics for small-N variance/stdev (n=1 -> 0, n=0 -> null)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 5 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "x" });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: null });

    const executor = new ToolExecutor(workbook);
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:A3", measures: ["count", "variance", "stdev", "sum", "mean", "min", "max"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({
      count: 1,
      variance: 0,
      stdev: 0,
      sum: 5,
      mean: 5,
      min: 5,
      max: 5,
    });

    const resultEmpty = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A2:A3", measures: ["count", "variance", "stdev", "sum", "mean", "min", "max"] },
    });
    expect(resultEmpty.ok).toBe(true);
    if (!resultEmpty.ok || resultEmpty.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(resultEmpty.data?.statistics).toEqual({
      count: 0,
      variance: null,
      stdev: null,
      sum: null,
      mean: null,
      min: null,
      max: null,
    });
  });

  it("computes correlation without requiring materialized pairs", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: 4 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: 6 });

    const executor = new ToolExecutor(workbook);
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B3", measures: ["correlation"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics.correlation).toBeCloseTo(1, 12);
  });
});

