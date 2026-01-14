import { describe, expect, it, vi } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

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

  it("preserves mode tie-breaking semantics when median/quartiles are also requested", async () => {
    // Tie case: 1 and 2 both appear twice. Legacy `mode()` resolves ties by first occurrence
    // in scan order (Map insertion order), so this should be 2 (not 1).
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    // Scan order is A1->D1: [2,1,1,2].
    await executor.execute({
      name: "set_range",
      parameters: { range: "Sheet1!A1:D1", values: [[2, 1, 1, 2]] },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:D1", measures: ["mode", "median", "quartiles"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics.mode).toBe(2);
    expect(result.data?.statistics.median).toBeCloseTo(1.5, 12);
    expect(result.data?.statistics.q2).toBeCloseTo(1.5, 12);
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

  it("short-circuits correlation-only requests for non-2-column ranges without reading the range", async () => {
    const spreadsheet: any = {
      readRange: vi.fn(() => {
        throw new Error("readRange should not be called for invalid correlation ranges");
      }),
    };
    const executor = new ToolExecutor(spreadsheet);
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:C2", measures: ["correlation"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics.correlation).toBeNull();
    expect(spreadsheet.readRange).not.toHaveBeenCalled();
  });

  it("short-circuits correlation-only requests for non-2-column ranges under DLP REDACT without reading the range (but still counts redactions)", async () => {
    const spreadsheet: any = {
      readRange: vi.fn(() => {
        throw new Error("readRange should not be called for invalid correlation ranges under DLP REDACT");
      }),
    };

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(spreadsheet, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true,
            },
          },
        },
        classification_records: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
        audit_logger,
      },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:C2", measures: ["correlation"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics.correlation).toBeNull();
    expect(spreadsheet.readRange).not.toHaveBeenCalled();

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      tool: "compute_statistics",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:C2",
      redactedCellCount: 1,
    });
    expect(event.decision?.decision).toBe("redact");
  });
});
