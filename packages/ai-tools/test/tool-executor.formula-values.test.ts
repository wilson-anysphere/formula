import { describe, expect, it, vi } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../security/dlp/src/selectors.js";

describe("ToolExecutor include_formula_values", () => {
  it("read_range surfaces computed values for formula cells when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[2]]);
  });

  it("compute_statistics includes numeric values from formula cells when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 4 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B1", measures: ["mean", "sum", "count"] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({ mean: 3, sum: 6, count: 2 });
  });

  it("filter_range compares formula cells using computed values when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=1+1", value: 2 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A2",
        has_header: true,
        criteria: [{ column: "A", operator: "greater", value: 1 }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");
    expect(result.data?.matching_rows).toEqual([2]);
    expect(result.data?.count).toBe(1);
  });

  it("does not surface formula values under DLP REDACT decisions even when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 100 });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: 4 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      include_formula_values: true,
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
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 0,
              col: 1,
            },
            classification: { level: "Restricted", labels: [] },
          },
        ],
        audit_logger,
      },
    });

    const read = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:C1" },
    });

    expect(read.ok).toBe(true);
    expect(read.tool).toBe("read_range");
    if (!read.ok || read.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(read.data?.values).toEqual([[null, "[REDACTED]", 4]]);

    const stats = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:C1", measures: ["mean", "sum", "count"] },
    });

    expect(stats.ok).toBe(true);
    expect(stats.tool).toBe("compute_statistics");
    if (!stats.ok || stats.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    // A1 is a formula cell with a computed value (2), but the range DLP decision is REDACT due to B1.
    // Formula-derived values should not influence derived computations under REDACT.
    expect(stats.data?.statistics).toEqual({ mean: 4, sum: 4, count: 1 });

    expect(audit_logger.log).toHaveBeenCalled();
  });
});

