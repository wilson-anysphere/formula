import { describe, expect, it, vi } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../security/dlp/src/selectors.js";

describe("ToolExecutor include_formula_values", () => {
  it("defaults to treating formula cells as null even when a computed value is present", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 4 });

    const executor = new ToolExecutor(workbook);

    const read = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:B1" },
    });
    expect(read.ok).toBe(true);
    expect(read.tool).toBe("read_range");
    if (!read.ok || read.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(read.data?.values).toEqual([[null, 4]]);

    const stats = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B1", measures: ["mean", "sum", "count"] },
    });
    expect(stats.ok).toBe(true);
    expect(stats.tool).toBe("compute_statistics");
    if (!stats.ok || stats.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(stats.data?.statistics).toEqual({ mean: 4, sum: 4, count: 1 });
  });

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

  it("read_range returns formulas and computed values together when include_formulas=true and include_formula_values is enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[2]]);
    expect(result.data?.formulas).toEqual([["=1+1"]]);
  });

  it("read_range surfaces computed values for formula cells under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });

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
        audit_logger,
      },
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[2]]);
    expect(result.data?.formulas).toEqual([["=1+1"]]);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
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

  it("compute_statistics includes numeric values from formula cells under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 4 });

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
        audit_logger,
      },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B1", measures: ["mean", "sum", "count"] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({ mean: 3, sum: 6, count: 2 });

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
    expect(event.redactedCellCount).toBe(0);
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

  it("filter_range compares formula cells using computed values under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=1+1", value: 2 });

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
        audit_logger,
      },
    });

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

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
    expect(event.redactedCellCount).toBe(0);
  });

  it("does not fall back to comparing formula text when include_formula_values is enabled but a formula cell has no computed value", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    // Simulate a backend that stores the formula but has not computed/filled `value` yet.
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=SECRET()", value: null });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A2",
        has_header: true,
        criteria: [{ column: "A", operator: "contains", value: "SECRET" }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");
    expect(result.data?.matching_rows).toEqual([]);
    expect(result.data?.count).toBe(0);
  });

  it("sort_range compares formula cells using computed values when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=10", value: 10 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { formula: "=2", value: 2 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "sort_range",
      parameters: { range: "Sheet1!A1:A3", has_header: true, sort_by: [{ column: "A", order: "asc" }] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("sort_range");
    if (!result.ok || result.tool !== "sort_range") throw new Error("Unexpected tool result");

    // If we sorted by formula text ("=10" < "=2"), the order would be unchanged.
    // With include_formula_values, the computed values (2 < 10) determine order.
    expect(workbook.getCell(parseA1Cell("Sheet1!A2")).formula).toBe("=2");
    expect(workbook.getCell(parseA1Cell("Sheet1!A2")).value).toBe(2);
    expect(workbook.getCell(parseA1Cell("Sheet1!A3")).formula).toBe("=10");
    expect(workbook.getCell(parseA1Cell("Sheet1!A3")).value).toBe(10);
  });

  it("sort_range compares formula cells using computed values under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=10", value: 10 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { formula: "=2", value: 2 });

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
        audit_logger,
      },
    });

    const result = await executor.execute({
      name: "sort_range",
      parameters: { range: "Sheet1!A1:A3", has_header: true, sort_by: [{ column: "A", order: "asc" }] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("sort_range");
    if (!result.ok || result.tool !== "sort_range") throw new Error("Unexpected tool result");

    expect(workbook.getCell(parseA1Cell("Sheet1!A2")).formula).toBe("=2");
    expect(workbook.getCell(parseA1Cell("Sheet1!A3")).formula).toBe("=10");

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
    expect(event.redactedCellCount).toBe(0);
  });

  it("detect_anomalies can include numeric values from formula cells when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A4"), { formula: "=100", value: 100 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A4", method: "zscore", threshold: 1.4 },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "zscore") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies.map((a) => a.cell)).toEqual(["Sheet1!A4"]);
    expect(result.data.anomalies[0]?.value).toBe(100);
  });

  it("detect_anomalies can include formula values under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A4"), { formula: "=100", value: 100 });

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
        audit_logger,
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

    expect(result.data.anomalies.map((a) => a.cell)).toEqual(["Sheet1!A4"]);
    expect(result.data.anomalies[0]?.value).toBe(100);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
    expect(event.redactedCellCount).toBe(0);
  });

  it("create_pivot_table can include numeric values from formula cells when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Category" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: "Unused" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "A" });
    // Simulate a backend that stores computed values as formatted strings.
    workbook.setCell(parseA1Cell("Sheet1!B2"), { formula: "=1+9", value: "10" });
    workbook.setCell(parseA1Cell("Sheet1!C2"), { value: 0 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: "A" });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: "20" });
    workbook.setCell(parseA1Cell("Sheet1!C3"), { value: 0 });

    const executor = new ToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C3",
        destination: "Sheet1!D1",
        rows: ["Category"],
        columns: [],
        values: [{ field: "Value", aggregation: "sum" }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");
    if (!result.ok || result.tool !== "create_pivot_table") throw new Error("Unexpected tool result");
    expect(result.data?.destination_range).toBe("Sheet1!D1:E3");

    expect(workbook.getCell(parseA1Cell("Sheet1!D2")).value).toBe("A");
    expect(workbook.getCell(parseA1Cell("Sheet1!E2")).value).toBe(30);
    expect(workbook.getCell(parseA1Cell("Sheet1!E3")).value).toBe(30);

    // Trigger pivot refresh by mutating an unused column inside the pivot source range.
    await executor.execute({ name: "write_cell", parameters: { cell: "Sheet1!C2", value: 1 } });
    expect(workbook.getCell(parseA1Cell("Sheet1!E2")).value).toBe(30);
    expect(workbook.getCell(parseA1Cell("Sheet1!E3")).value).toBe(30);
  });

  it("create_pivot_table can include formula values under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Category" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: "Unused" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "A" });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { formula: "=1+9", value: "10" });
    workbook.setCell(parseA1Cell("Sheet1!C2"), { value: 0 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: "A" });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: "20" });
    workbook.setCell(parseA1Cell("Sheet1!C3"), { value: 0 });

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
        audit_logger,
      },
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C3",
        destination: "Sheet1!D1",
        rows: ["Category"],
        columns: [],
        values: [{ field: "Value", aggregation: "sum" }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");
    if (!result.ok || result.tool !== "create_pivot_table") throw new Error("Unexpected tool result");
    expect(result.data?.destination_range).toBe("Sheet1!D1:E3");

    expect(workbook.getCell(parseA1Cell("Sheet1!E2")).value).toBe(30);
    expect(workbook.getCell(parseA1Cell("Sheet1!E3")).value).toBe(30);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event.decision?.decision).toBe("allow");
    expect(event.redactedCellCount).toBe(0);
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
      parameters: { range: "Sheet1!A1:C1", include_formulas: true },
    });

    expect(read.ok).toBe(true);
    expect(read.tool).toBe("read_range");
    if (!read.ok || read.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(read.data?.values).toEqual([[null, "[REDACTED]", 4]]);
    expect(read.data?.formulas).toEqual([["=1+1", "[REDACTED]", null]]);

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

    const anomalies = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:C1", method: "zscore", threshold: 0.6 },
    });
    expect(anomalies.ok).toBe(true);
    expect(anomalies.tool).toBe("detect_anomalies");
    if (!anomalies.ok || anomalies.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!anomalies.data || anomalies.data.method !== "zscore") throw new Error("Unexpected anomaly result");
    expect(anomalies.data.anomalies).toEqual([]);

    expect(audit_logger.log).toHaveBeenCalled();
  });

  it("does not use formula values for filter_range comparisons under DLP REDACT decisions", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Secret" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: "ok" });

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
              row: 1,
              col: 1,
            },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:B3",
        has_header: true,
        criteria: [{ column: "A", operator: "greater", value: 1 }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");

    // DLP decision for the range is REDACT due to B2; formula values should not be used as a signal.
    expect(result.data?.matching_rows).toEqual([3]);
    expect(result.data?.count).toBe(1);
  });

  it("does not use formula values for pivots under DLP REDACT decisions", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Category" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: "Secret" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "A" });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { formula: "=1+9", value: 10 });
    workbook.setCell(parseA1Cell("Sheet1!C2"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: "A" });
    workbook.setCell(parseA1Cell("Sheet1!B3"), { value: 20 });
    workbook.setCell(parseA1Cell("Sheet1!C3"), { value: "ok" });

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
              row: 1,
              col: 2,
            },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C3",
        destination: "Sheet1!D1",
        rows: ["Category"],
        columns: [],
        values: [{ field: "Value", aggregation: "sum" }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");
    if (!result.ok || result.tool !== "create_pivot_table") throw new Error("Unexpected tool result");

    // DLP decision for the range is REDACT due to C2; formula values should not influence the pivot.
    expect(workbook.getCell(parseA1Cell("Sheet1!E2")).value).toBe(20);
    expect(workbook.getCell(parseA1Cell("Sheet1!E3")).value).toBe(20);
  });
});
