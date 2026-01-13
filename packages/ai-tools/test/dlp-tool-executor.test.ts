import { describe, expect, it, vi } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { parseA1Cell, parseA1Range } from "../src/spreadsheet/a1.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";

import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../security/dlp/src/selectors.js";

describe("ToolExecutor DLP enforcement", () => {
  it("read_range redacts restricted cells when policy allows redaction", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: 123 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 0,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:C1", include_formulas: true }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["ok", "[REDACTED]", 123]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]", null]]);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "read_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:C1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("read_range enforces table-based column selectors when a resolver is provided", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 100 });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: 3 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.COLUMN,
              documentId: "doc-1",
              sheetId: "Sheet1",
              tableId: "t1",
              columnId: "cB"
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        table_column_resolver: {
          getColumnIndex(sheetId: string, tableId: string, columnId: string) {
            if (sheetId === "Sheet1" && tableId === "t1" && columnId === "cB") return 1; // Column B
            return null;
          }
        },
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:C1", include_formulas: true }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([[1, "[REDACTED]", 3]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]", null]]);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "read_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:C1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("read_range does not leak object-valued restricted cells when policy allows redaction", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: { secret: "TopSecret" } as any });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
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

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:B1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["ok", "[REDACTED]"]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]"]]);

    const serialized = JSON.stringify(result);
    expect(serialized).not.toContain("TopSecret");
  });

  it("read_range redacts restricted cells even when the tool call uses a display sheet name", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    workbook.setCell(parseA1Cell("Sheet2!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet2!B1"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet2!C1"), { value: 123 });

    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      default_sheet: "Sheet2",
      sheet_name_resolver: sheetNameResolver,
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet2",
              row: 0,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Budget!A1:C1", include_formulas: true }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    // User-facing range uses display name.
    expect(result.data?.range).toBe("Budget!A1:C1");
    expect(result.data?.values).toEqual([["ok", "[REDACTED]", 123]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]", null]]);

    // Audit logging is stable-id based (sheetId).
    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "read_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet2!A1:C1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("canonicalizes dlp.sheet_id when it is provided as a display sheet name", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    workbook.setCell(parseA1Cell("Sheet2!A1"), { value: "ok" });
    workbook.setCell(parseA1Cell("Sheet2!B1"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet2!C1"), { value: 123 });

    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      // Note: both default_sheet and dlp.sheet_id are passed as the *display* name.
      // The executor should canonicalize them to the stable sheet id for internal use.
      default_sheet: "Budget",
      sheet_name_resolver: sheetNameResolver,
      dlp: {
        document_id: "doc-1",
        sheet_id: "Budget",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet2",
              row: 0,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "A1:C1", include_formulas: true }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.range).toBe("Budget!A1:C1");
    expect(result.data?.values).toEqual([["ok", "[REDACTED]", 123]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]", null]]);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "read_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet2!A1:C1",
      redactedCellCount: 1
    });
    expect(event.sheetId).toBe("Sheet2");
    expect(event.decision?.decision).toBe("redact");
  });

  it("read_range blocks when policy maxAllowed is null", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "data" });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: null,
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1" }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/DLP policy blocks reading Sheet1!A1/);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "read_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1",
      redactedCellCount: 0
    });
    expect(event.decision?.decision).toBe("block");
  });

  it("compute_statistics excludes restricted cells when redacting", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 100 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 0,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B1", measures: ["mean", "min", "max"] }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({ mean: 1, min: 1, max: 1 });

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "compute_statistics",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:B1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("compute_statistics excludes values from table-based restricted column selectors under REDACT policy", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 100 });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: 3 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.COLUMN,
              documentId: "doc-1",
              sheetId: "Sheet1",
              tableId: "t1",
              columnId: "cB"
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        table_column_resolver: {
          getColumnIndex(sheetId: string, tableId: string, columnId: string) {
            if (sheetId === "Sheet1" && tableId === "t1" && columnId === "cB") return 1; // Column B
            return null;
          }
        },
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:C1", measures: ["mean", "min", "max"] }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({ mean: 2, min: 1, max: 3 });

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "compute_statistics",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:C1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("compute_statistics correlation does not incorporate restricted pairs", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: 10 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: 20 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 1,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range: "Sheet1!A1:B2", measures: ["correlation"] }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

    // With the second pair redacted, correlation falls back to a single-point calculation (0).
    expect(result.data?.statistics.correlation).toBe(0);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "compute_statistics",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:B2",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("create_pivot_table excludes restricted source cells when redacting", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 1,
              col: 2
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:C5",
        values: [
          ["Region", "Product", "Sales"],
          ["East", "A", 100],
          ["East", "B", 150],
          ["West", "A", 200],
          ["West", "B", 250]
        ]
      }
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C5",
        rows: ["Region"],
        columns: ["Product"],
        values: [{ field: "Sales", aggregation: "sum" }],
        destination: "Sheet1!E1"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");

    const out = workbook
      .readRange(parseA1Range("Sheet1!E1:H4"))
      .map((row) => row.map((cell) => cell.value));

    expect(out).toEqual([
      ["Region", "A - Sum of Sales", "B - Sum of Sales", "Grand Total - Sum of Sales"],
      ["East", null, 150, 150],
      ["West", 200, 250, 450],
      ["Grand Total", 200, 400, 600]
    ]);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "create_pivot_table",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:C5",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("sort_range is denied when DLP requires redaction for the range", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Name" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Salary" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "Alice" });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: 100 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 1,
              col: 1
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "sort_range",
      parameters: { range: "Sheet1!A1:B2", sort_by: [{ column: "B", order: "desc" }], has_header: true }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/DLP policy blocks sorting Sheet1!A1:B2/);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "sort_range",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:B2",
      redactedCellCount: 0
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("write_cell does not reveal equality with restricted content via changed=false", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 42 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 0,
              col: 0
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 42 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    // If we returned the true `changed` value, this would be `false` and would leak that the old value was 42.
    expect(result.data?.changed).toBe(true);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "write_cell",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("apply_formula_column is denied when policy blocks cloud processing", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Header" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 1 });

    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: null,
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        audit_logger
      }
    });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: { column: "B", formula_template: "=A{row}*2", start_row: 2, end_row: -1 }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/DLP policy blocks applying formulas to Sheet1!B2/);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "apply_formula_column",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!B2",
      redactedCellCount: 0
    });
    expect(event.decision?.decision).toBe("block");
  });

  it("detect_anomalies does not let restricted values influence anomaly scores under REDACT policy", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const audit_logger = { log: vi.fn() };
    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 4,
              col: 0
            },
            classification: { level: "Restricted", labels: [] }
          }
        ],
        audit_logger
      }
    });

    // A5 is restricted. If its value (200) influenced z-score computation, A4 (100)
    // would no longer be considered an anomaly at this threshold.
    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A5",
        values: [[1], [1], [1], [100], [200]]
      }
    });

    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A5", method: "zscore", threshold: 1.4 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "zscore") throw new Error("Unexpected anomaly result");

    const flagged = result.data.anomalies.map((a) => a.cell).sort();
    expect(flagged).toEqual(["Sheet1!A4"]);
    expect(result.data.anomalies[0]?.value).toBe(100);
    expect(result.data.anomalies[0]?.score).toBeGreaterThanOrEqual(1.4);

    expect(audit_logger.log).toHaveBeenCalledTimes(1);
    const event = audit_logger.log.mock.calls[0]?.[0];
    expect(event).toMatchObject({
      type: "ai.tool.dlp",
      tool: "detect_anomalies",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: "Sheet1!A1:A5",
      redactedCellCount: 1
    });
    expect(event.decision?.decision).toBe("redact");
  });

  it("filter_range does not treat disallowed criterion cells as matching the DLP redaction placeholder under REDACT policy", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Value" });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: "Secret" });
    // Literal placeholder-like values should still be filterable when the cell is allowed.
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: "[REDACTED]" });
    workbook.setCell(parseA1Cell("Sheet1!A4"), { value: "FRED" });

    const executor = new ToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true
            }
          }
        },
        classification_records: [
          {
            selector: {
              scope: CLASSIFICATION_SCOPE.CELL,
              documentId: "doc-1",
              sheetId: "Sheet1",
              row: 1,
              col: 0
            },
            classification: { level: "Restricted", labels: [] }
          }
        ]
      }
    });

    // If we evaluated disallowed cells against a placeholder, row 2 would incorrectly match.
    const equalsResult = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A4",
        has_header: true,
        criteria: [{ column: "A", operator: "equals", value: "[REDACTED]" }]
      }
    });
    expect(equalsResult.ok).toBe(true);
    expect(equalsResult.tool).toBe("filter_range");
    if (!equalsResult.ok || equalsResult.tool !== "filter_range") throw new Error("Unexpected tool result");
    expect(equalsResult.data?.matching_rows).toEqual([3]);
    expect(equalsResult.data?.count).toBe(1);

    const containsResult = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A4",
        has_header: true,
        criteria: [{ column: "A", operator: "contains", value: "RED" }]
      }
    });
    expect(containsResult.ok).toBe(true);
    expect(containsResult.tool).toBe("filter_range");
    if (!containsResult.ok || containsResult.tool !== "filter_range") throw new Error("Unexpected tool result");
    expect(containsResult.data?.matching_rows).toEqual([3, 4]);
    expect(containsResult.data?.count).toBe(2);
  });
});
