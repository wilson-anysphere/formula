import { describe, expect, it, vi } from "vitest";

import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import {
  createPreviewApprovalHandler,
  getSpreadsheetToolDefinitions,
  isSpreadsheetMutationTool,
  SpreadsheetLLMToolExecutor,
  type PreviewApprovalRequest
} from "../src/llm/integration.js";

import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../security/dlp/src/selectors.js";

describe("llm integration helpers", () => {
  it("marks mutating tools as requiring approval when configured", () => {
    const defs = getSpreadsheetToolDefinitions({ require_approval_for_mutations: true });
    const read = defs.find((t) => t.name === "read_range");
    const write = defs.find((t) => t.name === "write_cell");

    expect(read?.requiresApproval).toBe(false);
    expect(write?.requiresApproval).toBe(true);
    expect(isSpreadsheetMutationTool("apply_formatting")).toBe(true);
    expect(isSpreadsheetMutationTool("compute_statistics")).toBe(false);
  });

  it("filters tool definitions based on ToolPolicy", () => {
    const defs = getSpreadsheetToolDefinitions({ toolPolicy: { allowCategories: ["read", "analysis"] } });
    const names = defs.map((t) => t.name);
    expect(names).toContain("read_range");
    expect(names).toContain("compute_statistics");
    expect(names).not.toContain("write_cell");
    expect(names).not.toContain("fetch_external_data");
  });

  it("always requires approval for fetch_external_data", () => {
    const defs = getSpreadsheetToolDefinitions();
    const fetchExternal = defs.find((t) => t.name === "fetch_external_data");
    expect(fetchExternal?.requiresApproval).toBe(true);
  });

  it("createPreviewApprovalHandler auto-approves safe changes and delegates risky ones", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });

    const onApprovalRequired = vi.fn(async (_request: PreviewApprovalRequest) => false);
    const handler = createPreviewApprovalHandler({
      spreadsheet: workbook,
      on_approval_required: onApprovalRequired
    });

    const safe = await handler({ name: "write_cell", arguments: { cell: "Sheet1!B1", value: 2 } });
    expect(safe).toBe(true);
    expect(onApprovalRequired).not.toHaveBeenCalled();

    const risky = await handler({ name: "write_cell", arguments: { cell: "Sheet1!A1", value: null } });
    expect(risky).toBe(false);
    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
    const request = onApprovalRequired.mock.calls[0]?.[0];
    expect(request).toBeDefined();
    expect(request!.preview.summary.deletes).toBe(1);
  });

  it("SpreadsheetLLMToolExecutor adapts LLM tool calls to ToolExecutor", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new SpreadsheetLLMToolExecutor(workbook);

    const result = await executor.execute({
      id: "call-1",
      name: "write_cell",
      arguments: { cell: "Sheet1!A1", value: 42 }
    });

    expect(result.ok).toBe(true);
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(42);
  });

  it("SpreadsheetLLMToolExecutor resolves display sheet names when sheet_name_resolver is provided", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      default_sheet: "Sheet2",
      sheet_name_resolver: sheetNameResolver
    });

    const result = await executor.execute({
      id: "call-1",
      name: "write_cell",
      arguments: { cell: "Budget!A1", value: 42 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    expect(result.data?.cell).toBe("Budget!A1");

    expect(workbook.getCell(parseA1Cell("Sheet2!A1")).value).toBe(42);
    // InMemoryWorkbook would create missing sheets; ensure we did not create "Budget".
    expect(workbook.listSheets()).toEqual(["Sheet2"]);
  });

  it("SpreadsheetLLMToolExecutor forwards include_formula_values to ToolExecutor", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });

    const executor = new SpreadsheetLLMToolExecutor(workbook, { include_formula_values: true });
    const result = await executor.execute({
      id: "call-1",
      name: "read_range",
      arguments: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[2]]);
  });

  it("SpreadsheetLLMToolExecutor does not surface formula values under DLP REDACT decisions", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "secret" });
    workbook.setCell(parseA1Cell("Sheet1!C1"), { value: 4 });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
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
      },
    } as any);

    const result = await executor.execute({
      id: "call-1",
      name: "read_range",
      arguments: { range: "Sheet1!A1:C1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[null, "[REDACTED]", 4]]);
    expect(result.data?.formulas).toEqual([["=1+1", "[REDACTED]", null]]);
  });

  it("SpreadsheetLLMToolExecutor surfaces formula values under DLP ALLOW decisions when enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { formula: "=1+1", value: 2 });

    const audit_logger = { log: vi.fn() };
    const executor = new SpreadsheetLLMToolExecutor(workbook, {
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
    } as any);

    const result = await executor.execute({
      id: "call-1",
      name: "read_range",
      arguments: { range: "Sheet1!A1:A1", include_formulas: true },
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

  it("does not expose fetch_external_data when host external fetch is disabled", () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new SpreadsheetLLMToolExecutor(workbook);
    expect(executor.tools.map((t) => t.name)).not.toContain("fetch_external_data");
  });

  it("does not expose fetch_external_data when host allowlist is missing (defense in depth)", () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new SpreadsheetLLMToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: [] });
    expect(executor.tools.map((t) => t.name)).not.toContain("fetch_external_data");
  });

  it("does not expose create_chart when SpreadsheetApi lacks chart support", () => {
    const workbook: any = new InMemoryWorkbook(["Sheet1"]);
    workbook.createChart = undefined;
    const executor = new SpreadsheetLLMToolExecutor(workbook);
    expect(executor.tools.map((t) => t.name)).not.toContain("create_chart");
  });

  it("denies disallowed tool calls without executing them", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const setCellSpy = vi.spyOn(workbook, "setCell");
    const executor = new SpreadsheetLLMToolExecutor(workbook, { toolPolicy: { allowCategories: ["read"] } });

    const result = await executor.execute({
      id: "call-1",
      name: "write_cell",
      arguments: { cell: "Sheet1!A1", value: 42 }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(setCellSpy).not.toHaveBeenCalled();
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBeNull();
  });

  it("SpreadsheetLLMToolExecutor enforces allowed_tools (disallowed tools cannot execute)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });

    const executor = new SpreadsheetLLMToolExecutor(workbook, { allowed_tools: ["read_range"] });
    expect(executor.tools.map((t) => t.name)).toEqual(["read_range"]);

    const denied = await executor.execute({
      id: "call-1",
      name: "write_cell",
      arguments: { cell: "Sheet1!A1", value: 42 }
    });

    expect(denied.ok).toBe(false);
    expect(denied.error?.code).toBe("permission_denied");
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(1);

    const unknown = await executor.execute({
      id: "call-2",
      name: "nonexistent_tool",
      arguments: {}
    });
    expect(unknown.ok).toBe(false);
    expect(unknown.error?.code).toBe("not_implemented");
  });
});
