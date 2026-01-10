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
});
