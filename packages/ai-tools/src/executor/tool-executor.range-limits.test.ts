import { describe, expect, it } from "vitest";

import { ToolExecutor } from "./tool-executor.js";
import type { SpreadsheetApi } from "../spreadsheet/api.js";
import type { CellData } from "../spreadsheet/types.js";

class CountingSpreadsheet implements SpreadsheetApi {
  readCalls = 0;
  setCalls = 0;
  formatCalls = 0;

  listSheets(): string[] {
    return ["Sheet1"];
  }

  listNonEmptyCells(): any[] {
    return [];
  }

  getCell(): CellData {
    return { value: null };
  }

  setCell(): void {
    this.setCalls += 1;
  }

  readRange(): CellData[][] {
    this.readCalls += 1;
    throw new Error("readRange should not be called for oversized tool ranges");
  }

  writeRange(): void {
    // no-op
  }

  applyFormatting(): number {
    this.formatCalls += 1;
    return 0;
  }

  getLastUsedRow(): number {
    return 0;
  }

  clone(): SpreadsheetApi {
    return this;
  }
}

describe("ToolExecutor range size limits", () => {
  it("blocks sort_range before materializing a huge CellData[][]", async () => {
    const api = new CountingSpreadsheet();
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "sort_range",
      parameters: {
        // 26 * 8000 = 208k cells (exceeds the default max_tool_range_cells of 200k).
        range: "Sheet1!A1:Z8000",
        sort_by: [{ column: "A", order: "asc" }],
      },
    });

    expect(result.ok).toBe(false);
    expect(result.tool).toBe("sort_range");
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toContain("max_tool_range_cells");
    expect(api.readCalls).toBe(0);
  });

  it("does not block apply_formatting with max_tool_range_cells (formatting is host-guarded)", async () => {
    const api = new CountingSpreadsheet();
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "apply_formatting",
      parameters: {
        // 26 * 8000 = 208k cells (exceeds the default max_tool_range_cells of 200k).
        range: "Sheet1!A1:Z8000",
        format: { bold: true },
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("apply_formatting");
    if (!result.ok || result.tool !== "apply_formatting") {
      throw new Error(`Expected apply_formatting to succeed: ${result.error?.message ?? "unknown error"}`);
    }
    expect(result.data?.formatted_cells).toBe(0);
    expect(api.formatCalls).toBe(1);
    // Ensure we still don't materialize cell grids for formatting operations.
    expect(api.readCalls).toBe(0);
  });

  it("blocks apply_formula_column before writing an unbounded number of cells", async () => {
    const api = new CountingSpreadsheet();
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: {
        column: "A",
        formula_template: "=A{row}",
        start_row: 1,
        end_row: 300_000,
      },
    });

    expect(result.ok).toBe(false);
    expect(result.tool).toBe("apply_formula_column");
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toContain("max_tool_range_cells");
    expect(api.setCalls).toBe(0);
  });
});
