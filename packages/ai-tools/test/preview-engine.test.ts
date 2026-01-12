import { afterEach, describe, expect, it, vi } from "vitest";
import { PreviewEngine } from "../src/preview/preview-engine.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import type { CreateChartResult, CreateChartSpec, SpreadsheetApi } from "../src/spreadsheet/api.js";

describe("PreviewEngine", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("flags large edits for approval and truncates change list", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine({ max_preview_changes: 20, approval_cell_threshold: 100 });

    const values = Array.from({ length: 10 }, (_, r) =>
      Array.from({ length: 15 }, (_, c) => `${r + 1}:${c + 1}`)
    );

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "set_range",
          parameters: {
            range: "Sheet1!A1:O10",
            values
          }
        }
      ],
      workbook
    );

    expect(preview.summary.total_changes).toBe(150);
    expect(preview.changes.length).toBe(20);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.startsWith("Large edit"))).toBe(true);
  });

  it("detects deletes and requires approval", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 123 });

    const previewEngine = new PreviewEngine();
    const preview = await previewEngine.generatePreview(
      [
        {
          name: "write_cell",
          parameters: { cell: "Sheet1!A1", value: null }
        }
      ],
      workbook
    );

    expect(preview.summary.deletes).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.includes("Deletes"))).toBe(true);
  });

  it("requires approval for fetch_external_data previews", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "fetch_external_data",
          parameters: {
            source_type: "api",
            url: "https://example.com/data",
            destination: "Sheet1!A1"
          }
        }
      ],
      workbook
    );

    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons).toContain("External data access requested");
  });

  it("never performs network access during preview, even when executor options enable it", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine();

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called during preview");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "fetch_external_data",
          parameters: {
            source_type: "api",
            url: "https://example.com/data",
            destination: "Sheet1!A1"
          }
        }
      ],
      workbook,
      { allow_external_data: true, allowed_external_hosts: ["example.com"] }
    );

    expect(fetchMock).not.toHaveBeenCalled();
    expect(preview.requires_approval).toBe(true);
    expect(preview.tool_results[0]?.ok).toBe(false);
    expect(preview.tool_results[0]?.error?.code).toBe("permission_denied");
  });

  it("supports create_chart previews when the SpreadsheetApi exposes createChart", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "create_chart",
          parameters: {
            chart_type: "bar",
            data_range: "Sheet1!A1:B3",
            title: "Sales"
          }
        }
      ],
      workbook
    );

    expect(preview.summary.total_changes).toBe(0);
    expect(preview.requires_approval).toBe(false);
    expect(preview.tool_results[0]?.ok).toBe(true);
  });

  it("resolves display sheet names via executorOptions.sheet_name_resolver", async () => {
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    class StrictWorkbook implements SpreadsheetApi {
      private readonly inner: InMemoryWorkbook;
      private readonly chartCalls: CreateChartSpec[];

      constructor(sheetNames: string[], chartCalls: CreateChartSpec[]) {
        this.inner = new InMemoryWorkbook(sheetNames);
        this.chartCalls = chartCalls;
      }

      private assertSheetExists(sheet: string): void {
        if (!this.inner.listSheets().includes(sheet)) {
          throw new Error(`Unknown sheet "${sheet}"`);
        }
      }

      clone(): SpreadsheetApi {
        // Keep the same chart call recorder across clones so PreviewEngine tests can
        // observe what the simulated workbook invoked.
        const next = new StrictWorkbook([], this.chartCalls);
        (next as any).inner = this.inner.clone();
        return next;
      }

      listSheets(): string[] {
        return this.inner.listSheets();
      }
      listNonEmptyCells(sheet?: string) {
        if (sheet) this.assertSheetExists(sheet);
        return this.inner.listNonEmptyCells(sheet);
      }
      getCell(address: any) {
        this.assertSheetExists(address.sheet);
        return this.inner.getCell(address);
      }
      setCell(address: any, cell: any) {
        this.assertSheetExists(address.sheet);
        return this.inner.setCell(address, cell);
      }
      readRange(range: any) {
        this.assertSheetExists(range.sheet);
        return this.inner.readRange(range);
      }
      writeRange(range: any, cells: any) {
        this.assertSheetExists(range.sheet);
        return this.inner.writeRange(range, cells);
      }
      applyFormatting(range: any, format: any) {
        this.assertSheetExists(range.sheet);
        return this.inner.applyFormatting(range, format);
      }
      getLastUsedRow(sheet: string) {
        this.assertSheetExists(sheet);
        return this.inner.getLastUsedRow(sheet);
      }
      createChart(spec: CreateChartSpec): CreateChartResult {
        this.chartCalls.push(spec);
        return { chart_id: `chart_${this.chartCalls.length}` };
      }
    }

    const chartCalls: CreateChartSpec[] = [];
    const workbook = new StrictWorkbook(["Sheet2"], chartCalls);
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "write_cell",
          parameters: { cell: "Budget!A1", value: 1 }
        }
      ],
      workbook,
      { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver }
    );

    expect(preview.tool_results[0]?.ok).toBe(true);
    expect(preview.tool_results[0]?.tool).toBe("write_cell");
    expect((preview.tool_results[0] as any)?.data?.cell).toBe("Budget!A1");

    // Preview does not mutate the live workbook.
    expect(workbook.listNonEmptyCells()).toEqual([]);
  });

  it("passes stable sheet ids to createChart during preview when using sheet_name_resolver", async () => {
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    class CapturingWorkbook extends InMemoryWorkbook {
      constructor(sheetNames: string[], private readonly calls: CreateChartSpec[]) {
        super(sheetNames);
      }

      override clone(): CapturingWorkbook {
        const next = new CapturingWorkbook([], this.calls);
        // Copy internal state via base clone, then copy its private fields onto the wrapper.
        const cloned = super.clone() as any;
        (next as any).sheets = cloned.sheets;
        (next as any).nextChartId = cloned.nextChartId;
        (next as any).charts = cloned.charts;
        return next;
      }

      override createChart(spec: CreateChartSpec): CreateChartResult {
        this.calls.push(spec);
        return { chart_id: `chart_${this.calls.length}` };
      }
    }

    const chartCalls: CreateChartSpec[] = [];
    const workbook = new CapturingWorkbook(["Sheet2"], chartCalls);
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "create_chart",
          parameters: { chart_type: "bar", data_range: "Budget!A1:B3", title: "Sales" }
        }
      ],
      workbook,
      { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver }
    );

    expect(preview.tool_results[0]?.ok).toBe(true);
    expect(preview.tool_results[0]?.tool).toBe("create_chart");
    expect((preview.tool_results[0] as any)?.data?.data_range).toBe("Budget!A1:B3");

    expect(chartCalls).toHaveLength(1);
    expect(chartCalls[0]?.data_range).toBe("Sheet2!A1:B3");
  });

  it("requires approval for large formatting edits even when the cell-level diff is empty (tool-reported counts)", async () => {
    // Some spreadsheet backends store formatting in layered defaults / range runs without
    // materializing per-cell entries. PreviewEngine should still require approval based on
    // tool-reported cell counts.
    class FormattingLayerWorkbook implements SpreadsheetApi {
      private readonly inner: InMemoryWorkbook;
      readonly formattingOps: Array<{ range: any; format: any }> = [];

      constructor(inner?: InMemoryWorkbook, ops?: Array<{ range: any; format: any }>) {
        this.inner = inner ?? new InMemoryWorkbook(["Sheet1"]);
        if (ops) this.formattingOps.push(...ops);
      }

      clone(): SpreadsheetApi {
        return new FormattingLayerWorkbook(this.inner.clone() as InMemoryWorkbook, this.formattingOps.slice());
      }

      listSheets(): string[] {
        return this.inner.listSheets();
      }

      listNonEmptyCells(sheet?: string) {
        return this.inner.listNonEmptyCells(sheet);
      }

      getCell(address: any) {
        return this.inner.getCell(address);
      }

      setCell(address: any, cell: any) {
        return this.inner.setCell(address, cell);
      }

      readRange(range: any) {
        return this.inner.readRange(range);
      }

      writeRange(range: any, cells: any) {
        return this.inner.writeRange(range, cells);
      }

      applyFormatting(range: any, format: any): number {
        // Record the op but do not materialize cells.
        this.formattingOps.push({ range, format });
        const rows = Math.max(0, range.endRow - range.startRow + 1);
        const cols = Math.max(0, range.endCol - range.startCol + 1);
        return rows * cols;
      }

      getLastUsedRow(sheet: string): number {
        return this.inner.getLastUsedRow(sheet);
      }
    }

    const workbook = new FormattingLayerWorkbook();
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "apply_formatting",
          parameters: { range: "Sheet1!A1:A1048576", format: { bold: true } }
        }
      ],
      workbook,
      // ToolExecutor enforces a default `max_tool_range_cells` guard (200k). Override it here so we
      // can validate PreviewEngine's approval gating based on tool-reported cell counts for
      // formatting backends that don't materialize per-cell entries.
      { max_tool_range_cells: 2_000_000 }
    );

    // Diff should miss the edit because formatting isn't stored as per-cell entries.
    expect(preview.summary.total_changes).toBe(0);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons).toContain("Large edit (1048576 cells)");
    expect(preview.warnings.some((w) => /diff may be incomplete/i.test(w))).toBe(true);

    // Ensure preview simulation didn't mutate the live workbook.
    expect(workbook.listNonEmptyCells()).toEqual([]);
    expect(workbook.formattingOps).toEqual([]);
  });
});
