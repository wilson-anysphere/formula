import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { ToolExecutor, PreviewEngine } from "../../../../../packages/ai-tools/src/index.js";
import { workbookFromSpreadsheetApi } from "../../../../../packages/ai-rag/src/index.js";
import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_SCOPE } from "../../../../../packages/security/dlp/src/selectors.js";

import { DocumentControllerSpreadsheetApi } from "./documentControllerSpreadsheetApi.js";
import { createSheetNameResolverFromIdToNameMap } from "../../sheet/sheetNameResolver.js";
import { getLocale, setLocale } from "../../i18n/index.js";

describe("DocumentControllerSpreadsheetApi", () => {
  it("resolves display names to stable sheet ids (no phantom sheet creation after rename)", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet2", "A1", 1);

    const sheetNames = new Map<string, string>([["Sheet2", "Budget"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetNames);

    const api = new DocumentControllerSpreadsheetApi(controller, { sheetNameResolver });
    expect(api.listSheets()).toEqual(["Budget"]);

    const executor = new ToolExecutor(api, { default_sheet: "Sheet2" });
    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Budget!A1", value: 99 }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet2", "A1").value).toBe(99);
    expect(controller.getSheetIds()).toContain("Sheet2");
    expect(controller.getSheetIds()).not.toContain("Budget");

    expect(() => api.getCell({ sheet: "DoesNotExist", row: 1, col: 1 })).toThrow(/Unknown sheet/i);
    expect(controller.getSheetIds()).not.toContain("DoesNotExist");
  });

  it("does not resurrect deleted sheets when sheetNameResolver mappings are stale", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);
    controller.setCellValue("Sheet2", "A1", 2);

    const sheetNames = new Map<string, string>([["Sheet2", "Budget"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetNames);

    const api = new DocumentControllerSpreadsheetApi(controller, { sheetNameResolver });

    // Delete the sheet but keep the resolver mapping around (simulating a stale UI cache).
    controller.deleteSheet("Sheet2");
    expect(controller.getSheetIds()).toEqual(["Sheet1"]);

    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });
    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Budget!A1", value: 99 },
    });

    expect(result.ok).toBe(false);
    // Ensure the tool call did not recreate the deleted sheet.
    expect(controller.getSheetIds()).toEqual(["Sheet1"]);
    expect(controller.getSheetIds()).not.toContain("Sheet2");
    expect(controller.getSheetIds()).not.toContain("Budget");

    expect(() => api.getCell({ sheet: "Budget", row: 1, col: 1 })).toThrow(/Unknown sheet/i);
    expect(controller.getSheetIds()).toEqual(["Sheet1"]);
  });

  it("allows ToolExecutor to apply updates through DocumentController", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "A1", value: 99 }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(99);
  });

  it("does not create a history entry when setCell is a no-op", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);
    const before = controller.getStackDepths().undo;

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.setCell({ sheet: "Sheet1", row: 1, col: 1 }, { value: 1 });

    expect(controller.getStackDepths().undo).toBe(before);
  });

  it("does not create a history entry when setCell is a no-op for formulas (even when computed values are available)", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "B1", 1);
    controller.setCellFormula("Sheet1", "A1", "B1+1");
    const before = controller.getStackDepths().undo;

    const api = new DocumentControllerSpreadsheetApi(controller, {
      getCellComputedValueForSheet: (sheetId, cell) => {
        if (sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) return 2;
        return null;
      },
    });
    api.setCell({ sheet: "Sheet1", row: 1, col: 1 }, { value: null, formula: "=B1+1" });

    expect(controller.getStackDepths().undo).toBe(before);
  });

  it("roundtrips supported formatting between ai-tools CellFormat and DocumentController styles", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const beforeStyleTableSize = controller.styleTable.size;

    const result = await executor.execute({
      name: "apply_formatting",
      parameters: {
        range: "A1",
        format: {
          bold: true,
          italic: true,
          font_size: 14,
          font_color: "#FF00FF00",
          background_color: "#FFFFFF00",
          number_format: "$#,##0.00",
          horizontal_align: "center"
        }
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("apply_formatting");
    if (!result.ok || result.tool !== "apply_formatting") throw new Error("Unexpected tool result");
    expect(result.data?.formatted_cells).toBe(1);

    const cellState = controller.getCell("Sheet1", "A1");
    expect(cellState.styleId).toBeGreaterThan(0);
    expect(controller.styleTable.size).toBeGreaterThan(beforeStyleTableSize);

    const style = controller.styleTable.get(cellState.styleId);
    expect(style.font?.bold).toBe(true);
    expect(style.font?.italic).toBe(true);
    expect(style.font?.size).toBe(14);
    expect(style.font?.color).toBe("#FF00FF00");
    expect(style.fill?.pattern).toBe("solid");
    expect(style.fill?.fgColor).toBe("#FFFFFF00");
    expect(style.numberFormat).toBe("$#,##0.00");
    expect(style.alignment?.horizontal).toBe("center");

    const roundTrip = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
    expect(roundTrip.format).toEqual({
      bold: true,
      italic: true,
      font_size: 14,
      font_color: "#FF00FF00",
      background_color: "#FFFFFF00",
      number_format: "$#,##0.00",
      horizontal_align: "center"
    });
  });

  it("merges incremental apply_formatting patches like ai-tools (does not clobber prior fields)", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    await executor.execute({
      name: "apply_formatting",
      parameters: { range: "A1", format: { bold: true } }
    });

    await executor.execute({
      name: "apply_formatting",
      parameters: { range: "A1", format: { italic: true } }
    });

    const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
    expect(cell.format).toEqual({ bold: true, italic: true });
  });

  it("returns an error when DocumentController refuses to apply formatting (safety caps)", async () => {
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const controller = new DocumentController();
      const api = new DocumentControllerSpreadsheetApi(controller);
      const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

      // Full-width formatting over >50k rows is rejected by DocumentController's safety cap.
      const result = await executor.execute({
        name: "apply_formatting",
        parameters: { range: "A1:XFD60000", format: { bold: true } },
      });

      expect(result.ok).toBe(false);
      expect(result.tool).toBe("apply_formatting");
      if (result.ok) throw new Error("Expected apply_formatting to fail");
      expect(result.error?.message ?? "").toMatch(/Formatting could not be applied/i);
    } finally {
      warnSpy.mockRestore();
    }
  });

  it("uses the display sheet name in apply_formatting error messages when a sheetNameResolver is provided", async () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const controller = new DocumentController();
      controller.setCellValue("Sheet2", "A1", 1);

      const sheetNames = new Map<string, string>([["Sheet2", "Budget"]]);
      const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetNames);
      const api = new DocumentControllerSpreadsheetApi(controller, { sheetNameResolver });
      const executor = new ToolExecutor(api, { default_sheet: "Sheet2" });

      const result = await executor.execute({
        name: "apply_formatting",
        // Full-width formatting over >50k rows is rejected by DocumentController's safety cap.
        parameters: { range: "A1:XFD60000", format: { bold: true } },
      });

      expect(result.ok).toBe(false);
      expect(result.tool).toBe("apply_formatting");
      if (result.ok) throw new Error("Expected apply_formatting to fail");
      expect(result.error?.message ?? "").toContain("Budget!A1:XFD60000");
      expect(result.error?.message ?? "").not.toContain("Sheet2!A1:XFD60000");
    } finally {
      warnSpy.mockRestore();
    }
  });

  it("preserves existing formatting when write_cell updates values", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "A1", value: 2 }
    });

    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
  });

  it("filters out cells that are empty per ai-tools semantics (no value, formula, or supported format)", () => {
    const controller = new DocumentController();
    controller.setRangeFormat(
      "Sheet1",
      "A1",
      { border: { left: { style: "thin", color: "#FF000000" } } },
      { label: "Border only" }
    );

    const state = controller.getCell("Sheet1", "A1");
    expect(state.value).toBeNull();
    expect(state.formula).toBeNull();
    expect(state.styleId).toBeGreaterThan(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.listNonEmptyCells("Sheet1")).toEqual([]);
    expect(api.getLastUsedRow("Sheet1")).toBe(0);
  });

  it("counts supported formatting-only cells for getLastUsedRow()", () => {
    const controller = new DocumentController();
    controller.setRangeFormat("Sheet1", "A5", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.getLastUsedRow("Sheet1")).toBe(5);
  });

  it("does not include formatting-only cells in listNonEmptyCells()", () => {
    const controller = new DocumentController();
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.listNonEmptyCells("Sheet1")).toEqual([]);
  });

  it("does not call exportSheetForSemanticDiff for listNonEmptyCells() or getLastUsedRow()", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const spy = vi.spyOn(controller, "exportSheetForSemanticDiff");

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.listNonEmptyCells();
    api.getLastUsedRow("Sheet1");

    expect(spy).not.toHaveBeenCalled();
  });

  it("ignores legacy flat CellFormat keys stored in DocumentController styles for listNonEmptyCells()", () => {
    const controller = new DocumentController();
    // Simulate the pre-fix adapter behavior that wrote ai-tools CellFormat objects directly
    // into the DocumentController style table (flat keys like `bold`).
    controller.setRangeFormat("Sheet1", "A1", { bold: true, background_color: "#FFFFFF00" }, { label: "Legacy bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.listNonEmptyCells("Sheet1")).toEqual([]);
  });

  it("returns effective (layered) formatting for getCell() even when cellState.styleId is 0", () => {
    const controller = new DocumentController();
    controller.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } }, { label: "Bold column A" });

    const cellState = controller.getCell("Sheet1", "A1");
    expect(cellState.styleId).toBe(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
  });

  it("readRange() includes inherited formatting for empty cells", () => {
    const controller = new DocumentController();
    controller.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } }, { label: "Bold column A" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const cells = api.readRange({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 2, endCol: 1 });
    expect(cells).toEqual([
      [{ value: null, format: { bold: true } }],
      [{ value: null, format: { bold: true } }]
    ]);
  });

  it("readRange() includes range-run formatting for empty cells (large rectangle formatting layer)", () => {
    const controller = new DocumentController();
    // This range is large enough to trigger the compressed range-run formatting layer in DocumentController.
    controller.setRangeFormat("Sheet1", "A1:C20000", { font: { bold: true } }, { label: "Bold large rectangle" });

    const cellState = controller.getCell("Sheet1", "A1");
    expect(cellState.styleId).toBe(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const cells = api.readRange({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 2, endCol: 2 });
    expect(cells).toEqual([
      [
        { value: null, format: { bold: true } },
        { value: null, format: { bold: true } }
      ],
      [
        { value: null, format: { bold: true } },
        { value: null, format: { bold: true } }
      ]
    ]);
  });

  it("does not leak mutable references to DocumentController cell values from listNonEmptyCells()", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", {
      text: "Rich Bold",
      runs: [{ start: 0, end: 4, style: { bold: true } }]
    });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const entries = api.listNonEmptyCells("Sheet1");
    const value = entries[0]?.cell.value as any;
    expect(value?.text).toBe("Rich Bold");

    value.text = "Mutated";

    const after = controller.getCell("Sheet1", "A1").value as any;
    expect(after?.text).toBe("Rich Bold");
  });

  it("returns values & formulas compatible with workbookFromSpreadsheetApi (and clones object values)", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", "Hello");
    controller.setCellFormula("Sheet1", "B2", "SUM(A1:A1)");
    controller.setCellValue("Sheet1", "C3", {
      text: "Rich Bold",
      runs: [{ start: 0, end: 4, style: { bold: true } }]
    });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const workbook = workbookFromSpreadsheetApi({ spreadsheet: api as any, workbookId: "wb-doc" });
    const sheet = workbook.sheets.find((s: any) => s.name === "Sheet1");
    expect(sheet).toBeTruthy();
    if (!sheet) throw new Error("Missing sheet");

    expect(sheet.cells.get("0,0")).toEqual({ value: "Hello", formula: null });
    expect(sheet.cells.get("1,1")).toEqual({ value: null, formula: "=SUM(A1:A1)" });

    const c3 = sheet.cells.get("2,2");
    expect(c3?.value).toMatchObject({ text: "Rich Bold" });
    // Ensure mutating the returned workbook does not mutate the document.
    (c3 as any).value.text = "Mutated";
    expect((controller.getCell("Sheet1", "C3").value as any)?.text).toBe("Rich Bold");
  });

  it("does not call exportSheetForSemanticDiff from listNonEmptyCells()", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);
    const spy = vi.spyOn(controller, "exportSheetForSemanticDiff");

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.listNonEmptyCells("Sheet1");

    expect(spy).not.toHaveBeenCalled();
  });

  it("applies per-cell formatting provided to writeRange()", () => {
    const controller = new DocumentController();
    const api = new DocumentControllerSpreadsheetApi(controller);

    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 2 },
      [[{ value: 1, format: { bold: true } }, { value: 2 }]]
    );

    const a1 = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
    expect(a1.format).toEqual({ bold: true });

    const state = controller.getCell("Sheet1", "A1");
    expect(controller.styleTable.get(state.styleId).font?.bold).toBe(true);
    expect(controller.getCell("Sheet1", "B1").styleId).toBe(0);
  });

  it("does not copy inherited (layered) non-ai-tools styles into per-cell styles when writeRange applies formatting", () => {
    const controller = new DocumentController();
    const api = new DocumentControllerSpreadsheetApi(controller);

    // Apply a column default border via layered formatting (no per-cell materialization).
    controller.setRangeFormat("Sheet1", "A1:A1048576", {
      border: { left: { style: "thin", color: "#FF000000" } }
    });

    expect(controller.getCell("Sheet1", "A1").styleId).toBe(0);

    // Trigger the writeRange `hasAnyFormat` path by writing a supported ai-tools format.
    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 1 },
      [[{ value: null, format: { italic: true } }]]
    );

    const after = controller.getCell("Sheet1", "A1");
    expect(after.styleId).toBeGreaterThan(0);

    // The per-cell style should only contain ai-tools supported overrides (no border materialization).
    const cellStyle = controller.styleTable.get(after.styleId);
    expect(cellStyle.border).toBeUndefined();

    // The effective style should still include the inherited column border.
    const effectiveBeforeClear = controller.getCellFormat("Sheet1", "A1");
    expect(effectiveBeforeClear.border?.left?.style).toBe("thin");
    expect(effectiveBeforeClear.border?.left?.color).toBe("#FF000000");

    // Clearing the column formatting should remove the border from the effective style.
    controller.setRangeFormat("Sheet1", "A1:A1048576", null);
    const effectiveAfterClear = controller.getCellFormat("Sheet1", "A1");
    expect(effectiveAfterClear.border).toBeUndefined();
  });

  it("preserves inherited column formatting when writeRange writes partial per-cell formats (layered formats)", () => {
    const controller = new DocumentController();
    // Apply full-column formatting via the layered column style layer.
    controller.setColFormat("Sheet1", 0, {
      font: { bold: true },
      fill: { pattern: "solid", fgColor: "#FFFFFF00" }
    });

    const sheetModel = controller.model.sheets.get("Sheet1");
    expect(sheetModel).toBeDefined();
    // Column-level formatting should not materialize per-cell styleIds.
    expect(sheetModel!.cells.size).toBe(0);
    expect(controller.getCell("Sheet1", "A1").styleId).toBe(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 2, endCol: 1 },
      [[{ value: 1, format: { italic: true } }], [{ value: null }]]
    );

    // Writing A1's style should not materialize formatting into the rest of the column.
    expect(sheetModel!.cells.size).toBe(1);
    expect(controller.getCell("Sheet1", "A2").styleId).toBe(0);

    const effectiveA1 = controller.getCellFormat("Sheet1", "A1");
    expect(effectiveA1.font?.bold).toBe(true);
    expect(effectiveA1.font?.italic).toBe(true);
    expect(effectiveA1.fill?.fgColor).toBe("#FFFFFF00");
  });

  it("does not materialize inherited range-run formatting into per-cell styles when writeRange includes matching formats", () => {
    const controller = new DocumentController();
    // Force range-run formatting (compressed per-column runs) by formatting a large rectangle.
    controller.setRangeFormat("Sheet1", "A1:B25001", { font: { bold: true } }, { label: "Bold range" });

    const sheetModel = controller.model.sheets.get("Sheet1");
    expect(sheetModel).toBeDefined();
    expect(sheetModel!.formatRunsByCol?.size).toBeGreaterThan(0);
    // Range-run formatting should not eagerly materialize per-cell styleIds.
    expect(sheetModel!.cells.size).toBe(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    // Writing a range that includes the *effective* format (as returned by readRange/sort_range)
    // should not force that inherited formatting into per-cell styles.
    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 2, endCol: 1 },
      [[{ value: null, format: { bold: true } }], [{ value: null, format: { bold: true } }]]
    );

    expect(sheetModel!.cells.size).toBe(0);
  });

  it("does not drop existing per-cell formatting overrides that match inherited formatting when writeRange round-trips formats", () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [[1]]);
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold A1" });

    const beforeCell = controller.getCell("Sheet1", "A1");
    expect(beforeCell.styleId).toBeGreaterThan(0);

    // Apply column-level bold via layered formatting (colStyleIds).
    controller.setColFormat("Sheet1", 0, { font: { bold: true } });
    expect(controller.getCellFormat("Sheet1", "A1").font?.bold).toBe(true);

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.writeRange({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 1 }, [[{ value: 1, format: { bold: true } }]]);

    // writeRange should not clear A1's existing per-cell bold override just because it is also
    // satisfied by the inherited column formatting.
    const afterCell = controller.getCell("Sheet1", "A1");
    expect(afterCell.styleId).toBeGreaterThan(0);

    // Clearing the inherited column formatting should leave the per-cell formatting intact.
    controller.setColFormat("Sheet1", 0, null);
    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
  });

  it("clears stale formatting when writeRange moves formatted cells (no contamination)", () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1:B1", [[1, 2]]);
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold A1" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    // Move bold formatting from A1 -> B1 by writing a new matrix where only B1 has format.
    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 2 },
      [[{ value: 1 }, { value: 2, format: { bold: true } }]]
    );

    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toBeUndefined();
    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 2 }).format).toEqual({ bold: true });
  });

  it("does not merge stale supported formatting when writeRange overwrites a cell with a different per-cell format", () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1:A2", [[1], [2]]);
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold A1" });
    controller.setRangeFormat("Sheet1", "A2", { font: { italic: true } }, { label: "Italic A2" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    api.writeRange(
      { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 2, endCol: 1 },
      [[{ value: 1, format: { italic: true } }], [{ value: 2, format: { bold: true } }]]
    );

    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ italic: true });
    expect(api.getCell({ sheet: "Sheet1", row: 2, col: 1 }).format).toEqual({ bold: true });
  });

  it("preserves formatting when set_range updates values", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    await executor.execute({
      name: "set_range",
      parameters: { range: "A1", values: [[5]] }
    });

    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
  });

  it("sort_range moves formatting with values without duplicating styles", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [[3], [1], [2]]);
    controller.setRangeFormat("Sheet1", "A2", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "sort_range",
      parameters: {
        range: "A1:A3",
        sort_by: [{ column: "A", order: "asc" }],
        has_header: false
      }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(1);
    expect(controller.getCell("Sheet1", "A2").value).toBe(2);
    expect(controller.getCell("Sheet1", "A3").value).toBe(3);

    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
    expect(api.getCell({ sheet: "Sheet1", row: 2, col: 1 }).format).toBeUndefined();
    expect(api.getCell({ sheet: "Sheet1", row: 3, col: 1 }).format).toBeUndefined();
  });

  it("sort_range does not bake inherited column formatting into per-cell styles", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [[3], [1], [2]]);
    controller.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } }, { label: "Bold column A" });

    // Sanity: cell-level style ids remain default; formatting is inherited from the column layer.
    expect(controller.getCell("Sheet1", "A1").styleId).toBe(0);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "sort_range",
      parameters: {
        range: "A1:A3",
        sort_by: [{ column: "A", order: "asc" }],
        has_header: false
      }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(1);
    expect(controller.getCell("Sheet1", "A2").value).toBe(2);
    expect(controller.getCell("Sheet1", "A3").value).toBe(3);

    // Formatting should remain effective but should stay in the column formatting layer (styleId stays 0).
    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 }).format).toEqual({ bold: true });
    expect(controller.getCell("Sheet1", "A1").styleId).toBe(0);
    expect(controller.getCell("Sheet1", "A2").styleId).toBe(0);
    expect(controller.getCell("Sheet1", "A3").styleId).toBe(0);
  });

  it("validates writeRange dimensions like the ai-tools InMemoryWorkbook", () => {
    const controller = new DocumentController();
    const api = new DocumentControllerSpreadsheetApi(controller);

    expect(() =>
      api.writeRange({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 2 }, [[{ value: 1 }]])
    ).toThrow(/expected 2 columns/i);
  });

  it("supports PreviewEngine diffing without mutating the live controller", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 10);
    controller.setCellValue("Sheet1", "B1", { text: "Rich Bold", runs: [{ start: 0, end: 4, style: { bold: true } }] });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const previewEngine = new PreviewEngine({ approval_cell_threshold: 0 });

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "write_cell",
          parameters: { cell: "A1", value: 20 }
        }
      ],
      api,
      { default_sheet: "Sheet1", max_tool_range_cells: 2_000_000 }
    );

    expect(preview.summary.total_changes).toBe(1);
    expect(preview.summary.modifies).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(10);

    // Sanity check: tool call references normalize to Sheet1.
    expect(preview.changes[0]?.cell).toBe("Sheet1!A1");
  });

  it("does not include formatting-only changes in PreviewEngine diffs", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 10);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const previewEngine = new PreviewEngine({ approval_cell_threshold: 0 });

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "apply_formatting",
          parameters: { range: "A1", format: { bold: true } }
        }
      ],
      api,
      { default_sheet: "Sheet1", max_tool_range_cells: 2_000_000 }
    );

    // listNonEmptyCells is optimized for value/formula indexing and omits formatting-only
    // cells, so PreviewEngine diffs will not surface pure formatting changes when used
    // with this adapter.
    expect(preview.summary.total_changes).toBe(0);
    // Approval is still based on tool-reported edit sizes, even if the cell-level diff is empty.
    expect(preview.requires_approval).toBe(true);

    // Ensure the preview simulation didn't mutate the live controller.
    expect(controller.getCell("Sheet1", "A1").styleId).toBe(0);

    const previewEngineNoApproval = new PreviewEngine({ approval_cell_threshold: 1 });
    const previewNoApproval = await previewEngineNoApproval.generatePreview(
      [
        {
          name: "apply_formatting",
          parameters: { range: "A1", format: { bold: true } }
        }
      ],
      api,
      { default_sheet: "Sheet1" }
    );
    expect(previewNoApproval.requires_approval).toBe(false);
  });

  it("requires approval for large formatting edits even when PreviewEngine diffs cannot materialize cells (layered formats)", async () => {
    const controller = new DocumentController();
    // Ensure Sheet1 exists without placing any cells in column A (the formatted column).
    controller.setCellValue("Sheet1", "B1", 123);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const previewEngine = new PreviewEngine();

    const sheetModel = controller.model.sheets.get("Sheet1");
    const beforeCellCount = sheetModel?.cells?.size ?? 0;

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "apply_formatting",
          parameters: { range: "A1:A1048576", format: { bold: true } }
        }
      ],
      api,
      // Allow full-column formatting for this test so we can validate PreviewEngine approval
      // gating even when the underlying spreadsheet stores formatting in layered defaults
      // (no per-cell diffs).
      { default_sheet: "Sheet1", max_tool_range_cells: 2_000_000 }
    );

    // The DocumentController formatting layer stores full-column formatting without creating per-cell entries,
    // so PreviewEngine's cell-level diff should miss the edit.
    expect(preview.summary.total_changes).toBe(0);

    // But we should still require approval based on tool-reported cell counts.
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.startsWith("Large edit"))).toBe(true);

    expect(preview.warnings.some((warning) => /diff may be incomplete/i.test(warning))).toBe(true);

    // Ensure the preview simulation didn't mutate the live controller.
    expect(controller.model.sheets.get("Sheet1")?.cells?.size ?? 0).toBe(beforeCellCount);
    expect(controller.getCellFormat("Sheet1", "A1").font?.bold).toBeUndefined();
    expect(controller.getCell("Sheet1", "B1").value).toBe(123);
  });

  it("requires approval for large formatting edits even when PreviewEngine diffs cannot materialize cells (range runs)", async () => {
    const controller = new DocumentController();
    // Ensure Sheet1 exists without placing any cells inside the formatted rectangle.
    controller.setCellValue("Sheet1", "D1", 123);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const previewEngine = new PreviewEngine();

    const sheetModel = controller.model.sheets.get("Sheet1");
    const beforeCellCount = sheetModel?.cells?.size ?? 0;

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "apply_formatting",
          // This range is large enough to be stored in DocumentController's compressed range-run formatting layer,
          // which does not materialize per-cell entries.
          parameters: { range: "A1:C20000", format: { bold: true } }
        }
      ],
      api,
      { default_sheet: "Sheet1" }
    );

    expect(preview.summary.total_changes).toBe(0);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.startsWith("Large edit"))).toBe(true);
    expect(preview.warnings.some((warning) => /diff may be incomplete/i.test(warning))).toBe(true);

    // Ensure the preview simulation didn't mutate the live controller.
    expect(controller.model.sheets.get("Sheet1")?.cells?.size ?? 0).toBe(beforeCellCount);
    expect(controller.getCellFormat("Sheet1", "A1").font?.bold).toBeUndefined();
    expect(controller.getCell("Sheet1", "D1").value).toBe(123);
  });

  it("returns an error when DocumentController refuses apply_formatting edits", async () => {
    const controller = new DocumentController();
    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const spy = vi.spyOn(controller, "setRangeFormat").mockReturnValue(false);
    try {
      const result = await executor.execute({
        name: "apply_formatting",
        parameters: { range: "A1", format: { bold: true } }
      });

      expect(result.ok).toBe(false);
      expect(result.tool).toBe("apply_formatting");
      if (result.ok) throw new Error("Expected apply_formatting to fail");
      expect(result.error?.message ?? "").toMatch(/Formatting could not be applied/i);
      expect(spy).toHaveBeenCalled();
    } finally {
      spy.mockRestore();
    }
  });

  it("normalizes formulas to include leading '=' when reading through the adapter", async () => {
    const controller = new DocumentController();
    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "A1", value: "SUM(B1:B3)", is_formula: true }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").formula).toBe("=SUM(B1:B3)");

    const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
    expect(cell.formula).toBe("=SUM(B1:B3)");
  });

  it("preserves cached/computed values on formula cells (value + formula together)", async () => {
    const controller = new DocumentController();
    // Simulate a workbook import that includes both the formula text and a cached/computed result.
    // DocumentController's user-edit APIs set `value:null` for formulas, but snapshot import paths
    // can populate both.
    (controller as any).model.setCell("Sheet1", 0, 0, { value: 2, formula: "=1+1", styleId: 0 });

    const api = new DocumentControllerSpreadsheetApi(controller);

    const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
    expect(cell).toMatchObject({ value: 2, formula: "=1+1" });

    const range = api.readRange({ sheet: "Sheet1", startRow: 1, endRow: 1, startCol: 1, endCol: 1 });
    expect(range[0]?.[0]).toMatchObject({ value: 2, formula: "=1+1" });

    const executor = new ToolExecutor(api, { default_sheet: "Sheet1", include_formula_values: true });
    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "A1:A1" }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[2]]);
  });

  it("can surface live computed formula values via getCellComputedValueForSheet (gated by include_formula_values)", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "B1", 1);
    controller.setCellFormula("Sheet1", "A1", "B1+1");

    const api = new DocumentControllerSpreadsheetApi(controller, {
      getCellComputedValueForSheet: (sheetId, cell) => {
        if (sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) return 2;
        return null;
      },
    });

    // Adapter should surface a computed `value` even though DocumentController stores `value:null` for formulas.
    expect(api.getCell({ sheet: "Sheet1", row: 1, col: 1 })).toMatchObject({ value: 2, formula: "=B1+1" });

    // ToolExecutor should still treat formula values as opt-in.
    const executorNoValues = new ToolExecutor(api, { default_sheet: "Sheet1" });
    const noValues = await executorNoValues.execute({
      name: "read_range",
      parameters: { range: "A1:A1", include_formulas: true },
    });
    expect(noValues.ok).toBe(true);
    expect(noValues.tool).toBe("read_range");
    if (!noValues.ok || noValues.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(noValues.data?.values).toEqual([[null]]);
    expect(noValues.data?.formulas).toEqual([["=B1+1"]]);

    const executorWithValues = new ToolExecutor(api, { default_sheet: "Sheet1", include_formula_values: true });
    const withValues = await executorWithValues.execute({
      name: "read_range",
      parameters: { range: "A1:A1", include_formulas: true },
    });
    expect(withValues.ok).toBe(true);
    expect(withValues.tool).toBe("read_range");
    if (!withValues.ok || withValues.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(withValues.data?.values).toEqual([[2]]);
    expect(withValues.data?.formulas).toEqual([["=B1+1"]]);
  });

  it("clone() uses a local formula evaluator for computed values (and avoids computing them in listNonEmptyCells)", () => {
    const beforeLocale = getLocale();
    // Use a locale with localized function names + ';' argument separators to ensure
    // the clone evaluator stays locale-aware.
    setLocale("de-DE");
    try {
      const controller = new DocumentController();
      controller.setCellValue("Sheet1", "B1", 1);
      controller.setCellValue("Sheet1", "C1", 2);
      controller.setCellFormula("Sheet1", "A1", "SUMME(B1;C1)");

      // Provide a dummy live provider so clone() will enable local evaluation. The clone should NOT use this.
      const api = new DocumentControllerSpreadsheetApi(controller, {
        getCellComputedValueForSheet: () => 999,
      });
      const cloned = api.clone() as any as DocumentControllerSpreadsheetApi;

      // Local evaluator should compute based on the cloned DocumentController state.
      expect(cloned.getCell({ sheet: "Sheet1", row: 1, col: 1 })).toMatchObject({ value: 3, formula: "=SUMME(B1;C1)" });

      // PreviewEngine diffs rely on listNonEmptyCells; clones intentionally do not compute formula values there.
      const a1Entry = cloned.listNonEmptyCells("Sheet1").find((e) => e.address.row === 1 && e.address.col === 1);
      expect(a1Entry?.cell.formula).toBe("=SUMME(B1;C1)");
      expect(a1Entry?.cell.value).toBeNull();

      // Updating a dependency cell should invalidate any cached computed values.
      cloned.setCell({ sheet: "Sheet1", row: 1, col: 2 }, { value: 10 });
      expect(cloned.getCell({ sheet: "Sheet1", row: 1, col: 1 })).toMatchObject({ value: 12, formula: "=SUMME(B1;C1)" });
    } finally {
      setLocale(beforeLocale);
    }
  });

  it("clone() prefers document.documentElement.lang for locale-aware evaluation when i18n locale is not wired", () => {
    const beforeLocale = getLocale();
    const beforeDocument = (globalThis as any).document;
    try {
      (globalThis as any).document = { documentElement: { lang: "de_DE.UTF-8" } };
      // Keep i18n locale at the default (en-US) while the document is set to a de locale.
      setLocale("en-US");
      (globalThis as any).document.documentElement.lang = "de_DE.UTF-8";

      const controller = new DocumentController();
      controller.setCellValue("Sheet1", "B1", 1);
      controller.setCellValue("Sheet1", "C1", 2);
      controller.setCellFormula("Sheet1", "A1", "SUMME(B1;C1)");

      // Provide a dummy live provider so clone() will enable local evaluation. The clone should NOT use this.
      const api = new DocumentControllerSpreadsheetApi(controller, {
        getCellComputedValueForSheet: () => 999,
      });
      const cloned = api.clone() as any as DocumentControllerSpreadsheetApi;

      expect(cloned.getCell({ sheet: "Sheet1", row: 1, col: 1 })).toMatchObject({ value: 3, formula: "=SUMME(B1;C1)" });
    } finally {
      setLocale(beforeLocale);
      if (beforeDocument === undefined) {
        delete (globalThis as any).document;
      } else {
        (globalThis as any).document = beforeDocument;
      }
    }
  });

  it("read_range returns primitive values + formulas without per-cell controller.getCell calls", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 123);
    controller.setCellValue("Sheet1", "B1", "hello");
    controller.setCellValue("Sheet1", "C1", true);
    // Store without leading '=' to ensure adapter normalization still matches tool semantics.
    controller.setCellFormula("Sheet1", "D1", "SUM(A1:C1)");

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const getCellSpy = vi.spyOn(controller, "getCell");
    getCellSpy.mockClear();

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "A1:D1", include_formulas: true }
    });

    expect(getCellSpy).not.toHaveBeenCalled();

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[123, "hello", true, null]]);
    expect(result.data?.formulas).toEqual([[null, null, null, "=SUM(A1:C1)"]]);
  });

  it("readRange does not leak mutable references for object cell values", () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", {
      text: "Rich Bold",
      runs: [{ start: 0, end: 4, style: { bold: true } }]
    });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const cells = api.readRange({ sheet: "Sheet1", startRow: 1, endRow: 1, startCol: 1, endCol: 1 });

    const value = cells[0]?.[0]?.value as any;
    expect(value?.text).toBe("Rich Bold");
    value.text = "Mutated";

    const after = controller.getCell("Sheet1", "A1").value as any;
    expect(after?.text).toBe("Rich Bold");
  });

  it("read_range DLP redaction works against DocumentController adapter output shapes", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", "ok");
    controller.setCellValue("Sheet1", "B1", "secret");
    controller.setCellValue("Sheet1", "C1", 123);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, {
      default_sheet: "Sheet1",
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
        ]
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
  });

  it("read_range DLP redaction still applies when the tool call uses a display sheet name", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet2", "A1", "ok");
    controller.setCellValue("Sheet2", "B1", "secret");
    controller.setCellValue("Sheet2", "C1", 123);

    const sheetNames = new Map<string, string>([["Sheet2", "Budget"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetNames);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, {
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
        ]
      }
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Budget!A1:C1", include_formulas: true }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.range).toBe("Budget!A1:C1");
    expect(result.data?.values).toEqual([["ok", "[REDACTED]", 123]]);
    expect(result.data?.formulas).toEqual([[null, "[REDACTED]", null]]);
  });
});
