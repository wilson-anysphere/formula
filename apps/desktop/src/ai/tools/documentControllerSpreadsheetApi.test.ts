import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { ToolExecutor, PreviewEngine } from "../../../../../packages/ai-tools/src/index.js";

import { DocumentControllerSpreadsheetApi } from "./documentControllerSpreadsheetApi.js";

describe("DocumentControllerSpreadsheetApi", () => {
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

  it("includes supported formatting-only cells in listNonEmptyCells()", () => {
    const controller = new DocumentController();
    controller.setRangeFormat("Sheet1", "A1", { font: { bold: true } }, { label: "Bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.listNonEmptyCells("Sheet1")).toEqual([
      { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: null, format: { bold: true } } }
    ]);
  });

  it("supports legacy flat CellFormat keys stored in DocumentController styles", () => {
    const controller = new DocumentController();
    // Simulate the pre-fix adapter behavior that wrote ai-tools CellFormat objects directly
    // into the DocumentController style table (flat keys like `bold`).
    controller.setRangeFormat("Sheet1", "A1", { bold: true, background_color: "#FFFFFF00" }, { label: "Legacy bold" });

    const api = new DocumentControllerSpreadsheetApi(controller);
    expect(api.listNonEmptyCells("Sheet1")).toEqual([
      { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: null, format: { bold: true, background_color: "#FFFFFF00" } } }
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
      { default_sheet: "Sheet1" }
    );

    expect(preview.summary.total_changes).toBe(1);
    expect(preview.summary.modifies).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(10);

    // Sanity check: tool call references normalize to Sheet1.
    expect(preview.changes[0]?.cell).toBe("Sheet1!A1");
  });

  it("detects formatting-only changes in PreviewEngine diffs", async () => {
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
      { default_sheet: "Sheet1" }
    );

    expect(preview.summary.total_changes).toBe(1);
    expect(preview.summary.modifies).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(preview.changes[0]?.cell).toBe("Sheet1!A1");

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
});
