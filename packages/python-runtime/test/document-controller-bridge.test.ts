import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import { NativePythonRuntime } from "@formula/python-runtime/native";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

describe("DocumentControllerBridge", () => {
  it("lets Python scripts write values + formulas into a DocumentController", async () => {
    const doc = new DocumentController();
    const api = new DocumentControllerBridge(doc, { activeSheetId: "Sheet1" });
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

sheet = formula.active_sheet
sheet["A1"] = 42
sheet["A2"] = "=A1*2"
`;

    await runtime.execute(script, { api });

    expect(doc.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(42);
    expect(doc.getCell("Sheet1", { row: 1, col: 0 }).formula).toBe("=A1*2");
  });

  it("returns effective (layered) formats when formats are inherited from a full-column style", () => {
    const MAX_ROW = 1_048_575; // Excel rows are 1..1048576 (0-based index max).

    class LayeredFormatDoc {
      private readonly columnFormats = new Map<string, any>();

      setRangeFormat(
        sheetId: string,
        range: { start: { row: number; col: number }; end: { row: number; col: number } },
        stylePatch: any,
      ) {
        const isFullColumn = range.start.col === range.end.col && range.start.row === 0 && range.end.row === MAX_ROW;
        if (!isFullColumn) {
          throw new Error("Test doc only supports full-column formats");
        }
        this.columnFormats.set(`${sheetId}:${range.start.col}`, stylePatch ?? {});
      }

      // Legacy per-cell API still exists, but returns no direct formatting (styleId=0).
      getCell(_sheetId: string, _coord: { row: number; col: number }) {
        return { styleId: 0 };
      }

      // New effective format API used by the bridge.
      getCellFormat(sheetId: string, coord: { row: number; col: number }) {
        return this.columnFormats.get(`${sheetId}:${coord.col}`) ?? {};
      }
    }

    const doc = new LayeredFormatDoc();
    const api = new DocumentControllerBridge(doc as any, { activeSheetId: "Sheet1" });

    api.set_range_format({
      range: { sheet_id: "Sheet1", start_row: 0, end_row: MAX_ROW, start_col: 0, end_col: 0 },
      format: { font: { bold: true } },
    });

    const format = api.get_range_format({
      range: { sheet_id: "Sheet1", start_row: 999, end_row: 999, start_col: 0, end_col: 0 },
    });

    expect(format).toMatchObject({ font: { bold: true } });
  });
});
