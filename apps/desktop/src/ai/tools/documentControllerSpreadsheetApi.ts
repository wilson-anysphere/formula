import { DocumentController } from "../../document/documentController.js";

import type { CellAddress, RangeAddress } from "../../../../../packages/ai-tools/src/spreadsheet/a1.js";
import type { CellEntry, SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import type { CellData, CellFormat } from "../../../../../packages/ai-tools/src/spreadsheet/types.js";

function toCellData(cellState: any): CellData {
  return {
    value: cellState?.value ?? null,
    ...(cellState?.formula ? { formula: String(cellState.formula) } : {}),
    ...(cellState?.format ? { format: cellState.format as CellFormat } : {})
  };
}

function toControllerCoord(address: { row: number; col: number }): { row: number; col: number } {
  return { row: address.row - 1, col: address.col - 1 };
}

function toControllerRange(range: RangeAddress): { start: { row: number; col: number }; end: { row: number; col: number } } {
  return {
    start: { row: range.startRow - 1, col: range.startCol - 1 },
    end: { row: range.endRow - 1, col: range.endCol - 1 }
  };
}

function parseRowColKey(key: string): { row: number; col: number } {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid sheet cell key: ${key}`);
  }
  return { row, col };
}

/**
 * Adapter that lets `packages/ai-tools` execute tool calls against the real
 * `DocumentController` workbook model (used by the desktop app).
 *
 * This makes it possible to reuse the tool executor + preview engine on top of
 * the UI controller before the Rust calc engine integration exists.
 */
export class DocumentControllerSpreadsheetApi implements SpreadsheetApi {
  readonly controller: DocumentController;

  constructor(controller: DocumentController) {
    this.controller = controller;
  }

  listSheets(): string[] {
    const sheets = (this.controller as any).model?.sheets;
    if (!sheets || typeof sheets.keys !== "function") return ["Sheet1"];
    const ids = Array.from(sheets.keys());
    // DocumentController creates sheets lazily; expose the default sheet name so
    // downstream consumers (RAG context, etc) can still reason about "Sheet1"
    // even before any edits.
    return ids.length > 0 ? ids : ["Sheet1"];
  }

  listNonEmptyCells(sheet?: string): CellEntry[] {
    const sheets = (this.controller as any).model?.sheets;
    if (!sheets || typeof sheets.get !== "function") return [];

    const sheetIds = sheet ? [sheet] : Array.from(sheets.keys());
    const entries: CellEntry[] = [];
    for (const sheetId of sheetIds) {
      const sheetModel = sheets.get(sheetId);
      if (!sheetModel) continue;
      for (const [key, state] of sheetModel.cells?.entries?.() ?? []) {
        const { row, col } = parseRowColKey(key);
        entries.push({
          address: { sheet: sheetId, row: row + 1, col: col + 1 },
          cell: toCellData(state)
        });
      }
    }
    return entries;
  }

  getCell(address: CellAddress): CellData {
    const state = this.controller.getCell(address.sheet, toControllerCoord(address));
    return toCellData(state);
  }

  setCell(address: CellAddress, cell: CellData): void {
    const coord = toControllerCoord(address);
    if (cell.formula) {
      this.controller.setCellFormula(address.sheet, coord, cell.formula, { label: "AI write_cell" });
    } else {
      this.controller.setCellValue(address.sheet, coord, cell.value ?? null, { label: "AI write_cell" });
    }

    if (cell.format && Object.keys(cell.format).length > 0) {
      this.controller.setRangeFormat(address.sheet, { start: coord, end: coord }, cell.format as any, {
        label: "AI apply_formatting"
      });
    }
  }

  readRange(range: RangeAddress): CellData[][] {
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellData[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this.getCell({ sheet: range.sheet, row: r, col: c }));
      }
      rows.push(row);
    }
    return rows;
  }

  writeRange(range: RangeAddress, cells: CellData[][]): void {
    const values = cells.map((row) =>
      row.map((cell) => (cell?.formula ? { formula: cell.formula } : cell?.value ?? null))
    );
    this.controller.setRangeValues(range.sheet, toControllerRange(range), values, { label: "AI set_range" });
  }

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number {
    this.controller.setRangeFormat(range.sheet, toControllerRange(range), format as any, { label: "AI apply_formatting" });
    const rows = range.endRow - range.startRow + 1;
    const cols = range.endCol - range.startCol + 1;
    return rows * cols;
  }

  getLastUsedRow(sheet: string): number {
    const sheets = (this.controller as any).model?.sheets;
    const sheetModel = sheets?.get?.(sheet);
    if (!sheetModel) return 0;
    let max = 0;
    for (const key of sheetModel.cells?.keys?.() ?? []) {
      const { row } = parseRowColKey(key);
      max = Math.max(max, row + 1);
    }
    return max;
  }

  clone(): SpreadsheetApi {
    const snapshot = this.controller.encodeState();
    const cloned = new DocumentController();
    cloned.applyState(snapshot);
    return new DocumentControllerSpreadsheetApi(cloned);
  }
}
