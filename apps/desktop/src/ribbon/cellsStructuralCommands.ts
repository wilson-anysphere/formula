import type { DocumentController } from "../document/documentController.js";
import type { GridLimits, Range } from "../selection/types";
import { DEFAULT_DESKTOP_LOAD_MAX_COLS, DEFAULT_DESKTOP_LOAD_MAX_ROWS } from "../workbook/load/clampUsedRange.js";

export type CellsStructuralCommandId =
  | "home.cells.insert.insertSheetRows"
  | "home.cells.insert.insertSheetColumns"
  | "home.cells.delete.deleteSheetRows"
  | "home.cells.delete.deleteSheetColumns";

export type CellsStructuralCommandApp = {
  isEditing(): boolean;
  isReadOnly?(): boolean;
  getDocument(): DocumentController;
  getCurrentSheetId(): string;
  getSelectionRanges(): Range[];
  getActiveCell(): { row: number; col: number };
  getGridLimits(): GridLimits;
  focus(): void;
};

function normalizeSelectionRange(range: Range): { startRow: number; endRow: number; startCol: number; endCol: number } {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function resolveGridLimits(app: CellsStructuralCommandApp): { maxRows: number; maxCols: number } {
  const raw = app.getGridLimits();
  const maxRows =
    Number.isInteger(raw?.maxRows) && raw.maxRows > 0 ? raw.maxRows : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
  const maxCols =
    Number.isInteger(raw?.maxCols) && raw.maxCols > 0 ? raw.maxCols : DEFAULT_DESKTOP_LOAD_MAX_COLS;
  return { maxRows, maxCols };
}

/**
 * Execute Excel-style "Insert/Delete Sheet Rows/Columns" commands from the Home > Cells group.
 *
 * This is factored out of `main.ts` so it can be unit-tested without booting the full desktop UI.
 */
export function executeCellsStructuralRibbonCommand(app: CellsStructuralCommandApp, commandId: string): boolean {
  const id = commandId as CellsStructuralCommandId;
  if (
    id !== "home.cells.insert.insertSheetRows" &&
    id !== "home.cells.insert.insertSheetColumns" &&
    id !== "home.cells.delete.deleteSheetRows" &&
    id !== "home.cells.delete.deleteSheetColumns"
  ) {
    return false;
  }

  // Match SpreadsheetApp guards: never mutate while editing.
  if (app.isEditing()) return true;
  if (typeof app.isReadOnly === "function" && app.isReadOnly()) return true;

  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  const ranges = app.getSelectionRanges();
  const active = app.getActiveCell();
  const limits = resolveGridLimits(app);

  const primaryRange = ranges.length > 0 ? normalizeSelectionRange(ranges[0]!) : null;
  const startRow = primaryRange?.startRow ?? active.row;
  const endRow = primaryRange?.endRow ?? active.row;
  const startCol = primaryRange?.startCol ?? active.col;
  const endCol = primaryRange?.endCol ?? active.col;

  const isFullRowBand = startCol === 0 && endCol === limits.maxCols - 1;
  const isFullColBand = startRow === 0 && endRow === limits.maxRows - 1;

  switch (id) {
    case "home.cells.insert.insertSheetRows": {
      const row0 = isFullRowBand ? startRow : active.row;
      const count = isFullRowBand ? endRow - startRow + 1 : 1;
      doc.insertRows(sheetId, row0, count, { label: "Insert Rows", source: "ribbon" });
      app.focus();
      return true;
    }
    case "home.cells.insert.insertSheetColumns": {
      const col0 = isFullColBand ? startCol : active.col;
      const count = isFullColBand ? endCol - startCol + 1 : 1;
      doc.insertCols(sheetId, col0, count, { label: "Insert Columns", source: "ribbon" });
      app.focus();
      return true;
    }
    case "home.cells.delete.deleteSheetRows": {
      const row0 = isFullRowBand ? startRow : active.row;
      const count = isFullRowBand ? endRow - startRow + 1 : 1;
      doc.deleteRows(sheetId, row0, count, { label: "Delete Rows", source: "ribbon" });
      app.focus();
      return true;
    }
    case "home.cells.delete.deleteSheetColumns": {
      const col0 = isFullColBand ? startCol : active.col;
      const count = isFullColBand ? endCol - startCol + 1 : 1;
      doc.deleteCols(sheetId, col0, count, { label: "Delete Columns", source: "ribbon" });
      app.focus();
      return true;
    }
    default:
      return false;
  }
}
