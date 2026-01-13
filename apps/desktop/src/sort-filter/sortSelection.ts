import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import { showToast } from "../extensions/ui.js";
import { normalizeSelectionRange } from "../formatting/selectionSizeGuard.js";
import type { CellCoord, Range } from "../selection/types";
import type { SortOrder } from "./types";

export const DEFAULT_SORT_CELL_LIMIT = 50_000;

type CellSnapshot = { value: unknown; formula: string | null; styleId: number };

export type SortRangeRowsResult =
  | { applied: true; rowCount: number; colCount: number; keyCol: number }
  | { applied: false; reason: "invalidRange" | "tooLarge"; cellCount?: number; maxCells?: number };

function isCellBlank(value: unknown): boolean {
  return value == null || value === "";
}

function valueForCompare(raw: unknown): { kind: "blank" | "number" | "string"; value: number | string } {
  if (isCellBlank(raw)) return { kind: "blank", value: "" };

  if (typeof raw === "number" && Number.isFinite(raw)) {
    return { kind: "number", value: raw };
  }

  if (typeof raw === "string") {
    return { kind: "string", value: raw };
  }

  if (typeof raw === "boolean") {
    return { kind: "string", value: raw ? "TRUE" : "FALSE" };
  }

  if (raw && typeof raw === "object") {
    const maybeText = (raw as any)?.text;
    if (typeof maybeText === "string") {
      return { kind: "string", value: maybeText };
    }
  }

  return { kind: "string", value: String(raw) };
}

function compareEffectiveValues(aRaw: unknown, bRaw: unknown, order: SortOrder): number {
  const a = valueForCompare(aRaw);
  const b = valueForCompare(bRaw);

  // Always sort blanks last (Excel-like), regardless of ascending/descending.
  if (a.kind === "blank" && b.kind === "blank") return 0;
  if (a.kind === "blank") return 1;
  if (b.kind === "blank") return -1;

  const direction = order === "descending" ? -1 : 1;

  if (a.kind === "number" && b.kind === "number") {
    const delta = (a.value as number) - (b.value as number);
    if (delta === 0) return 0;
    return delta > 0 ? direction : -direction;
  }

  const cmp = String(a.value).localeCompare(String(b.value));
  if (cmp === 0) return 0;
  return direction * cmp;
}

function effectiveCellValue(cell: { value: unknown; formula: string | null }): unknown {
  return cell.value ?? cell.formula;
}

/**
 * Sort rows within a rectangular range in-place.
 *
 * This is the "pure" implementation used by ribbon commands; it avoids any direct UI calls
 * so it can be unit tested without a DOM.
 */
export function sortRangeRowsInDocument(
  doc: DocumentController,
  sheetId: string,
  range: Range,
  activeCell: CellCoord,
  options: { order: SortOrder; maxCells?: number } = { order: "ascending" },
): SortRangeRowsResult {
  const { startRow, endRow, startCol, endCol } = normalizeSelectionRange(range);
  if (startRow < 0 || startCol < 0) return { applied: false, reason: "invalidRange" };

  const rowCount = endRow - startRow + 1;
  const colCount = endCol - startCol + 1;
  if (rowCount <= 0 || colCount <= 0) return { applied: false, reason: "invalidRange" };

  const cellCount = rowCount * colCount;
  const maxCells = options.maxCells ?? DEFAULT_SORT_CELL_LIMIT;
  if (cellCount > maxCells) {
    return { applied: false, reason: "tooLarge", cellCount, maxCells };
  }

  const activeInRange =
    activeCell.row >= startRow && activeCell.row <= endRow && activeCell.col >= startCol && activeCell.col <= endCol;
  const keyCol = activeInRange ? activeCell.col : startCol;

  const rows: CellSnapshot[][] = [];
  const keys: unknown[] = [];

  for (let row = startRow; row <= endRow; row++) {
    const rowCells: CellSnapshot[] = [];
    for (let col = startCol; col <= endCol; col++) {
      const cell = doc.getCell(sheetId, { row, col });
      rowCells.push({ value: cell.value, formula: cell.formula ?? null, styleId: cell.styleId ?? 0 });
    }
    rows.push(rowCells);
    const keyCell = rowCells[keyCol - startCol];
    keys.push(effectiveCellValue(keyCell));
  }

  const order = options.order;
  const sortedRowIndices = rows
    .map((_, idx) => idx)
    .sort((aIdx, bIdx) => {
      const cmp = compareEffectiveValues(keys[aIdx], keys[bIdx], order);
      if (cmp !== 0) return cmp;
      // Stable tiebreaker: preserve original order for equal keys.
      return aIdx - bIdx;
    });

  const sortedValues = sortedRowIndices.map((idx) => rows[idx]!);

  doc.beginBatch({ label: "Sort" });
  try {
    doc.setRangeValues(sheetId, { row: startRow, col: startCol }, sortedValues);
    doc.endBatch();
  } catch (err) {
    // Keep undo history consistent if anything goes wrong while applying the sort.
    doc.cancelBatch();
    throw err;
  }

  return { applied: true, rowCount, colCount, keyCol };
}

/**
 * UI wrapper for ribbon commands. Resolves the selection/active cell from the app and
 * shows toasts for unsupported scenarios.
 */
export function sortSelection(app: SpreadsheetApp, options: { order: SortOrder }): void {
  const ranges = app.getSelectionRanges();
  const activeCell = app.getActiveCell();

  let range: Range | null = null;
  if (ranges.length === 0) {
    range = { startRow: activeCell.row, endRow: activeCell.row, startCol: activeCell.col, endCol: activeCell.col };
  } else if (ranges.length === 1) {
    range = ranges[0] ?? null;
  } else {
    showToast("Sorting multiple ranges isn't supported yet. Select a single rectangular range and try again.", "warning");
    app.focus();
    return;
  }

  if (!range) {
    app.focus();
    return;
  }

  const sheetId = app.getCurrentSheetId();
  const doc = app.getDocument();

  const result = sortRangeRowsInDocument(doc, sheetId, range, activeCell, { order: options.order });
  if (!result.applied) {
    if (result.reason === "tooLarge") {
      showToast(
        `Selection too large to sort (${(result.cellCount ?? 0).toLocaleString()} cells). ` +
          `Reduce the selection to under ${(result.maxCells ?? DEFAULT_SORT_CELL_LIMIT).toLocaleString()} cells and try again.`,
        "warning",
      );
    } else {
      showToast("Unable to sort selection.", "error");
    }
  }

  app.focus();
}
