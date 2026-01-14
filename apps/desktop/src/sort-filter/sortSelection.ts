import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import { showToast } from "../extensions/ui.js";
import { normalizeSelectionRange } from "../formatting/selectionSizeGuard.js";
import type { CellCoord, Range } from "../selection/types";
import type { SpreadsheetValue } from "../spreadsheet/evaluateFormula";
import { parseImageCellValue } from "../shared/imageCellValue.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";

import type { SortKey, SortOrder, SortSpec } from "./types";

// Sorting re-materializes the full selection into JS arrays (values + style ids).
// Keep this bounded so users can't accidentally trigger million-cell allocations.
export const DEFAULT_SORT_CELL_LIMIT = 50_000;

type CellSnapshot = { value: unknown; formula: string | null; styleId: number };

/**
 * DocumentController creates sheets lazily when referenced by `getCell()` / `getSheetView()`.
 *
 * Sort operations read a large amount of cell state and can run while UI state is briefly
 * pointing at a deleted sheet (e.g. during undo/redo of sheet structure edits). Treat those
 * ids as "known missing" so sorting can't resurrect a deleted sheet as a side-effect.
 */
function isSheetKnownMissing(doc: DocumentController, sheetId: string): boolean {
  const id = String(sheetId ?? "").trim();
  if (!id) return true;

  const docAny = doc as any;
  const sheets: any = docAny?.model?.sheets;
  const sheetMeta: any = docAny?.sheetMeta;

  if (
    sheets &&
    typeof sheets.has === "function" &&
    typeof sheets.size === "number" &&
    sheetMeta &&
    typeof sheetMeta.has === "function" &&
    typeof sheetMeta.size === "number"
  ) {
    const workbookHasAnySheets = sheets.size > 0 || sheetMeta.size > 0;
    if (!workbookHasAnySheets) return false;
    return !sheets.has(id) && !sheetMeta.has(id);
  }

  return false;
}

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
    const image = parseImageCellValue(raw);
    if (image) {
      // In-cell images do not have a meaningful numeric sort key. When alt text is present
      // treat it as a string; otherwise treat the image as blank so it sorts last.
      if (image.altText != null) return { kind: "string", value: image.altText };
      return { kind: "blank", value: "" };
    }

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
  options: { order: SortOrder; maxCells?: number; getCellValue?: GetSortValue } = { order: "ascending" },
): SortRangeRowsResult {
  // Avoid resurrecting deleted sheets when callers hold a stale sheet id.
  if (isSheetKnownMissing(doc, sheetId)) return { applied: false, reason: "invalidRange" };

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
  const getCellValue = options.getCellValue;

  for (let row = startRow; row <= endRow; row += 1) {
    const rowCells: CellSnapshot[] = [];
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = doc.getCell(sheetId, { row, col }) as { value: unknown; formula: string | null; styleId: number };
      rowCells.push({ value: cell.value, formula: cell.formula ?? null, styleId: cell.styleId ?? 0 });
    }
    rows.push(rowCells);
    const keyCell = rowCells[keyCol - startCol];
    keys.push(getCellValue ? getCellValue({ row, col: keyCol }) : effectiveCellValue(keyCell!));
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
  // Sorting mutates cell values/styles. Even though `DocumentController.canEditCell` will prevent
  // edits in collab read-only roles, guard early here so command palette / programmatic execution
  // produces an explicit (and cheaper) UX outcome.
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const appAny = app as any;
    if (typeof appAny?.isReadOnly === "function" && appAny.isReadOnly() === true) {
      showCollabEditRejectedToast([{ rejectionKind: "sort", rejectionReason: "permission" }]);
      try {
        app.focus();
      } catch {
        // ignore
      }
      return;
    }
  } catch {
    // ignore
  }

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
 
   // Reject partial sorts when any cell in the target range is non-writable (protected/encrypted).
   // `DocumentController` filters disallowed deltas per-cell, which can corrupt row integrity if
   // we attempt a sort and some writes are silently skipped.
   //
   // When `canEditCell` is unavailable (local/non-collab), treat all cells as writable.
   const normalized = normalizeSelectionRange(range);
   const normRowCount = normalized.endRow - normalized.startRow + 1;
   const normColCount = normalized.endCol - normalized.startCol + 1;
   const normCellCount = normRowCount * normColCount;
   if (normCellCount > DEFAULT_SORT_CELL_LIMIT) {
     showToast(
       `Selection too large to sort (${normCellCount.toLocaleString()} cells). ` +
         `Reduce the selection to under ${DEFAULT_SORT_CELL_LIMIT.toLocaleString()} cells and try again.`,
       "warning",
     );
     app.focus();
     return;
   }
 
   // eslint-disable-next-line @typescript-eslint/no-explicit-any
   const canEditCell = (doc as any)?.canEditCell as
     | ((cell: { sheetId: string; row: number; col: number }) => boolean)
     | null
     | undefined;
   if (typeof canEditCell === "function") {
     let blockedCell: { row: number; col: number } | null = null;
     try {
       outer: for (let row = normalized.startRow; row <= normalized.endRow; row += 1) {
         for (let col = normalized.startCol; col <= normalized.endCol; col += 1) {
           if (!canEditCell.call(doc, { sheetId, row, col })) {
             blockedCell = { row, col };
             break outer;
           }
         }
       }
     } catch {
       blockedCell = null;
     }
     if (blockedCell) {
       const rejection = (() => {
         try {
           // eslint-disable-next-line @typescript-eslint/no-explicit-any
           const appAny = app as any;
           const infer = appAny?.inferCollabEditRejection;
           if (typeof infer === "function") {
             const inferred = infer.call(appAny, { sheetId, row: blockedCell.row, col: blockedCell.col }) as any;
             if (inferred && typeof inferred === "object" && typeof inferred.rejectionReason === "string") {
               return inferred as {
                 rejectionReason: "permission" | "encryption" | "unknown";
                 encryptionKeyId?: string;
                 encryptionPayloadUnsupported?: boolean;
               };
             }
           }
         } catch {
           // ignore
         }
         return { rejectionReason: "permission" as const };
       })();
       showCollabEditRejectedToast([
         {
           rejectionKind: "sort",
           ...rejection,
         },
       ]);
       app.focus();
       return;
     }
   }

  const result = sortRangeRowsInDocument(doc, sheetId, range, activeCell, {
    order: options.order,
    getCellValue: (cell) => app.getCellComputedValueForSheet(sheetId, cell),
  });
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

export type GetSortValue = (cell: { row: number; col: number }) => SpreadsheetValue;

function clampKeys(keys: SortKey[], width: number): SortKey[] {
  const out: SortKey[] = [];
  for (const key of keys) {
    if (!key || typeof key.column !== "number" || typeof key.order !== "string") continue;
    const col = Math.trunc(key.column);
    if (col < 0 || col >= width) continue;
    out.push({ column: col, order: key.order });
  }
  return out;
}

/**
 * Apply a multi-key sort spec to the current selection.
 *
 * - `spec.keys[].column` are 0-based indices relative to `selection.startCol`.
 * - When `spec.hasHeader` is true, the first row of the selection is kept fixed.
 */
export function applySortSpecToSelection(params: {
  doc: DocumentController;
  sheetId: string;
  selection: Range;
  spec: SortSpec;
  getCellValue: GetSortValue;
  maxCells?: number;
  label?: string;
}): boolean {
  // Avoid resurrecting deleted sheets when callers hold a stale sheet id.
  if (isSheetKnownMissing(params.doc, params.sheetId)) return false;

  const { startRow, endRow, startCol, endCol } = normalizeSelectionRange(params.selection);
  if (startRow < 0 || startCol < 0) return false;

  const rowCount = endRow - startRow + 1;
  const colCount = endCol - startCol + 1;
  if (rowCount <= 0 || colCount <= 0) return false;

  const cellCount = rowCount * colCount;
  const maxCells = params.maxCells ?? DEFAULT_SORT_CELL_LIMIT;
  if (cellCount > maxCells) return false;

  const keys = clampKeys(params.spec.keys, colCount);
  if (keys.length === 0) return false;

  const dataStartRow = params.spec.hasHeader ? startRow + 1 : startRow;
  // Nothing to sort.
  if (dataStartRow > endRow) return true;
  const dataHeight = endRow - dataStartRow + 1;
  if (dataHeight <= 1) return true;
 
  // Reject partial sorts when any cell in the sort payload is non-writable. `DocumentController`
  // filters disallowed deltas per-cell; for sort that can corrupt row integrity.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const canEditCell = (params.doc as any)?.canEditCell as
    | ((cell: { sheetId: string; row: number; col: number }) => boolean)
    | null
    | undefined;
  if (typeof canEditCell === "function") {
    try {
      for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          if (!canEditCell.call(params.doc, { sheetId: params.sheetId, row, col })) {
            return false;
          }
        }
      }
     } catch {
       // Best-effort: if `canEditCell` throws, fall through to attempting the sort.
     }
   }

  type RowRecord = {
    originalIndex: number;
    keyValues: SpreadsheetValue[];
    cells: CellSnapshot[];
  };

  const rows: RowRecord[] = [];
  for (let i = 0; i < dataHeight; i += 1) {
    const row = dataStartRow + i;

    const keyValues = keys.map((k) => params.getCellValue({ row, col: startCol + k.column }));

    const cells: CellSnapshot[] = [];
    for (let c = 0; c < colCount; c += 1) {
      const state = params.doc.getCell(params.sheetId, { row, col: startCol + c }) as {
        value: unknown;
        formula: string | null;
        styleId: number;
      };
      cells.push({ value: state.value, formula: state.formula ?? null, styleId: state.styleId ?? 0 });
    }

    rows.push({ originalIndex: i, keyValues, cells });
  }

  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i += 1) {
      const key = keys[i]!;
      const cmp = compareEffectiveValues(a.keyValues[i], b.keyValues[i], key.order);
      if (cmp !== 0) return cmp;
    }
    // Stable tie-breaker.
    return a.originalIndex - b.originalIndex;
  });

  const sortedValues: CellSnapshot[][] = rows.map((r) => r.cells);

  const label = params.label ?? "Sort";
  params.doc.beginBatch({ label });
  try {
    params.doc.setRangeValues(params.sheetId, { row: dataStartRow, col: startCol }, sortedValues);
    params.doc.endBatch();
  } catch (err) {
    params.doc.cancelBatch();
    throw err;
  }

  return true;
}
