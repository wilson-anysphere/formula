import type { DocumentController } from "../document/documentController.js";
import { showInputBox, showToast } from "../extensions/ui.js";
import type { Range } from "../selection/types";

export const MAX_AXIS_RESIZE_INDICES = 10_000;

export type AxisSizingKind = "rowHeight" | "colWidth";

export type AxisSizingApp = {
  getSelectionRanges(): Range[];
  getCurrentSheetId(): string;
  getDocument(): DocumentController;
  focus(): void;
  isEditing(): boolean;
  /**
   * Optional read-only indicator (used in collab viewer/commenter sessions).
   *
   * In read-only roles, axis sizing should still be allowed as a **local-only** view interaction.
   * Collaboration binders are responsible for preventing these mutations from being persisted into
   * the shared Yjs document.
   */
  isReadOnly?: () => boolean;
};

type NormalizedRange = { startRow: number; endRow: number; startCol: number; endCol: number };

function normalizeSelectionRange(range: Range): NormalizedRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

export function selectedRowIndices(ranges: Range[]): number[] {
  const rows = new Set<number>();
  for (const range of ranges) {
    const r = normalizeSelectionRange(range);
    for (let row = r.startRow; row <= r.endRow; row += 1) rows.add(row);
  }
  return [...rows].sort((a, b) => a - b);
}

export function selectedColIndices(ranges: Range[]): number[] {
  const cols = new Set<number>();
  for (const range of ranges) {
    const r = normalizeSelectionRange(range);
    for (let col = r.startCol; col <= r.endCol; col += 1) cols.add(col);
  }
  return [...cols].sort((a, b) => a - b);
}

function estimateSelectedRowUpperBound(ranges: Range[]): number {
  let rowUpperBound = 0;
  for (const range of ranges) {
    const r = normalizeSelectionRange(range);
    rowUpperBound += Math.max(0, r.endRow - r.startRow + 1);
    if (rowUpperBound > MAX_AXIS_RESIZE_INDICES) break;
  }
  return rowUpperBound;
}

function estimateSelectedColUpperBound(ranges: Range[]): number {
  let colUpperBound = 0;
  for (const range of ranges) {
    const r = normalizeSelectionRange(range);
    colUpperBound += Math.max(0, r.endCol - r.startCol + 1);
    if (colUpperBound > MAX_AXIS_RESIZE_INDICES) break;
  }
  return colUpperBound;
}

export async function promptAndApplyAxisSizing(
  app: AxisSizingApp,
  kind: AxisSizingKind,
  options: { isEditing?: () => boolean } = {},
): Promise<void> {
  const isEditing = options.isEditing ?? (() => app.isEditing() || (globalThis as any).__formulaSpreadsheetIsEditing === true);
  if (isEditing()) return;

  // `selectedRowIndices()` / `selectedColIndices()` enumerate every row/col in every selection range into a Set.
  // On Excel-scale sheets, this can freeze/crash the UI (e.g. select-all => 1M rows).
  // Guard here so we reject huge selections *before* prompting for input.
  const selection = app.getSelectionRanges();
  const axisUpperBound = kind === "rowHeight" ? estimateSelectedRowUpperBound(selection) : estimateSelectedColUpperBound(selection);
  if (axisUpperBound > MAX_AXIS_RESIZE_INDICES) {
    showToast(
      kind === "rowHeight"
        ? "Selection too large to resize rows. Select fewer rows and try again."
        : "Selection too large to resize columns. Select fewer columns and try again.",
      "warning",
    );
    return;
  }

  const label = kind === "rowHeight" ? "Row Height" : "Column Width";
  const placeHolder = kind === "rowHeight" ? "Enter a row height (px)" : "Enter a column width (px)";
  const input = await showInputBox({ prompt: label, placeHolder });
  if (input == null) return;
  const size = Number(input);
  if (!Number.isFinite(size) || size <= 0) {
    showToast(kind === "rowHeight" ? "Row height must be a positive number." : "Column width must be a positive number.", "error");
    return;
  }

  const indices = kind === "rowHeight" ? selectedRowIndices(selection) : selectedColIndices(selection);
  if (indices.length === 0) return;

  const sheetId = app.getCurrentSheetId();
  const doc = app.getDocument();
  doc.beginBatch({ label });
  try {
    for (const index of indices) {
      if (kind === "rowHeight") doc.setRowHeight(sheetId, index, size, { label });
      else doc.setColWidth(sheetId, index, size, { label });
    }
  } finally {
    doc.endBatch();
  }

  app.focus();
}
