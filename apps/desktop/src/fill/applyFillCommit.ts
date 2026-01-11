import { computeFillEdits, type CellRange, type FillMode, type FillSourceCell } from "@formula/fill-engine";

import type { DocumentController } from "../document/documentController.js";

export interface ApplyFillCommitOptions {
  document: DocumentController;
  sheetId: string;
  sourceRange: CellRange;
  targetRange: CellRange;
  mode: FillMode;
  /**
   * Optional hook to prevent overwriting protected/locked cells.
   *
   * Return `false` to skip emitting/applying an edit for the cell.
   */
  canWriteCell?: (row: number, col: number) => boolean;
}

/**
 * Apply a fill-handle commit to the desktop DocumentController.
 *
 * This uses `DocumentController.beginBatch/endBatch` so a single fill drag is
 * recorded as one undo step.
 */
export function applyFillCommitToDocumentController(options: ApplyFillCommitOptions): { editsApplied: number } {
  const { document: doc, sheetId, sourceRange, targetRange, mode, canWriteCell } = options;

  const height = Math.max(0, sourceRange.endRow - sourceRange.startRow);
  const width = Math.max(0, sourceRange.endCol - sourceRange.startCol);
  if (height === 0 || width === 0) return { editsApplied: 0 };

  const sourceCells: FillSourceCell[][] = [];
  for (let row = sourceRange.startRow; row < sourceRange.endRow; row++) {
    const outRow: FillSourceCell[] = [];
    for (let col = sourceRange.startCol; col < sourceRange.endCol; col++) {
      const cell = doc.getCell(sheetId, { row, col }) as { value: unknown; formula: string | null };
      const formula = typeof cell?.formula === "string" && cell.formula.trim() !== "" ? cell.formula : null;
      const value = (cell?.value ?? null) as any;
      const input = (formula ?? value) as any;
      outRow.push({ input, value });
    }
    sourceCells.push(outRow);
  }

  const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode, canWriteCell });
  if (edits.length === 0) return { editsApplied: 0 };

  doc.beginBatch({ label: "Fill" });
  try {
    for (const edit of edits) {
      doc.setCellInput(sheetId, { row: edit.row, col: edit.col }, edit.value);
    }
  } finally {
    doc.endBatch();
  }

  return { editsApplied: edits.length };
}
