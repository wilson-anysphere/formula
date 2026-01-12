import { computeFillEdits, type CellRange, type CellScalar, type FillMode, type FillSourceCell } from "@formula/fill-engine";

import type { DocumentController } from "../document/documentController.js";

export interface ApplyFillCommitOptions {
  document: DocumentController;
  sheetId: string;
  sourceRange: CellRange;
  targetRange: CellRange;
  mode: FillMode;
  /**
   * Optional hook to provide computed values for formula cells.
   *
   * This is only needed when `mode === "copy"` and the consumer wants to fill
   * formulas as values (requires the evaluated value, not the stored input).
   */
  getCellComputedValue?: (row: number, col: number) => CellScalar;
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
  const { document: doc, sheetId, sourceRange, targetRange, mode, canWriteCell, getCellComputedValue } = options;

  const height = Math.max(0, sourceRange.endRow - sourceRange.startRow);
  const width = Math.max(0, sourceRange.endCol - sourceRange.startCol);
  if (height === 0 || width === 0) return { editsApplied: 0 };

  const toScalar = (value: unknown): CellScalar => {
    if (value == null) return null;
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
    const maybeRich = value as { text?: unknown } | null;
    if (maybeRich && typeof maybeRich.text === "string") return maybeRich.text;
    return String(value);
  };

  const sourceCells: FillSourceCell[][] = [];
  for (let row = sourceRange.startRow; row < sourceRange.endRow; row++) {
    const outRow: FillSourceCell[] = [];
    for (let col = sourceRange.startCol; col < sourceRange.endCol; col++) {
      const cell = doc.getCell(sheetId, { row, col }) as { value: unknown; formula: string | null };
      const formula = typeof cell?.formula === "string" && cell.formula.trim() !== "" ? cell.formula : null;
      if (formula) {
        const computed = getCellComputedValue ? getCellComputedValue(row, col) : null;
        outRow.push({ input: formula, value: computed ?? null });
        continue;
      }
      const scalar = toScalar(cell?.value ?? null);
      outRow.push({ input: scalar, value: scalar });
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
