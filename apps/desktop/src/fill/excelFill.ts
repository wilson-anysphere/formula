import type { DocumentController } from "../document/documentController.js";
import type { Range } from "../selection/types";
import { applyFillCommitToDocumentController } from "./applyFillCommit";
import type { CellRange as FillEngineRange } from "@formula/fill-engine";

function normalizeSelectionRange(range: Range): Range {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

export function applyExcelFillDown(options: {
  document: DocumentController;
  sheetId: string;
  ranges: Range[];
  /**
   * Optional hook to supply computed values when fill mode is "copy". (Not needed for
   * the Ctrl+D/Ctrl+R "formulas" mode, but supported for future-proofing.)
   */
  getCellComputedValue?: (row: number, col: number) => string | number | boolean | null;
}): { editsApplied: number } {
  const { document: doc, sheetId, ranges, getCellComputedValue } = options;
  if (!ranges || ranges.length === 0) return { editsApplied: 0 };

  let editsApplied = 0;
  doc.beginBatch({ label: "Fill Down" });
  try {
    for (const rawRange of ranges) {
      const range = normalizeSelectionRange(rawRange);
      if (range.endRow <= range.startRow) continue;

      const sourceRange: FillEngineRange = {
        startRow: range.startRow,
        endRow: range.startRow + 1,
        startCol: range.startCol,
        endCol: range.endCol + 1,
      };
      const targetRange: FillEngineRange = {
        startRow: range.startRow + 1,
        endRow: range.endRow + 1,
        startCol: range.startCol,
        endCol: range.endCol + 1,
      };

      editsApplied += applyFillCommitToDocumentController({
        document: doc,
        sheetId,
        sourceRange,
        targetRange,
        mode: "formulas",
        ...(getCellComputedValue ? { getCellComputedValue } : {}),
      }).editsApplied;
    }
  } finally {
    doc.endBatch();
  }

  return { editsApplied };
}

export function applyExcelFillRight(options: {
  document: DocumentController;
  sheetId: string;
  ranges: Range[];
  getCellComputedValue?: (row: number, col: number) => string | number | boolean | null;
}): { editsApplied: number } {
  const { document: doc, sheetId, ranges, getCellComputedValue } = options;
  if (!ranges || ranges.length === 0) return { editsApplied: 0 };

  let editsApplied = 0;
  doc.beginBatch({ label: "Fill Right" });
  try {
    for (const rawRange of ranges) {
      const range = normalizeSelectionRange(rawRange);
      if (range.endCol <= range.startCol) continue;

      const sourceRange: FillEngineRange = {
        startRow: range.startRow,
        endRow: range.endRow + 1,
        startCol: range.startCol,
        endCol: range.startCol + 1,
      };
      const targetRange: FillEngineRange = {
        startRow: range.startRow,
        endRow: range.endRow + 1,
        startCol: range.startCol + 1,
        endCol: range.endCol + 1,
      };

      editsApplied += applyFillCommitToDocumentController({
        document: doc,
        sheetId,
        sourceRange,
        targetRange,
        mode: "formulas",
        ...(getCellComputedValue ? { getCellComputedValue } : {}),
      }).editsApplied;
    }
  } finally {
    doc.endBatch();
  }

  return { editsApplied };
}

