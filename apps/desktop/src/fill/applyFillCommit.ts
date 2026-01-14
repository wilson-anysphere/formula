import {
  computeFillEdits,
  type CellRange,
  type CellScalar,
  type FillEdit,
  type FillMode,
  type FillSourceCell,
} from "@formula/fill-engine";

import type { DocumentController } from "../document/documentController.js";
import { parseImageCellValue } from "../shared/imageCellValue.js";

export type RewriteFormulasForCopyDeltaRequest = { formula: string; deltaRow: number; deltaCol: number };

function normalizeFormulaText(formula: unknown): string | null {
  if (formula == null) return null;
  const trimmed = String(formula).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

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

  // Avoid allocating a fresh `{row,col}` object for every cell visit during fills.
  const coordScratch = { row: 0, col: 0 };

  const escapeLiteralTextForCellInput = (text: string): string => {
    if (text.trimStart().startsWith("=") || text.startsWith("'")) return `'${text}`;
    return text;
  };

  const toScalar = (value: unknown): CellScalar => {
    if (value == null) return null;
    if (typeof value === "string") return escapeLiteralTextForCellInput(value);
    if (typeof value === "number" || typeof value === "boolean") return value;
    const maybeRich = value as { text?: unknown } | null;
    if (maybeRich && typeof maybeRich.text === "string") return escapeLiteralTextForCellInput(maybeRich.text);

    const image = parseImageCellValue(value);
    if (image) return escapeLiteralTextForCellInput(image.altText ?? "[Image]");

    return String(value);
  };

  const sourceCells: FillSourceCell[][] = [];
  for (let row = sourceRange.startRow; row < sourceRange.endRow; row++) {
    const outRow: FillSourceCell[] = [];
    coordScratch.row = row;
    for (let col = sourceRange.startCol; col < sourceRange.endCol; col++) {
      coordScratch.col = col;
      const cell = doc.getCell(sheetId, coordScratch) as { value: unknown; formula: string | null };
      const formula = normalizeFormulaText(cell?.formula);
      if (formula) {
        let computed = getCellComputedValue ? getCellComputedValue(row, col) : null;
        if (mode === "copy" && typeof computed === "string") {
          computed = escapeLiteralTextForCellInput(computed);
        }
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
      coordScratch.row = edit.row;
      coordScratch.col = edit.col;
      doc.setCellInput(sheetId, coordScratch, edit.value);
    }
  } finally {
    doc.endBatch();
  }

  return { editsApplied: edits.length };
}

function mod(value: number, modulus: number): number {
  const rem = value % modulus;
  return rem < 0 ? rem + modulus : rem;
}

function isFormulaInput(value: CellScalar): value is string {
  if (typeof value !== "string") return false;
  // Match the engine's formula detection semantics: allow leading whitespace before '='
  // but require at least one non-whitespace character after it.
  const trimmed = value.trimStart();
  if (!trimmed.startsWith("=")) return false;
  return trimmed.slice(1).trim() !== "";
}

type FillAxis = "vertical" | "horizontal";

function detectFillAxis(sourceRange: CellRange, targetRange: CellRange): FillAxis {
  const sameCols = targetRange.startCol === sourceRange.startCol && targetRange.endCol === sourceRange.endCol;
  const sameRows = targetRange.startRow === sourceRange.startRow && targetRange.endRow === sourceRange.endRow;

  const targetOutsideRows = targetRange.endRow <= sourceRange.startRow || targetRange.startRow >= sourceRange.endRow;
  const targetOutsideCols = targetRange.endCol <= sourceRange.startCol || targetRange.startCol >= sourceRange.endCol;

  if (sameCols && targetOutsideRows) return "vertical";
  if (sameRows && targetOutsideCols) return "horizontal";

  throw new Error(`Unsupported fill axis: source=${JSON.stringify(sourceRange)} target=${JSON.stringify(targetRange)}`);
}

export async function applyFillCommitToDocumentControllerWithFormulaRewrite(
  options: ApplyFillCommitOptions & {
    rewriteFormulasForCopyDelta: (requests: RewriteFormulasForCopyDeltaRequest[]) => Promise<string[]>;
    /**
     * Optional batch label for undo history (defaults to "Fill").
     */
    label?: string;
  }
): Promise<{ editsApplied: number }> {
  const { document: doc, sheetId } = options;

  const edits = await computeFillEditsForDocumentControllerWithFormulaRewrite(options);
  if (edits.length === 0) return { editsApplied: 0 };

  // Avoid allocating a fresh `{row,col}` object for every cell visit during fills.
  const coordScratch = { row: 0, col: 0 };

  doc.beginBatch({ label: options.label ?? "Fill" });
  try {
    for (const edit of edits) {
      coordScratch.row = edit.row;
      coordScratch.col = edit.col;
      doc.setCellInput(sheetId, coordScratch, edit.value);
    }
  } finally {
    doc.endBatch();
  }

  return { editsApplied: edits.length };
}

export async function computeFillEditsForDocumentControllerWithFormulaRewrite(
  options: ApplyFillCommitOptions & {
    rewriteFormulasForCopyDelta: (requests: RewriteFormulasForCopyDeltaRequest[]) => Promise<string[]>;
  }
): Promise<FillEdit[]> {
  const { document: doc, sheetId, sourceRange, targetRange, mode, canWriteCell, getCellComputedValue } = options;

  const height = Math.max(0, sourceRange.endRow - sourceRange.startRow);
  const width = Math.max(0, sourceRange.endCol - sourceRange.startCol);
  if (height === 0 || width === 0) return [];

  // Avoid allocating a fresh `{row,col}` object for every cell visit during fills.
  const coordScratch = { row: 0, col: 0 };

  const escapeLiteralTextForCellInput = (text: string): string => {
    if (text.trimStart().startsWith("=") || text.startsWith("'")) return `'${text}`;
    return text;
  };

  const toScalar = (value: unknown): CellScalar => {
    if (value == null) return null;
    if (typeof value === "string") return escapeLiteralTextForCellInput(value);
    if (typeof value === "number" || typeof value === "boolean") return value;
    const maybeRich = value as { text?: unknown } | null;
    if (maybeRich && typeof maybeRich.text === "string") return escapeLiteralTextForCellInput(maybeRich.text);

    const image = parseImageCellValue(value);
    if (image) return escapeLiteralTextForCellInput(image.altText ?? "[Image]");

    return String(value);
  };

  const sourceCells: FillSourceCell[][] = [];
  for (let row = sourceRange.startRow; row < sourceRange.endRow; row++) {
    const outRow: FillSourceCell[] = [];
    coordScratch.row = row;
    for (let col = sourceRange.startCol; col < sourceRange.endCol; col++) {
      coordScratch.col = col;
      const cell = doc.getCell(sheetId, coordScratch) as { value: unknown; formula: string | null };
      const formula = normalizeFormulaText(cell?.formula);
      if (formula) {
        let computed = getCellComputedValue ? getCellComputedValue(row, col) : null;
        if (mode === "copy" && typeof computed === "string") {
          computed = escapeLiteralTextForCellInput(computed);
        }
        outRow.push({ input: formula, value: computed ?? null });
        continue;
      }
      const scalar = toScalar(cell?.value ?? null);
      outRow.push({ input: scalar, value: scalar });
    }
    sourceCells.push(outRow);
  }

  const { edits } = computeFillEdits({ sourceRange, targetRange, sourceCells, mode, canWriteCell });
  if (edits.length === 0) return [];

  // Only rewrite formula fills (not "fill formulas as values" copy mode).
  if (mode === "copy") {
    return edits;
  }

  // Patch formula edits using engine rewrite semantics.
  //
  // `computeFillEdits` uses a best-effort A1 shifter; we still call it so we can reuse its
  // series/copy behavior, but we override any formula fills using the engine's AST-based rewrite.
  try {
    const axis = detectFillAxis(sourceRange, targetRange);
    const requests: RewriteFormulasForCopyDeltaRequest[] = [];
    const editIndexByRequestIndex: number[] = [];

    for (let i = 0; i < edits.length; i++) {
      const edit = edits[i]!;
      const row = edit.row;
      const col = edit.col;

      let sourceRow = row;
      let sourceCol = col;

      if (axis === "vertical") {
        sourceRow = sourceRange.startRow + mod(row - sourceRange.startRow, height);
        sourceCol = col;
      } else {
        sourceRow = row;
        sourceCol = sourceRange.startCol + mod(col - sourceRange.startCol, width);
      }

      const source = sourceCells[sourceRow - sourceRange.startRow]?.[sourceCol - sourceRange.startCol];
      if (!source) continue;
      if (!isFormulaInput(source.input)) continue;

      requests.push({ formula: source.input, deltaRow: row - sourceRow, deltaCol: col - sourceCol });
      editIndexByRequestIndex.push(i);
    }

    if (requests.length > 0) {
      const rewritten = await options.rewriteFormulasForCopyDelta(requests);
      if (Array.isArray(rewritten) && rewritten.length === requests.length) {
        for (let i = 0; i < rewritten.length; i++) {
          const editIndex = editIndexByRequestIndex[i];
          const next = rewritten[i];
          if (typeof editIndex !== "number") continue;
          if (typeof next !== "string") continue;
          const edit = edits[editIndex];
          if (edit) edit.value = next;
        }
      }
    }
  } catch {
    // Ignore rewrite failures and fall back to the best-effort fill-engine result.
  }

  return edits;
}
