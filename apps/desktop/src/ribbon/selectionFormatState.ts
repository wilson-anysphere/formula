import type { DocumentController } from "../document/documentController.js";
import type { Range } from "../selection/types";

export type SelectionHorizontalAlign = "left" | "center" | "right" | "mixed";

export type SelectionNumberFormat = string | "mixed" | null;

export type SelectionFormatState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  wrapText: boolean;
  align: SelectionHorizontalAlign;
  numberFormat: SelectionNumberFormat;
};

type NormalizedRange = { startRow: number; endRow: number; startCol: number; endCol: number };

function normalizeRange(range: Range): NormalizedRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function safeCellCount(range: NormalizedRange): number {
  const rows = range.endRow - range.startRow + 1;
  const cols = range.endCol - range.startCol + 1;
  return rows * cols;
}

function sampleAxisIndices(start: number, end: number, maxSamples: number): number[] {
  const len = end - start + 1;
  if (len <= 0) return [];
  if (len <= maxSamples) {
    return Array.from({ length: len }, (_, i) => start + i);
  }
  if (maxSamples <= 1) return [start];

  const out: number[] = [];
  for (let i = 0; i < maxSamples; i++) {
    const idx = start + Math.floor((i * (len - 1)) / (maxSamples - 1));
    if (out[out.length - 1] !== idx) out.push(idx);
  }
  return out;
}

function* sampleRangeCells(range: NormalizedRange, maxCells: number): Generator<{ row: number; col: number }, void> {
  if (maxCells <= 0) return;

  const rows = range.endRow - range.startRow + 1;
  const cols = range.endCol - range.startCol + 1;
  if (rows <= 0 || cols <= 0) return;

  // Sampling strategy:
  // - For small ranges, enumerate every cell.
  // - For large ranges, sample a coarse grid across the full area.
  //   This intentionally does *not* try to be exhaustive; callers can treat
  //   results as best-effort.
  const total = rows * cols;
  if (total <= maxCells) {
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        yield { row, col };
      }
    }
    return;
  }

  const approxPerAxis = Math.max(2, Math.floor(Math.sqrt(maxCells)));
  const rowSamples = sampleAxisIndices(range.startRow, range.endRow, approxPerAxis);
  const colSamples = sampleAxisIndices(range.startCol, range.endCol, approxPerAxis);

  let emitted = 0;
  for (const row of rowSamples) {
    for (const col of colSamples) {
      yield { row, col };
      emitted += 1;
      if (emitted >= maxCells) return;
    }
  }
}

type AggregationState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  wrapText: boolean;
  align: "left" | "center" | "right" | "mixed" | null;
  numberFormat: string | null | "mixed" | undefined;
  inspected: number;
  exhaustive: boolean;
};

/**
 * Compute a lightweight formatting state summary for the current selection.
 *
 * The intent is to drive Ribbon toggle state (Bold/Italic/Underline/etc.) without
 * iterating every cell in very large selections.
 *
 * Performance semantics:
 * - If the total selection cell count is <= `maxInspectCells`, all cells are inspected.
 * - Otherwise, we inspect up to `maxInspectCells` sampled cells across the selection.
 * - For properties that require certainty, callers can treat sampled results as "mixed".
 */
export function computeSelectionFormatState(
  doc: DocumentController,
  sheetId: string,
  selectionRanges: Range[],
  options: { maxInspectCells?: number } = {},
): SelectionFormatState {
  const maxInspectCells = options.maxInspectCells ?? 256;

  const ranges = selectionRanges.map(normalizeRange).filter((r) => safeCellCount(r) > 0);
  if (ranges.length === 0) {
    return {
      bold: false,
      italic: false,
      underline: false,
      wrapText: false,
      align: "left",
      numberFormat: null,
    };
  }

  // Determine whether we can inspect the selection exhaustively.
  let totalCells = 0;
  for (const r of ranges) {
    totalCells += safeCellCount(r);
    if (totalCells > maxInspectCells) break;
  }
  const exhaustive = totalCells <= maxInspectCells;

  /** @type {AggregationState} */
  const state: AggregationState = {
    bold: true,
    italic: true,
    underline: true,
    wrapText: true,
    align: null,
    numberFormat: undefined,
    inspected: 0,
    exhaustive,
  };

  const visited = new Set<string>();

  const mergeAlign = (raw: unknown) => {
    const value = raw === "center" || raw === "right" || raw === "left" ? (raw as "left" | "center" | "right") : "left";
    if (state.align == null) state.align = value;
    else if (state.align !== "mixed" && state.align !== value) state.align = "mixed";
  };

  const mergeNumberFormat = (raw: unknown) => {
    const value = typeof raw === "string" ? raw : null;
    if (state.numberFormat === undefined) state.numberFormat = value;
    else if (state.numberFormat !== "mixed" && state.numberFormat !== value) state.numberFormat = "mixed";
  };

  const inspectCell = (row: number, col: number) => {
    const key = `${row},${col}`;
    if (visited.has(key)) return;
    visited.add(key);
    state.inspected += 1;

    const cell = doc.getCell(sheetId, { row, col }) as any;
    const style = doc.styleTable.get(cell?.styleId ?? 0) as any;

    state.bold = state.bold && Boolean(style?.font?.bold);
    state.italic = state.italic && Boolean(style?.font?.italic);
    state.underline = state.underline && Boolean(style?.font?.underline);
    state.wrapText = state.wrapText && Boolean(style?.alignment?.wrapText);

    mergeAlign(style?.alignment?.horizontal);
    mergeNumberFormat(style?.numberFormat);
  };

  outer: for (const range of ranges) {
    if (exhaustive) {
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          inspectCell(row, col);
          if (state.inspected >= maxInspectCells) break outer;
        }
      }
    } else {
      const remaining = maxInspectCells - state.inspected;
      for (const cell of sampleRangeCells(range, remaining)) {
        inspectCell(cell.row, cell.col);
        if (state.inspected >= maxInspectCells) break outer;
      }
    }
  }

  // If we didn't inspect the full selection, treat "all cells match" properties as mixed.
  const isCertain = state.exhaustive;

  return {
    bold: isCertain ? state.bold : false,
    italic: isCertain ? state.italic : false,
    underline: isCertain ? state.underline : false,
    wrapText: isCertain ? state.wrapText : false,
    align: (() => {
      if (state.align === "mixed") return "mixed";
      if (!isCertain) return "mixed";
      return state.align ?? "left";
    })(),
    numberFormat: (() => {
      if (state.numberFormat === "mixed") return "mixed";
      if (!isCertain) return "mixed";
      return state.numberFormat ?? null;
    })(),
  };
}

