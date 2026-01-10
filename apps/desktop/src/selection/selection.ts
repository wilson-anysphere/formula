import type { CellCoord, GridLimits, Range, SelectionState, SelectionType } from "./types";
import { clampCell, clampRange, equalsRange, isSingleCellRange, rangeFromCells } from "./range";

export const DEFAULT_GRID_LIMITS: GridLimits = {
  // Excel limits as sane defaults.
  maxRows: 1_048_576,
  maxCols: 16_384
};

export function createSelection(cell: CellCoord = { row: 0, col: 0 }, limits: GridLimits = DEFAULT_GRID_LIMITS) {
  const clamped = clampCell(cell, limits);
  const range = { startRow: clamped.row, endRow: clamped.row, startCol: clamped.col, endCol: clamped.col };
  return buildSelection({ ranges: [range], active: clamped, anchor: clamped, activeRangeIndex: 0 }, limits);
}

export function setActiveCell(
  state: SelectionState,
  cell: CellCoord,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const clamped = clampCell(cell, limits);
  const range = { startRow: clamped.row, endRow: clamped.row, startCol: clamped.col, endCol: clamped.col };
  return buildSelection({ ranges: [range], active: clamped, anchor: clamped, activeRangeIndex: 0 }, limits);
}

export function extendSelectionToCell(
  state: SelectionState,
  cell: CellCoord,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const active = clampCell(cell, limits);
  const range = rangeFromCells(state.anchor, active);
  return buildSelection({ ranges: [range], active, anchor: state.anchor, activeRangeIndex: 0 }, limits);
}

export function addCellToSelection(
  state: SelectionState,
  cell: CellCoord,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const active = clampCell(cell, limits);
  const newRange: Range = { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
  const nextRanges = state.ranges.some((r) => equalsRange(r, newRange)) ? state.ranges : [...state.ranges, newRange];
  const activeRangeIndex = nextRanges.findIndex((r) => equalsRange(r, newRange));
  return buildSelection(
    {
      ranges: nextRanges,
      active,
      anchor: active,
      activeRangeIndex: Math.max(0, activeRangeIndex)
    },
    limits
  );
}

export function addRangeToSelection(
  state: SelectionState,
  range: Range,
  activeCell: CellCoord,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const clampedRange = clampRange(range, limits);
  const active = clampCell(activeCell, limits);
  const nextRanges = state.ranges.some((r) => equalsRange(r, clampedRange)) ? state.ranges : [...state.ranges, clampedRange];
  const activeRangeIndex = nextRanges.findIndex((r) => equalsRange(r, clampedRange));
  return buildSelection(
    {
      ranges: nextRanges,
      active,
      anchor: active,
      activeRangeIndex: Math.max(0, activeRangeIndex)
    },
    limits
  );
}

export function selectRows(
  state: SelectionState,
  startRow: number,
  endRow: number,
  opts: { additive?: boolean } = {},
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const clampedStart = Math.max(0, Math.min(limits.maxRows - 1, Math.trunc(startRow)));
  const clampedEnd = Math.max(0, Math.min(limits.maxRows - 1, Math.trunc(endRow)));
  const range: Range = clampRange(
    {
      startRow: clampedStart,
      endRow: clampedEnd,
      startCol: 0,
      endCol: limits.maxCols - 1
    },
    limits
  );

  const active: CellCoord = {
    row: range.startRow,
    col: clampCell(state.active, limits).col
  };

  if (opts.additive) {
    return addRangeToSelection(state, range, active, limits);
  }

  return buildSelection({ ranges: [range], active, anchor: active, activeRangeIndex: 0 }, limits);
}

export function selectColumns(
  state: SelectionState,
  startCol: number,
  endCol: number,
  opts: { additive?: boolean } = {},
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  const clampedStart = Math.max(0, Math.min(limits.maxCols - 1, Math.trunc(startCol)));
  const clampedEnd = Math.max(0, Math.min(limits.maxCols - 1, Math.trunc(endCol)));
  const range: Range = clampRange(
    {
      startRow: 0,
      endRow: limits.maxRows - 1,
      startCol: clampedStart,
      endCol: clampedEnd
    },
    limits
  );

  const active: CellCoord = {
    row: clampCell(state.active, limits).row,
    col: range.startCol
  };

  if (opts.additive) {
    return addRangeToSelection(state, range, active, limits);
  }

  return buildSelection({ ranges: [range], active, anchor: active, activeRangeIndex: 0 }, limits);
}

export function selectAll(limits: GridLimits = DEFAULT_GRID_LIMITS): SelectionState {
  const range: Range = { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: limits.maxCols - 1 };
  const active: CellCoord = { row: 0, col: 0 };
  return buildSelection({ ranges: [range], active, anchor: active, activeRangeIndex: 0 }, limits);
}

export function inferSelectionType(ranges: Range[], limits: GridLimits): SelectionType {
  if (ranges.length > 1) return "multi";

  const r = ranges[0];
  const isFullHeight = r.startRow === 0 && r.endRow === limits.maxRows - 1;
  const isFullWidth = r.startCol === 0 && r.endCol === limits.maxCols - 1;

  if (isFullHeight && isFullWidth) return "all";
  if (isFullHeight) return "column";
  if (isFullWidth) return "row";
  if (isSingleCellRange(r)) return "cell";
  return "range";
}

export function buildSelection(
  input: Pick<SelectionState, "ranges" | "active" | "anchor" | "activeRangeIndex">,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState {
  if (input.ranges.length === 0) {
    throw new Error("Selection must contain at least one range");
  }

  const ranges = input.ranges.map((r) => clampRange(r, limits));
  const active = clampCell(input.active, limits);
  const anchor = clampCell(input.anchor, limits);
  const activeRangeIndex = Math.max(0, Math.min(ranges.length - 1, Math.trunc(input.activeRangeIndex)));

  return {
    type: inferSelectionType(ranges, limits),
    ranges,
    active,
    anchor,
    activeRangeIndex
  };
}

