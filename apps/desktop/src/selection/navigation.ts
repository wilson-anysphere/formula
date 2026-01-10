import type { CellCoord, GridLimits, SelectionState, UsedRangeProvider } from "./types";
import { clampCell, cellInRange, rangeArea } from "./range";
import { DEFAULT_GRID_LIMITS, buildSelection, extendSelectionToCell, setActiveCell } from "./selection";

export type Direction = "up" | "down" | "left" | "right";

export interface KeyModifiers {
  shift: boolean;
  /**
   * Primary modifier: Ctrl on Windows/Linux, Cmd (meta) on macOS.
   * We accept both to make tests portable.
   */
  primary: boolean;
}

export function navigateSelectionByKey(
  state: SelectionState,
  key: string,
  mods: KeyModifiers,
  data: UsedRangeProvider,
  limits: GridLimits = DEFAULT_GRID_LIMITS
): SelectionState | null {
  switch (key) {
    case "ArrowUp":
      return move(state, "up", mods, data, limits);
    case "ArrowDown":
      return move(state, "down", mods, data, limits);
    case "ArrowLeft":
      return move(state, "left", mods, data, limits);
    case "ArrowRight":
      return move(state, "right", mods, data, limits);
    case "Tab":
      return tab(state, mods, limits);
    case "Enter":
      return enter(state, mods, limits);
    case "Home":
      if (mods.primary) return mods.shift ? extendSelectionToCell(state, { row: 0, col: 0 }, limits) : setActiveCell(state, { row: 0, col: 0 }, limits);
      return null;
    case "End":
      if (mods.primary) {
        const used = data.getUsedRange();
        const target = used ? { row: used.endRow, col: used.endCol } : { row: 0, col: 0 };
        return mods.shift ? extendSelectionToCell(state, target, limits) : setActiveCell(state, target, limits);
      }
      return null;
    default:
      return null;
  }
}

function move(
  state: SelectionState,
  direction: Direction,
  mods: KeyModifiers,
  data: UsedRangeProvider,
  limits: GridLimits
): SelectionState {
  const target = mods.primary ? jumpToEdge(state.active, direction, data, limits) : moveByOne(state.active, direction, limits);

  if (mods.shift) {
    return extendSelectionToCell(state, target, limits);
  }

  return setActiveCell(state, target, limits);
}

function moveByOne(cell: CellCoord, direction: Direction, limits: GridLimits): CellCoord {
  const delta = directionToDelta(direction);
  return clampCell({ row: cell.row + delta.dRow, col: cell.col + delta.dCol }, limits);
}

function directionToDelta(direction: Direction): { dRow: number; dCol: number } {
  switch (direction) {
    case "up":
      return { dRow: -1, dCol: 0 };
    case "down":
      return { dRow: 1, dCol: 0 };
    case "left":
      return { dRow: 0, dCol: -1 };
    case "right":
      return { dRow: 0, dCol: 1 };
  }
}

/**
 * Excel-like Ctrl+Arrow behavior:
 * - If current cell is non-empty, jump to the last non-empty cell before the first empty cell.
 * - If current cell is empty, jump to the next non-empty cell and then to the last contiguous non-empty.
 *
 * We bound the scan by the sheet used range to avoid traversing an "infinite" sheet.
 */
export function jumpToEdge(cell: CellCoord, direction: Direction, data: UsedRangeProvider, limits: GridLimits): CellCoord {
  const used = data.getUsedRange();
  if (!used) return cell;

  const { dRow, dCol } = directionToDelta(direction);
  const minRow = used.startRow;
  const maxRow = used.endRow;
  const minCol = used.startCol;
  const maxCol = used.endCol;

  const isEmpty = (row: number, col: number) => data.isCellEmpty({ row, col });

  // Clamp the starting point to the used range to guarantee termination.
  let row = Math.max(minRow, Math.min(maxRow, cell.row));
  let col = Math.max(minCol, Math.min(maxCol, cell.col));

  if (dRow !== 0) {
    const limit = dRow > 0 ? maxRow : minRow;
    if (row === limit) return clampCell({ row, col }, limits);

    // Excel behavior depends on the *next* cell in the direction rather than the current cell.
    // This ensures Ctrl+Arrow always moves somewhere (unless already at the boundary).
    if (isEmpty(row + dRow, col)) {
      // Skip empty cells.
      while (row !== limit && isEmpty(row + dRow, col)) {
        row += dRow;
      }

      if (row === limit) {
        return clampCell({ row: limit, col }, limits);
      }
    }

    // Step into the first non-empty cell (if any).
    row += dRow;

    // Then run to the end of the contiguous non-empty block.
    while (row !== limit && !isEmpty(row + dRow, col)) {
      row += dRow;
    }

    return clampCell({ row, col }, limits);
  }

  const limit = dCol > 0 ? maxCol : minCol;
  if (col === limit) return clampCell({ row, col }, limits);

  if (isEmpty(row, col + dCol)) {
    while (col !== limit && isEmpty(row, col + dCol)) {
      col += dCol;
    }

    if (col === limit) {
      return clampCell({ row, col: limit }, limits);
    }
  }

  col += dCol;
  while (col !== limit && !isEmpty(row, col + dCol)) {
    col += dCol;
  }

  return clampCell({ row, col }, limits);
}

function tab(state: SelectionState, mods: KeyModifiers, limits: GridLimits): SelectionState {
  const primaryRange = state.ranges[state.activeRangeIndex] ?? state.ranges[0];
  if (state.ranges.length === 1 && rangeArea(primaryRange) > 1 && cellInRange(state.active, primaryRange)) {
    const next = nextCellInRange(primaryRange, state.active, mods.shift ? "backward" : "forward");
    return buildSelection(
      { ranges: state.ranges, active: next, anchor: state.anchor, activeRangeIndex: state.activeRangeIndex },
      limits
    );
  }

  const target = moveByOne(state.active, mods.shift ? "left" : "right", limits);
  return setActiveCell(state, target, limits);
}

function enter(state: SelectionState, mods: KeyModifiers, limits: GridLimits): SelectionState {
  const primaryRange = state.ranges[state.activeRangeIndex] ?? state.ranges[0];
  if (state.ranges.length === 1 && rangeArea(primaryRange) > 1 && cellInRange(state.active, primaryRange)) {
    const next = nextCellInRangeForEnter(primaryRange, state.active, mods.shift ? "backward" : "forward");
    return buildSelection(
      { ranges: state.ranges, active: next, anchor: state.anchor, activeRangeIndex: state.activeRangeIndex },
      limits
    );
  }

  const target = moveByOne(state.active, mods.shift ? "up" : "down", limits);
  return setActiveCell(state, target, limits);
}

function nextCellInRange(
  range: { startRow: number; endRow: number; startCol: number; endCol: number },
  active: CellCoord,
  dir: "forward" | "backward"
): CellCoord {
  if (dir === "forward") {
    if (active.col < range.endCol) return { row: active.row, col: active.col + 1 };
    if (active.row < range.endRow) return { row: active.row + 1, col: range.startCol };
    return { row: range.startRow, col: range.startCol };
  }

  if (active.col > range.startCol) return { row: active.row, col: active.col - 1 };
  if (active.row > range.startRow) return { row: active.row - 1, col: range.endCol };
  return { row: range.endRow, col: range.endCol };
}

function nextCellInRangeForEnter(
  range: { startRow: number; endRow: number; startCol: number; endCol: number },
  active: CellCoord,
  dir: "forward" | "backward"
): CellCoord {
  if (dir === "forward") {
    if (active.row < range.endRow) return { row: active.row + 1, col: active.col };
    if (active.col < range.endCol) return { row: range.startRow, col: active.col + 1 };
    return { row: range.startRow, col: range.startCol };
  }

  if (active.row > range.startRow) return { row: active.row - 1, col: active.col };
  if (active.col > range.startCol) return { row: range.endRow, col: active.col - 1 };
  return { row: range.endRow, col: range.endCol };
}
