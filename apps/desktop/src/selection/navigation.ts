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
      return tab(state, mods, limits, data);
    case "Enter":
      return enter(state, mods, limits, data);
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
  const target = mods.primary ? jumpToEdge(state.active, direction, data, limits) : moveByOne(state.active, direction, limits, data);

  if (mods.shift) {
    return extendSelectionToCell(state, target, limits);
  }

  return setActiveCell(state, target, limits);
}

function moveByOne(cell: CellCoord, direction: Direction, limits: GridLimits, data?: UsedRangeProvider): CellCoord {
  const delta = directionToDelta(direction);
  const clamped = clampCell({ row: cell.row + delta.dRow, col: cell.col + delta.dCol }, limits);

  if (!data) return clamped;

  if (delta.dRow !== 0) {
    const nextRow = nextVisibleRow(clamped.row, delta.dRow, data, limits);
    return { row: nextRow, col: clamped.col };
  }

  const nextCol = nextVisibleCol(clamped.col, delta.dCol, data, limits);
  return { row: clamped.row, col: nextCol };
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
  const start = clampCell(cell, limits);
  const minRow = Math.max(0, Math.min(used.startRow, start.row));
  const maxRow = Math.min(limits.maxRows - 1, Math.max(used.endRow, start.row));
  const minCol = Math.max(0, Math.min(used.startCol, start.col));
  const maxCol = Math.min(limits.maxCols - 1, Math.max(used.endCol, start.col));

  const isEmpty = (row: number, col: number) => data.isCellEmpty({ row, col });
  const rowHidden = (row: number) => data.isRowHidden?.(row) ?? false;
  const colHidden = (col: number) => data.isColHidden?.(col) ?? false;

  const nextRow = (row: number, dir: number): number | null => {
    let r = row + dir;
    while (r >= minRow && r <= maxRow && rowHidden(r)) {
      r += dir;
    }
    if (r < minRow || r > maxRow) return null;
    return r;
  };

  const nextCol = (col: number, dir: number): number | null => {
    let c = col + dir;
    while (c >= minCol && c <= maxCol && colHidden(c)) {
      c += dir;
    }
    if (c < minCol || c > maxCol) return null;
    return c;
  };

  // Use the current cell as the scan origin, but bound the scan window by the sheet's used
  // range. This prevents Ctrl+Arrow from traversing an "infinite" sheet while still behaving
  // sensibly when the active cell is outside the used range.
  let row = start.row;
  let col = start.col;

  if (dRow !== 0) {
    const limit = dRow > 0 ? maxRow : minRow;
    if (row === limit) return clampCell({ row, col }, limits);

    // Excel behavior depends on the *next* cell in the direction rather than the current cell.
    let cursor = nextRow(row, dRow);
    if (cursor === null) return clampCell({ row, col }, limits);

    if (isEmpty(cursor, col)) {
      // Skip empty cells.
      while (cursor !== null && isEmpty(cursor, col)) {
        row = cursor;
        if (row === limit) break;
        cursor = nextRow(row, dRow);
      }

      if (row === limit || cursor === null) {
        return clampCell({ row, col }, limits);
      }
    }

    // Step into the first non-empty cell.
    row = cursor;
    cursor = nextRow(row, dRow);

    // Then run to the end of the contiguous non-empty block.
    while (cursor !== null && !isEmpty(cursor, col)) {
      row = cursor;
      cursor = nextRow(row, dRow);
    }

    return clampCell({ row, col }, limits);
  }

  const limit = dCol > 0 ? maxCol : minCol;
  if (col === limit) return clampCell({ row, col }, limits);

  let cursor = nextCol(col, dCol);
  if (cursor === null) return clampCell({ row, col }, limits);

  if (isEmpty(row, cursor)) {
    while (cursor !== null && isEmpty(row, cursor)) {
      col = cursor;
      if (col === limit) break;
      cursor = nextCol(col, dCol);
    }

    if (col === limit || cursor === null) {
      return clampCell({ row, col }, limits);
    }
  }

  col = cursor;
  cursor = nextCol(col, dCol);
  while (cursor !== null && !isEmpty(row, cursor)) {
    col = cursor;
    cursor = nextCol(col, dCol);
  }

  return clampCell({ row, col }, limits);
}

function tab(state: SelectionState, mods: KeyModifiers, limits: GridLimits, data: UsedRangeProvider): SelectionState {
  const primaryRange = state.ranges[state.activeRangeIndex] ?? state.ranges[0];
  if (state.ranges.length === 1 && rangeArea(primaryRange) > 1 && cellInRange(state.active, primaryRange)) {
    const next = nextCellInRange(primaryRange, state.active, mods.shift ? "backward" : "forward");
    return buildSelection(
      { ranges: state.ranges, active: next, anchor: state.anchor, activeRangeIndex: state.activeRangeIndex },
      limits
    );
  }

  // Excel-style tab traversal for a single active cell:
  // - Tab moves right; at the last column, wrap to the first column of the next row.
  // - Shift+Tab moves left; at the first column, wrap to the last column of the previous row.
  //
  // This is distinct from multi-cell selection traversal (handled above), where Tab cycles
  // within the selection bounds.
  const active = clampCell(state.active, limits);

  if (mods.shift) {
    if (active.col > 0) {
      const target = moveByOne(active, "left", limits, data);
      return setActiveCell(state, target, limits);
    }

    // At the start of the row: wrap to the previous row's last column (if any).
    if (active.row <= 0) return state;
    const prevRow = nextVisibleRow(active.row - 1, -1, data, limits);
    const lastCol = nextVisibleCol(limits.maxCols - 1, -1, data, limits);
    return setActiveCell(state, { row: prevRow, col: lastCol }, limits);
  }

  if (active.col < limits.maxCols - 1) {
    const target = moveByOne(active, "right", limits, data);
    return setActiveCell(state, target, limits);
  }

  // At the end of the row: wrap to the next row's first column (if any).
  if (active.row >= limits.maxRows - 1) return state;
  const nextRow = nextVisibleRow(active.row + 1, 1, data, limits);
  const firstCol = nextVisibleCol(0, 1, data, limits);
  return setActiveCell(state, { row: nextRow, col: firstCol }, limits);
}

function enter(state: SelectionState, mods: KeyModifiers, limits: GridLimits, data: UsedRangeProvider): SelectionState {
  const primaryRange = state.ranges[state.activeRangeIndex] ?? state.ranges[0];
  if (state.ranges.length === 1 && rangeArea(primaryRange) > 1 && cellInRange(state.active, primaryRange)) {
    const next = nextCellInRangeForEnter(primaryRange, state.active, mods.shift ? "backward" : "forward");
    return buildSelection(
      { ranges: state.ranges, active: next, anchor: state.anchor, activeRangeIndex: state.activeRangeIndex },
      limits
    );
  }

  const target = moveByOne(state.active, mods.shift ? "up" : "down", limits, data);
  return setActiveCell(state, target, limits);
}

function nextVisibleRow(row: number, dir: number, data: UsedRangeProvider, limits: GridLimits): number {
  if (!data.isRowHidden) return row;
  let r = row;
  while (r >= 0 && r < limits.maxRows && data.isRowHidden(r)) {
    r += dir;
  }
  // If we ran off the sheet bounds without finding a visible row, keep the selection anchored
  // on the last in-bounds row in the opposite direction (Excel-like: Arrow keys should not land
  // on hidden boundary rows).
  if (r < 0 || r >= limits.maxRows) {
    const fallback = r < 0 ? row - dir : row - dir;
    if (fallback >= 0 && fallback < limits.maxRows) return fallback;
    return Math.min(limits.maxRows - 1, Math.max(0, row));
  }
  return r;
}

function nextVisibleCol(col: number, dir: number, data: UsedRangeProvider, limits: GridLimits): number {
  if (!data.isColHidden) return col;
  let c = col;
  while (c >= 0 && c < limits.maxCols && data.isColHidden(c)) {
    c += dir;
  }
  if (c < 0 || c >= limits.maxCols) {
    const fallback = c < 0 ? col - dir : col - dir;
    if (fallback >= 0 && fallback < limits.maxCols) return fallback;
    return Math.min(limits.maxCols - 1, Math.max(0, col));
  }
  return c;
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
