import type { CellCoord, GridLimits, Range, UsedRangeProvider } from "./types";
import { clampCell, clampRange } from "./range";
import { DEFAULT_GRID_LIMITS } from "./selection";

export interface CurrentRegionOptions {
  /**
   * Hard cap on the number of cells visited in the flood fill. This exists purely
   * as a performance safety valve for extremely large sheets/regions.
   *
   * When exceeded, the function returns the sheet used range (bounded) rather
   * than attempting to finish the traversal.
   */
  maxVisitedCells?: number;
  /**
   * If the used-range window is small enough, we use a dense Uint8Array for
   * visited tracking (fast + memory efficient). Above this threshold we fall
   * back to a Set-based tracker to avoid allocating huge arrays.
   */
  maxDenseVisitedArea?: number;
}

/**
 * Compute Excel-like "current region" for Ctrl+Shift+* (Ctrl+Shift+8).
 *
 * The algorithm matches the spec for this repo:
 * - A cell is considered non-empty if it has a value or a formula (via `isCellEmpty`).
 * - We take the 4-neighborhood connected component of non-empty cells.
 * - Return the bounding rectangle of that component.
 * - The scan is bounded to the sheet used range to avoid traversing the full grid.
 * - If the active cell is empty and there is no non-empty neighbor, return just
 *   the active cell.
 */
export function computeCurrentRegionRange(
  activeCell: CellCoord,
  data: Pick<UsedRangeProvider, "getUsedRange" | "isCellEmpty">,
  limits: GridLimits = DEFAULT_GRID_LIMITS,
  options: CurrentRegionOptions = {}
): Range {
  const active = clampCell(activeCell, limits);

  const usedRaw = data.getUsedRange();
  if (!usedRaw) {
    return { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
  }

  const used = clampRange(usedRaw, limits);

  const cell = { row: 0, col: 0 };
  const isNonEmpty = (row: number, col: number): boolean => {
    cell.row = row;
    cell.col = col;
    return !data.isCellEmpty(cell);
  };

  const inUsedBounds = (row: number, col: number): boolean =>
    row >= used.startRow && row <= used.endRow && col >= used.startCol && col <= used.endCol;

  const pickSeed = (): CellCoord | null => {
    if (inUsedBounds(active.row, active.col) && isNonEmpty(active.row, active.col)) return active;

    // If the active cell is empty, Excel selects the region if the active cell is adjacent
    // to non-empty data. Use the first non-empty orthogonal neighbor as the seed.
    const neighbors: Array<[number, number]> = [
      [active.row - 1, active.col],
      [active.row + 1, active.col],
      [active.row, active.col - 1],
      [active.row, active.col + 1],
    ];

    for (const [row, col] of neighbors) {
      if (!inUsedBounds(row, col)) continue;
      if (!isNonEmpty(row, col)) continue;
      return { row, col };
    }

    return null;
  };

  const seed = pickSeed();
  if (!seed) {
    return { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
  }

  const height = used.endRow - used.startRow + 1;
  const width = used.endCol - used.startCol + 1;
  const area = height * width;

  const maxDenseVisitedArea = options.maxDenseVisitedArea ?? 5_000_000;
  const visitedDense = area > 0 && area <= maxDenseVisitedArea ? new Uint8Array(area) : null;
  const visitedSparse = visitedDense ? null : new Set<number>();

  const toId = (row: number, col: number): number => (row - used.startRow) * width + (col - used.startCol);

  const markVisited = (id: number): boolean => {
    if (visitedDense) {
      if (visitedDense[id]) return false;
      visitedDense[id] = 1;
      return true;
    }
    if (visitedSparse!.has(id)) return false;
    visitedSparse!.add(id);
    return true;
  };

  const queue: number[] = [];
  queue.push(toId(seed.row, seed.col));
  markVisited(queue[0]!);

  let visitedCount = 0;
  const maxVisitedCells = options.maxVisitedCells ?? 5_000_000;

  let minRow = seed.row;
  let maxRow = seed.row;
  let minCol = seed.col;
  let maxCol = seed.col;

  for (let qi = 0; qi < queue.length; qi += 1) {
    const id = queue[qi]!;
    visitedCount += 1;
    if (visitedCount > maxVisitedCells) {
      // Fail-safe: avoid pathological traversal costs. For very large regions,
      // selecting the whole used range is a reasonable approximation.
      return { ...used };
    }

    const rowOffset = Math.floor(id / width);
    const colOffset = id - rowOffset * width;
    const row = used.startRow + rowOffset;
    const col = used.startCol + colOffset;

    // Expand bounding box.
    if (row < minRow) minRow = row;
    if (row > maxRow) maxRow = row;
    if (col < minCol) minCol = col;
    if (col > maxCol) maxCol = col;

    // 4-neighborhood flood fill.
    const maybeEnqueue = (nr: number, nc: number) => {
      if (nr < used.startRow || nr > used.endRow || nc < used.startCol || nc > used.endCol) return;
      if (!isNonEmpty(nr, nc)) return;
      const nid = toId(nr, nc);
      if (!markVisited(nid)) return;
      queue.push(nid);
    };

    maybeEnqueue(row - 1, col);
    maybeEnqueue(row + 1, col);
    maybeEnqueue(row, col - 1);
    maybeEnqueue(row, col + 1);
  }

  return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
}

