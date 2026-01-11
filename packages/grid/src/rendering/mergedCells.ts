import type { CellRange } from "../model/CellProvider";

export interface CellRef {
  row: number;
  col: number;
}

function normalizeRange(range: CellRange): CellRange | null {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  if (startRow === endRow || startCol === endCol) return null;
  return { startRow, endRow, startCol, endCol };
}

/**
 * Efficient lookup structure for merged ranges.
 *
 * The index stores column spans per row for O(k) lookup where k is the number of
 * merged regions that touch a given row (typically small).
 */
export class MergedCellIndex {
  private readonly ranges: CellRange[];
  private readonly rowIndex: Map<number, Array<{ colStart: number; colEnd: number; rangeIndex: number }>>;

  constructor(ranges: CellRange[]) {
    this.ranges = [];
    for (const range of ranges) {
      const normalized = normalizeRange(range);
      if (normalized) this.ranges.push(normalized);
    }
    this.rowIndex = buildRowIndex(this.ranges);
  }

  getRanges(): readonly CellRange[] {
    return this.ranges;
  }

  /**
   * Returns the merged range that contains `cell`, if any.
   */
  rangeAt(cell: CellRef): CellRange | null {
    const spans = this.rowIndex.get(cell.row);
    if (!spans) return null;
    for (const span of spans) {
      if (cell.col < span.colStart || cell.col >= span.colEnd) continue;
      const range = this.ranges[span.rangeIndex];
      if (!range) continue;
      if (
        cell.row >= range.startRow &&
        cell.row < range.endRow &&
        cell.col >= range.startCol &&
        cell.col < range.endCol
      ) {
        return range;
      }
    }
    return null;
  }

  /**
   * Top-left anchor for a merged range (or the cell itself if not merged).
   */
  resolveCell(cell: CellRef): CellRef {
    const range = this.rangeAt(cell);
    if (!range) return cell;
    return { row: range.startRow, col: range.startCol };
  }

  isAnchor(cell: CellRef): boolean {
    const range = this.rangeAt(cell);
    return !!range && range.startRow === cell.row && range.startCol === cell.col;
  }

  /**
   * Cells inside a merged range that are not the anchor are skipped for text rendering.
   */
  shouldSkipCell(cell: CellRef): boolean {
    const range = this.rangeAt(cell);
    if (!range) return false;
    return !(range.startRow === cell.row && range.startCol === cell.col);
  }
}

/**
 * Returns `true` if the vertical gridline between `col` and `col+1` at `row`
 * lies inside a merged region and should not be drawn.
 */
export function isInteriorVerticalGridline(index: MergedCellIndex, row: number, col: number): boolean {
  if (row < 0 || col < 0) return false;
  const left = index.rangeAt({ row, col });
  if (!left) return false;
  return col + 1 < left.endCol;
}

/**
 * Returns `true` if the horizontal gridline between `row` and `row+1` at `col`
 * lies inside a merged region and should not be drawn.
 */
export function isInteriorHorizontalGridline(index: MergedCellIndex, row: number, col: number): boolean {
  if (row < 0 || col < 0) return false;
  const top = index.rangeAt({ row, col });
  if (!top) return false;
  return row + 1 < top.endRow;
}

export function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;
}

function buildRowIndex(ranges: CellRange[]): Map<number, Array<{ colStart: number; colEnd: number; rangeIndex: number }>> {
  const index = new Map<number, Array<{ colStart: number; colEnd: number; rangeIndex: number }>>();
  ranges.forEach((range, rangeIndex) => {
    for (let row = range.startRow; row < range.endRow; row++) {
      const spans = index.get(row) ?? [];
      spans.push({ colStart: range.startCol, colEnd: range.endCol, rangeIndex });
      index.set(row, spans);
    }
  });
  return index;
}

