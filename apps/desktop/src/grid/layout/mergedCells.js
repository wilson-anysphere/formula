/**
 * @typedef {{ startRow: number; startCol: number; endRow: number; endCol: number }} CellRange
 * @typedef {{ row: number; col: number }} CellRef
 * @typedef {{ x: number; y: number; width: number; height: number }} Rect
 * @typedef {{
 *   getColWidth(col: number): number;
 *   getRowHeight(row: number): number;
 *   getColLeft(col: number): number;
 *   getRowTop(row: number): number;
 * }} GridMetrics
 */

/**
 * Efficient lookup structure for merged ranges.
 *
 * The index stores spans per row for O(k) lookup where k is the number of merged
 * regions that touch a given row (typically small).
 */
export class MergedRegionIndex {
  /** @type {CellRange[]} */
  #regions;
  /** @type {Map<number, Array<{ colStart: number; colEnd: number; regionIndex: number }>>} */
  #rowIndex;

  /**
   * @param {CellRange[]} regions
   */
  constructor(regions) {
    this.#regions = regions.map(normalizeRange);
    this.#rowIndex = buildRowIndex(this.#regions);
  }

  /** @returns {readonly CellRange[]} */
  getRegions() {
    return this.#regions;
  }

  /**
   * Returns the merged range that contains `cell`, if any.
   * @param {CellRef} cell
   * @returns {CellRange | undefined}
   */
  rangeAt(cell) {
    const spans = this.#rowIndex.get(cell.row);
    if (!spans) return undefined;
    for (const span of spans) {
      if (cell.col < span.colStart || cell.col > span.colEnd) continue;
      const region = this.#regions[span.regionIndex];
      if (
        cell.row >= region.startRow &&
        cell.row <= region.endRow &&
        cell.col >= region.startCol &&
        cell.col <= region.endCol
      ) {
        return region;
      }
    }
    return undefined;
  }

  /**
   * Top-left anchor for a merged range (or the cell itself if not merged).
   * @param {CellRef} cell
   * @returns {CellRef}
   */
  resolveCell(cell) {
    const range = this.rangeAt(cell);
    if (!range) return cell;
    return { row: range.startRow, col: range.startCol };
  }

  /**
   * @param {CellRef} cell
   * @returns {boolean}
   */
  isAnchor(cell) {
    const range = this.rangeAt(cell);
    return !!range && range.startRow === cell.row && range.startCol === cell.col;
  }

  /**
   * Cells inside a merged range that are not the anchor are skipped for text rendering.
   * @param {CellRef} cell
   * @returns {boolean}
   */
  shouldSkipCell(cell) {
    const range = this.rangeAt(cell);
    if (!range) return false;
    return !(range.startRow === cell.row && range.startCol === cell.col);
  }
}

/**
 * Returns `true` if the vertical gridline between `col` and `col+1` at `row`
 * lies inside a merged region and should not be drawn.
 * @param {MergedRegionIndex} index
 * @param {number} row
 * @param {number} col
 */
export function isInteriorVerticalGridline(index, row, col) {
  const left = index.rangeAt({ row, col });
  if (!left) return false;
  return col >= left.startCol && col < left.endCol;
}

/**
 * Returns `true` if the horizontal gridline between `row` and `row+1` at `col`
 * lies inside a merged region and should not be drawn.
 * @param {MergedRegionIndex} index
 * @param {number} row
 * @param {number} col
 */
export function isInteriorHorizontalGridline(index, row, col) {
  const top = index.rangeAt({ row, col });
  if (!top) return false;
  return row >= top.startRow && row < top.endRow;
}

/**
 * @param {CellRange} range
 * @param {GridMetrics} metrics
 * @returns {Rect}
 */
export function mergedRangeRect(range, metrics) {
  const x = metrics.getColLeft(range.startCol);
  const y = metrics.getRowTop(range.startRow);

  let width = 0;
  for (let c = range.startCol; c <= range.endCol; c++) width += metrics.getColWidth(c);

  let height = 0;
  for (let r = range.startRow; r <= range.endRow; r++) height += metrics.getRowHeight(r);

  return { x, y, width, height };
}

/**
 * @param {CellRange} range
 * @returns {CellRange}
 */
function normalizeRange(range) {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, startCol, endRow, endCol };
}

/**
 * @param {CellRange[]} regions
 */
function buildRowIndex(regions) {
  /** @type {Map<number, Array<{ colStart: number; colEnd: number; regionIndex: number }>>} */
  const index = new Map();
  regions.forEach((region, regionIndex) => {
    for (let row = region.startRow; row <= region.endRow; row++) {
      const spans = index.get(row) ?? [];
      spans.push({ colStart: region.startCol, colEnd: region.endCol, regionIndex });
      index.set(row, spans);
    }
  });
  return index;
}

