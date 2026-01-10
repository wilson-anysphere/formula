export type SelectionType = "cell" | "range" | "multi" | "column" | "row" | "all";

export interface CellCoord {
  row: number;
  col: number;
}

export interface Range {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface GridLimits {
  /**
   * Maximum number of rows in the grid (Excel default is 1,048,576).
   * We keep this configurable for testing and future "infinite" sheets.
   */
  maxRows: number;
  /**
   * Maximum number of columns in the grid (Excel default is 16,384).
   */
  maxCols: number;
}

export interface SelectionState {
  type: SelectionType;
  ranges: Range[];
  /**
   * The "active" cell is the focused cell: the one that receives keyboard navigation
   * and is the target for in-place editing (F2).
   */
  active: CellCoord;
  /**
   * Anchor is the fixed end of a shift-extended selection. When the user holds Shift,
   * navigation moves `active` while `anchor` remains stable.
   */
  anchor: CellCoord;
  /**
   * When multiple ranges exist, one range is considered "active" for Tab/Enter
   * traversal semantics.
   */
  activeRangeIndex: number;
}

export interface UsedRangeProvider {
  /**
   * Return the bounding box of all non-empty cells, or null if the sheet is empty.
   */
  getUsedRange(): Range | null;
  /**
   * Return true if the cell should be treated as empty for navigation purposes.
   */
  isCellEmpty(cell: CellCoord): boolean;
}

