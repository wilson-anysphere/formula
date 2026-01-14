/**
 * Shared structural types (documentation-only; runtime is plain JS).
 *
 * These are intentionally loose to match the runtime's tolerant normalization.
 */

export type Cell = {
  /** Cell value */
  v?: any;
  /** Formula string, e.g. "=SUM(A1:A10)" */
  f?: string | null | undefined;
};

export type Sheet = {
  name: string;
  /**
   * Either a 2D matrix `[row][col]` or a sparse Map keyed by `"row,col"` or `"row:col"`.
   */
  cells: Cell[][] | Map<string, any>;
  /** Optional alternative to `cells` (treated as `[row][col]`). */
  values?: any[][];
};

export type WorkbookTable = {
  name: string;
  sheetName: string;
  /** inclusive 0-based */
  rect: { r0: number; c0: number; r1: number; c1: number };
};

export type NamedRange = {
  name: string;
  sheetName: string;
  /** inclusive 0-based */
  rect: { r0: number; c0: number; r1: number; c1: number };
};

export type Workbook = {
  id: string;
  sheets: Sheet[];
  tables?: WorkbookTable[];
  namedRanges?: NamedRange[];
};

export type WorkbookChunk = {
  id: string;
  workbookId: string;
  sheetName: string;
  kind: "table" | "namedRange" | "dataRegion" | "formulaRegion";
  title: string;
  rect: { r0: number; c0: number; r1: number; c1: number };
  /**
   * Sampled window of cells (bounded for embedding), aligned to `rect.r0/c0`.
   */
  cells: Cell[][];
  meta?: any;
};
