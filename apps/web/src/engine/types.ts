export type EngineCellScalar = number | string | boolean | null;

export type EngineSheetJson = {
  cells: Record<string, EngineCellScalar>;
  /**
   * Optional logical worksheet dimensions (row/column count).
   *
   * When set, this controls how whole-column/row references like `A:A` / `1:1`
   * are expanded by the WASM engine.
   */
  rowCount?: number;
  colCount?: number;
};

export type EngineWorkbookJson = {
  sheets: Record<string, EngineSheetJson>;
};
