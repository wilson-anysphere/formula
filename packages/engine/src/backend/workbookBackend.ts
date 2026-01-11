export type SheetInfo = {
  id: string;
  name: string;
};

export type WorkbookInfo = {
  path: string | null;
  origin_path: string | null;
  sheets: SheetInfo[];
};

export type CellValue = {
  value: unknown | null;
  formula: string | null;
  display_value: string;
};

export type RangeData = {
  values: CellValue[][];
  start_row: number;
  start_col: number;
};

export type RangeCellEdit = {
  value: unknown | null;
  formula: string | null;
};

export type SheetUsedRange = {
  start_row: number;
  end_row: number;
  start_col: number;
  end_col: number;
};

/**
 * WorkbookBackend v1
 *
 * Minimal cross-platform workbook surface used by the UI. Desktop implementations
 * call into Tauri commands, while web implementations route through a WASM engine
 * running in a Worker.
 */
export interface WorkbookBackend {
  newWorkbook(): Promise<WorkbookInfo>;

  openWorkbook?(path: string): Promise<WorkbookInfo>;
  openWorkbookFromBytes?(bytes: Uint8Array): Promise<WorkbookInfo>;

  getSheetUsedRange(sheetId: string): Promise<SheetUsedRange | null>;

  getRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  }): Promise<RangeData>;

  setCell(params: {
    sheetId: string;
    row: number;
    col: number;
    value: unknown | null;
    formula: string | null;
  }): Promise<void>;

  setRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    values: RangeCellEdit[][];
  }): Promise<void>;

  saveWorkbook?(path?: string): Promise<void>;
}

