export type GridBatch = { rowOffset: number; values: any[][] };

export function parquetToArrowTable(parquetBytes: Uint8Array, options?: any): Promise<any>;

export function parquetFileToArrowTable(handle: Blob, options?: any): Promise<any>;

export function parquetFileToGridBatches(handle: Blob, options?: any): AsyncGenerator<GridBatch>;

export function arrowTableToParquet(table: any, options?: any): Promise<Uint8Array>;

export function arrowTableToIPC(table: any): Uint8Array;

export function arrowTableFromIPC(bytes: Uint8Array | ArrayBuffer): any;

export function arrowTableFromColumns(columns: Record<string, any[] | ArrayLike<any>>): any;

export function arrowTableToGridBatches(table: any, options?: { batchSize?: number; includeHeader?: boolean }): AsyncGenerator<GridBatch>;

export class ArrowColumnarSheet {
  table: any;
  columnNames: string[];

  constructor(table: any);

  get rowCount(): number;

  get columnCount(): number;

  getCell(row: number, col: number): unknown;

  slice(range: { startRow: number; endRow: number; startCol: number; endCol: number }): ArrowColumnarSheet;
}
