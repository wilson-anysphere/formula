export type DataType = string;

export type Column = {
  name: string;
  type: DataType;
};

export function makeUniqueColumnNames(rawNames: string[]): string[];
export function inferColumnType(values: unknown[]): DataType;

export class DataTable {
  columns: Column[];
  rowCount: number;

  constructor(columns: Column[], rows: unknown[][]);

  static fromGrid(grid: unknown[][], options?: { hasHeaders?: boolean; inferTypes?: boolean }): DataTable;

  getRow(rowIndex: number): unknown[];

  toGrid(options?: { includeHeader?: boolean }): unknown[][];
}
