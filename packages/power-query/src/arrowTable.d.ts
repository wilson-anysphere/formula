import type { Column } from "./table.js";

export class ArrowTableAdapter {
  table: any;
  columns: Column[];
  rowCount: number;

  constructor(table: any, options?: any);

  getRow(rowIndex: number): unknown[];

  toGrid(options?: { includeHeader?: boolean }): unknown[][];
}
