export type CellValue = string | number | boolean | null;

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string;
  backgroundColor?: string;
}

export type RangeCoords = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export class TypedEventEmitter {
  on(event: string, listener: (payload: any) => void): () => void;
  off(event: string, listener: (payload: any) => void): void;
  emit(event: string, payload: any): void;
}

export class Range {
  readonly sheet: Sheet;
  readonly coords: RangeCoords;

  get address(): string;

  getValues(): CellValue[][];
  setValues(values: CellValue[][]): void;

  getValue(): CellValue;
  setValue(value: CellValue): void;

  setFormat(format: Partial<CellFormat> | null): void;
  getFormat(): CellFormat;
}

export class Sheet {
  readonly workbook: Workbook;
  readonly name: string;

  getRange(address: string): Range;

  setCellValue(address: string, value: CellValue): void;
  setRangeValues(address: string, values: CellValue[][]): void;

  getCellValue(row: number, col: number): CellValue;
  getCellFormat(row: number, col: number): CellFormat;
}

export type Selection = { sheetName: string; address: string };

export class Workbook {
  readonly events: TypedEventEmitter;

  addSheet(name: string): Sheet;
  getSheet(name: string): Sheet;

  getActiveSheet(): Sheet;
  setActiveSheet(name: string): void;

  getSelection(): Selection;
  setSelection(sheetName: string, address: string): void;

  snapshot(): Record<string, Record<string, { value: CellValue; format: Record<string, any> }>>;
}

export function columnLabelToIndex(label: string): number;
export function indexToColumnLabel(index: number): string;
export function parseCellAddress(a1: string): { row: number; col: number };
export function formatCellAddress(coord: { row: number; col: number }): string;
export function parseRangeAddress(a1: string): RangeCoords;
export function formatRangeAddress(range: RangeCoords): string;

export const FORMULA_API_DTS: string;

