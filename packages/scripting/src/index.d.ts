export type CellValue = string | number | boolean | null;
export type CellFormula = string | null;

export interface CellFormat {
  [key: string]: any;
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string | null;
  backgroundColor?: string | null;
}

export type RangeCoords = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export interface RangeLike {
  readonly address: string;
  getValues(): CellValue[][];
  setValues(values: CellValue[][]): void;
  getValue(): CellValue;
  setValue(value: CellValue): void;
  setFormat(format: Partial<CellFormat> | null): void;
  getFormat(): CellFormat;
}

export interface SheetLike {
  readonly name: string;
  getRange(address: string): RangeLike;
}

export class TypedEventEmitter {
  on(event: string, listener: (payload: any) => void): () => void;
  off(event: string, listener: (payload: any) => void): void;
  emit(event: string, payload: any): void;
}

export class Range implements RangeLike {
  readonly sheet: Sheet;
  readonly coords: RangeCoords;

  get address(): string;

  getValues(): CellValue[][];
  setValues(values: CellValue[][]): void;

  getFormulas(): CellFormula[][];
  setFormulas(formulas: CellFormula[][]): void;

  getFormats(): CellFormat[][];
  setFormats(formats: Partial<CellFormat>[][]): void;

  getValue(): CellValue;
  setValue(value: CellValue): void;

  getFormat(): CellFormat;
  setFormat(format: Partial<CellFormat> | null): void;
}

export class Sheet implements SheetLike {
  readonly workbook: Workbook;
  readonly name: string;

  getRange(address: string): Range;
  getCell(row: number, col: number): Range;
  getUsedRange(): Range;

  setCellValue(address: string, value: CellValue): void;
  setRangeValues(address: string, values: CellValue[][]): void;
  setCellFormula(address: string, formula: CellFormula): void;

  getCellValue(row: number, col: number): CellValue;
  getCellFormula(row: number, col: number): CellFormula;
  getCellFormat(row: number, col: number): CellFormat;
}

export type Selection = { sheetName: string; address: string };

export interface WorkbookLike {
  getSheet(name: string): SheetLike;
  getActiveSheet(): SheetLike;
  getSelection(): Selection;
  setSelection(sheetName: string, address: string): void;
}

export class Workbook implements WorkbookLike {
  readonly events: TypedEventEmitter;

  addSheet(name: string): Sheet;
  getSheet(name: string): Sheet;
  getSheets(): Sheet[];

  getActiveSheet(): Sheet;
  setActiveSheet(name: string): void;

  getSelection(): Selection;
  setSelection(sheetName: string, address: string): void;

  snapshot(): Record<string, Record<string, { value: CellValue; formula: CellFormula; format: Record<string, any> }>>;
}

export function columnLabelToIndex(label: string): number;
export function indexToColumnLabel(index: number): string;
export function parseCellAddress(a1: string): { row: number; col: number };
export function formatCellAddress(coord: { row: number; col: number }): string;
export function parseRangeAddress(a1: string): RangeCoords;
export function formatRangeAddress(range: RangeCoords): string;

export const FORMULA_API_DTS: string;
