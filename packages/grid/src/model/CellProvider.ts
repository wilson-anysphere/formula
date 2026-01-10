export interface CellStyle {
  fill?: string;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string;
  textAlign?: CanvasTextAlign;
}

export interface CellData {
  row: number;
  col: number;
  value: string | number | null;
  style?: CellStyle;
}

export interface CellRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface CellProvider {
  getCell(row: number, col: number): CellData | null;
  prefetch?(range: CellRange): void;
}

