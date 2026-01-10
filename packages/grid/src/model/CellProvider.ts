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
  /** Exclusive end row index. */
  endRow: number;
  startCol: number;
  /** Exclusive end column index. */
  endCol: number;
}

export type CellProviderUpdate =
  | { type: "cells"; range: CellRange }
  | { type: "invalidateAll" };

export interface CellProvider {
  getCell(row: number, col: number): CellData | null;
  prefetch?(range: CellRange): void;
  subscribe?(listener: (update: CellProviderUpdate) => void): () => void;
}
