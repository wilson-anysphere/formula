export interface CellStyle {
  fill?: string;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string;
  textAlign?: CanvasTextAlign;
  /**
   * Wrapping strategy for cell text.
   *
   * - `"none"`: single line (clip overflow)
   * - `"word"`: wrap at whitespace with char-wrap fallback
   * - `"char"`: wrap at grapheme cluster boundaries
   */
  wrapMode?: "none" | "word" | "char";
  /**
   * Explicit base direction for text rendering. `"auto"` uses the first strong
   * directional character to choose LTR/RTL.
   */
  direction?: "ltr" | "rtl" | "auto";
  verticalAlign?: "top" | "middle" | "bottom";
  /** Basic rotation support (clockwise, degrees). */
  rotationDeg?: number;
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
