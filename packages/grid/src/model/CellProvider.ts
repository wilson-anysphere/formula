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

export interface CellRichTextRun {
  /** Unicode code point start index (inclusive). */
  start: number;
  /** Unicode code point end index (exclusive). */
  end: number;
  style?: Record<string, unknown>;
}

export interface CellRichText {
  text: string;
  runs?: CellRichTextRun[];
}

export interface CellData {
  row: number;
  col: number;
  value: string | number | boolean | null;
  /**
   * Optional rich text payload for cell rendering.
   *
   * When present, `value` should typically be set to `richText.text` as a plain-text
   * fallback for accessibility / status strings.
   */
  richText?: CellRichText;
  style?: CellStyle;
  /**
   * Optional comment metadata used for rendering a cell comment indicator.
   *
   * This is intentionally minimal: higher layers can store full comment threads
   * elsewhere (e.g. in a Yjs doc) and only expose presence/resolved state here
   * for fast grid rendering.
   */
  comment?: { resolved?: boolean } | null;
}

export interface CellRange {
  startRow: number;
  /** Exclusive end row index. */
  endRow: number;
  startCol: number;
  /** Exclusive end column index. */
  endCol: number;
}

/**
 * A rectangular region of merged cells.
 *
 * The range uses the same exclusive-end semantics as {@link CellRange}. The
 * merged cell "anchor" is always the top-left cell at `startRow/startCol`.
 */
export type MergedCellRange = CellRange;

export type CellProviderUpdate =
  | { type: "cells"; range: CellRange }
  | { type: "invalidateAll" };

export interface CellProvider {
  getCell(row: number, col: number): CellData | null;
  /**
   * Returns the merged range that contains the given cell, if any.
   *
   * The returned range must use exclusive end coordinates (endRow/endCol), and
   * the anchor is assumed to be the top-left cell (startRow/startCol).
   */
  getMergedRangeAt?(row: number, col: number): MergedCellRange | null;
  /**
   * Returns merged ranges that intersect `range`.
   *
   * This is an optional bulk API used to build efficient per-viewport indexes.
   * Implementations should include merged ranges whose anchor is outside the
   * input range as long as they intersect it.
   */
  getMergedRangesInRange?(range: CellRange): MergedCellRange[];
  prefetch?(range: CellRange): void;
  subscribe?(listener: (update: CellProviderUpdate) => void): () => void;
}
