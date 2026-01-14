export type CellBorderLineStyle = "solid" | "dashed" | "dotted" | "double";

export interface CellBorderSpec {
  /**
   * Border width in CSS pixels at zoom=1 (the renderer scales this by the current zoom).
   */
  width: number;
  style: CellBorderLineStyle;
  color: string;
}

export interface CellBorders {
  top?: CellBorderSpec;
  right?: CellBorderSpec;
  bottom?: CellBorderSpec;
  left?: CellBorderSpec;
}

export interface CellDiagonalBorders {
  /** Bottom-left → top-right. */
  up?: CellBorderSpec;
  /** Top-left → bottom-right. */
  down?: CellBorderSpec;
}

export interface CellStyle {
  fill?: string;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string;
  /**
   * Canvas font style (e.g. `"normal"`, `"italic"`).
   *
   * Note: This maps directly into the `FontSpec.style` field used by `@formula/text-layout`.
  */
  fontStyle?: string;
  /**
   * Excel-style baseline shift for subscript/superscript rendering.
   *
   * Semantics align with OOXML `font.vertAlign` values.
   */
  fontVariantPosition?: "subscript" | "superscript";
  textAlign?: CanvasTextAlign;
  /**
   * Text indentation in CSS pixels at zoom=1 (the renderer scales this by the current zoom).
   *
   * Semantics: the text content box is indented inwards from the *aligned edge*:
   * - left/start alignment → extra padding on the left edge
   * - right/end alignment → extra padding on the right edge
   * - center alignment → indent is ignored (deterministic)
   */
  textIndentPx?: number;
  /**
   * Spreadsheet-like horizontal alignment semantics that cannot be represented by
   * {@link CanvasTextAlign}.
   *
   * - `"fill"`: repeat the cell text to fill the available width (Excel "Fill")
   * - `"justify"`: justify wrapped text by expanding spaces between words
   *
   * Renderers may still use {@link CellStyle.textAlign} as a fallback baseline alignment.
   */
  horizontalAlign?: "fill" | "justify";
  underline?: boolean;
  /**
   * Optional underline variant.
   *
   * Note: `underline=true` implies `"single"`. When set to `"double"`, the renderer draws two underline strokes
   * (Excel-like).
   */
  underlineStyle?: "single" | "double";
  strike?: boolean;
  borders?: CellBorders;
  diagonalBorders?: CellDiagonalBorders;
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
  /**
   * Per-run style overrides (best-effort).
   *
   * The grid renderer currently supports the same keys as the legacy desktop rich-text renderer:
   * - `bold?: boolean`
   * - `italic?: boolean`
   * - `underline?: string | boolean` (anything except `"none"` is treated as underlined)
   * - `color?: string` (engine colors use `#AARRGGBB`)
   * - `font?: string`
   * - `size_100pt?: number` (font size in 1/100 points)
   */
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
   * Optional in-cell image payload (Excel "image in cell" / `IMAGE()` values).
   *
   * The `imageId` is a stable identifier that can be resolved by host applications via
   * the {@link CanvasGridRenderer} image resolver API.
   */
  image?: { imageId: string; altText?: string; width?: number; height?: number };
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
