/**
 * Excel column width conversions.
 *
 * The core engine (and `formula_model::worksheet::ColProperties.width`) stores column widths
 * in Excel's "character" units (OOXML `<col width="...">`).
 *
 * The UI grid / sheet view state stores widths in **CSS pixels** (at `zoom = 1`).
 *
 * These helpers implement Excel's conversion for the *default* font metrics used by Excel:
 *
 * - Calibri 11
 * - max digit width (MDW) = 7px
 * - padding = 5px
 *
 * Note: Excel's width<->pixel mapping depends on the workbook's default font and can vary
 * across environments. This implementation is intended to match Excel's behavior well
 * enough for `CELL("width")` and round-tripping common default widths.
 */

export const EXCEL_DEFAULT_MAX_DIGIT_WIDTH_PX = 7;
export const EXCEL_DEFAULT_CELL_PADDING_PX = 5;

export type ExcelColumnWidthConversionOptions = {
  /**
   * Max digit width (MDW) in pixels for the default font.
   *
   * Excel uses this to define the meaning of a "character" width.
   */
  maxDigitWidthPx?: number;
  /**
   * Per-column padding in pixels added by Excel.
   */
  paddingPx?: number;
  /**
   * Number of decimals to round to when converting pixels -> character widths.
   *
   * Excel's UI (and typical OOXML output) uses 2 decimals.
   */
  decimals?: number;
};

function roundDecimals(value: number, decimals: number): number {
  if (!Number.isFinite(value)) return 0;
  const pow = 10 ** decimals;
  return Math.round(value * pow) / pow;
}

/**
 * Convert an Excel column width in "character" units (OOXML `col/@width`) to pixels.
 */
export function excelColWidthCharsToPixels(widthChars: number, options: ExcelColumnWidthConversionOptions = {}): number {
  if (!Number.isFinite(widthChars) || widthChars <= 0) return 0;

  const maxDigitWidthPx = options.maxDigitWidthPx ?? EXCEL_DEFAULT_MAX_DIGIT_WIDTH_PX;
  const paddingPx = options.paddingPx ?? EXCEL_DEFAULT_CELL_PADDING_PX;

  // Excel uses an extra offset term (`TRUNC(128 / MDW)`) to match its internal rounding.
  // See e.g. Apache POI's width conversion helpers and many XLSX libraries.
  const offset = Math.floor(128 / maxDigitWidthPx);

  return (
    Math.floor(((256 * widthChars + offset) / 256) * maxDigitWidthPx) +
    paddingPx
  );
}

/**
 * Convert pixel width (CSS px at `zoom = 1`) to an Excel column width in "character" units.
 */
export function pixelsToExcelColWidthChars(pixels: number, options: ExcelColumnWidthConversionOptions = {}): number {
  if (!Number.isFinite(pixels) || pixels <= 0) return 0;

  const maxDigitWidthPx = options.maxDigitWidthPx ?? EXCEL_DEFAULT_MAX_DIGIT_WIDTH_PX;
  const paddingPx = options.paddingPx ?? EXCEL_DEFAULT_CELL_PADDING_PX;
  const decimals = options.decimals ?? 2;

  // Excel pixel sizes are integer. Round to stabilize widths when UI code passes
  // fractional values (e.g. from zoom interpolation).
  const px = Math.round(pixels);
  const inner = Math.max(0, px - paddingPx);

  // Excel stores widths in 1/256th character increments internally. Quantize to that
  // precision, then round to the typical 2-decimal UI representation.
  const widthUnits256 = Math.round((inner * 256) / maxDigitWidthPx);
  const widthChars = widthUnits256 / 256;
  return roundDecimals(widthChars, decimals);
}

