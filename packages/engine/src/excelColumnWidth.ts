/**
 * Excel/OpenXML column width unit conversions.
 *
 * DocumentController stores interactive axis sizes in **CSS pixels at zoom=1** ("base" units).
 * The Rust formula model (and OOXML `col/@width`) stores column widths in **Excel character**
 * units: the number of `0` (digit) glyphs that fit in the cell, for the workbook's default
 * font, with 1/256 character precision.
 *
 * Excel's pixel conversion is documented in various OOXML/BIFF references and is commonly
 * implemented as:
 *   pixels = truncate(((256 * width + truncate(128 / MDW)) / 256) * MDW) + padding
 * where:
 *   MDW = max digit width of the default font (Calibri 11 @ 96dpi => 7px)
 *   padding = 5px
 *
 * We invert that relationship and quantize to 1/256th character increments so values round-trip
 * through the engine model consistently.
 */

const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
const EXCEL_COL_PADDING_PX = 5;
const EXCEL_COL_WIDTH_DENOM = 256;

/**
 * Convert a DocumentController "base" column width (CSS px @ zoom=1) into Excel/OOXML column
 * width units (characters, 1/256 precision).
 */
export function docColWidthPxToExcelChars(widthPx: number): number {
  if (!Number.isFinite(widthPx)) return NaN;

  // Excel's algorithm bakes in ~5px padding; guard small widths so we never emit negative
  // character widths.
  const px = Math.max(0, widthPx);
  const width256 = Math.floor(((px - EXCEL_COL_PADDING_PX) / EXCEL_MAX_DIGIT_WIDTH_PX) * EXCEL_COL_WIDTH_DENOM);
  return Math.max(0, width256) / EXCEL_COL_WIDTH_DENOM;
}

