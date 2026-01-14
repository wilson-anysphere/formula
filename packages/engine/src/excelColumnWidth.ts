/**
 * Excel/OpenXML column width unit conversions.
 *
 * Note: this file exists for backwards compatibility with older code that imported
 * `docColWidthPxToExcelChars` directly. New code should prefer the more general helpers in
 * `columnWidth.ts` (which also supports Excel chars -> pixels).
 */

import { pixelsToExcelColWidthChars } from "./columnWidth.ts";

/**
 * Convert a DocumentController "base" column width (CSS px @ zoom=1) into Excel/OOXML column
 * width units (characters, 1/256 precision).
 */
export function docColWidthPxToExcelChars(widthPx: number): number {
  return pixelsToExcelColWidthChars(widthPx);
}
