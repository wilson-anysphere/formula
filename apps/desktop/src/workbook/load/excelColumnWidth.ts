// Thin wrapper around the shared conversion helpers in `@formula/engine`.
//
// The UI sheet view stores widths in CSS px (zoom=1) while Excel stores widths in "character"
// units (OOXML `col/@width`). The shared helpers implement the deterministic conversion contract
// used across the desktop app + WASM engine.

import { excelColWidthCharsToPixels, pixelsToExcelColWidthChars } from "@formula/engine";

/**
 * Convert an Excel column width expressed in "character" units into a pixel width (zoom=1).
 *
 * This is the value stored in OOXML `<col width="...">` and surfaced by `CELL("width")`.
 */
export function excelColWidthCharsToPx(widthChars: number): number {
  return excelColWidthCharsToPixels(widthChars);
}

/**
 * Convert a pixel column width (zoom=1) back into Excel "character" units.
 *
 * This is primarily used for round-tripping UI-driven resizing into Excel-compatible formula
 * semantics (`CELL("width")`).
 */
export function excelColWidthPxToChars(widthPx: number): number {
  return pixelsToExcelColWidthChars(widthPx);
}
