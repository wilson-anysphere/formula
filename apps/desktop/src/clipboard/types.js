/**
 * Clipboard-specific data shapes.
 *
 * We intentionally align these with the DocumentController's `CellState` so clipboard
 * paste can round-trip through `setRangeValues` without bespoke adapters.
 *
 * @typedef {import("../document/cell.js").CellValue} CellValue
 * @typedef {import("../document/cell.js").CellFormat} CellFormat
 * @typedef {import("../document/cell.js").CellState} CellState
 *
 * @typedef {CellState[][]} CellGrid
 *
 * @typedef {{ html?: string, text?: string }} ClipboardContent
 * @typedef {{ text: string, html?: string }} ClipboardWritePayload
 *
 * @typedef {"all" | "values" | "formulas" | "formats"} PasteSpecialMode
 * @typedef {{ mode?: PasteSpecialMode }} PasteOptions
 */

export {};

