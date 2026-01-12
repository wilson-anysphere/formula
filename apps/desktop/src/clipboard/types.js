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
 * @typedef {{ html?: string, text?: string, rtf?: string, imagePng?: Uint8Array }} ClipboardContent
 * @typedef {{ text: string, html?: string, rtf?: string, imagePng?: Uint8Array }} ClipboardWritePayload
 *
 * @typedef {"all" | "values" | "formulas" | "formats"} PasteSpecialMode
 * @typedef {{ mode?: PasteSpecialMode }} PasteOptions
 */

export {};
