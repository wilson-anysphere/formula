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
 * @typedef {{
 *   html?: string,
 *   text?: string,
 *   rtf?: string,
 *   imagePng?: Uint8Array,
 *   pngBase64?: string
 * }} ClipboardContent
 *
 * `pngBase64` is a legacy/internal field kept for backwards compatibility.
 * Callers should prefer `imagePng` (raw PNG bytes).
 *
 * @typedef {{
 *   text: string,
 *   html?: string,
 *   rtf?: string,
 *   imagePng?: Uint8Array | ArrayBuffer | ArrayBufferView | Blob,
 *   pngBase64?: string
 * }} ClipboardWritePayload
 *
 * Notes:
 * - `imagePng` is the preferred API for image clipboard writes (raw PNG bytes).
 * - `pngBase64` is a legacy/internal field; base64 is only used as a wire format for Tauri IPC.
 *
 * @typedef {"all" | "values" | "formulas" | "formats"} PasteSpecialMode
 * @typedef {{ mode?: PasteSpecialMode }} PasteOptions
 */

export {};
