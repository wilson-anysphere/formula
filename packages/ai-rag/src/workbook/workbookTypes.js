/**
 * Shared structural types (documentation-only; runtime is plain JS).
 *
 * @typedef {Object} Cell
 * @property {string|number|boolean|null|undefined} [v] Cell value
 * @property {string|null|undefined} [f] Formula string, e.g. "=SUM(A1:A10)"
 *
 * Alternative supported input shapes (auto-normalized by `normalizeCell`):
 * - `{ value, formula }` (e.g. DocumentController / ai-tools cell data)
 * - raw scalar values; strings starting with "=" are treated as formulas
 *
 * @typedef {Object} Sheet
 * @property {string} name
 * @property {Cell[][] | Map<string, any>} cells
 *   Either a 2D matrix `[row][col]` or a sparse Map keyed by `"row,col"` or `"row:col"`.
 * @property {any[][]} [values] Optional alternative to `cells` (treated as `[row][col]`).
 *
 * @typedef {Object} WorkbookTable
 * @property {string} name
 * @property {string} sheetName
 * @property {{ r0: number, c0: number, r1: number, c1: number }} rect inclusive 0-based
 *
 * @typedef {Object} NamedRange
 * @property {string} name
 * @property {string} sheetName
 * @property {{ r0: number, c0: number, r1: number, c1: number }} rect inclusive 0-based
 *
 * @typedef {Object} Workbook
 * @property {string} id
 * @property {Sheet[]} sheets
 * @property {WorkbookTable[]} [tables]
 * @property {NamedRange[]} [namedRanges]
 *
 * @typedef {Object} WorkbookChunk
 * @property {string} id
 * @property {string} workbookId
 * @property {string} sheetName
 * @property {'table'|'namedRange'|'dataRegion'|'formulaRegion'} kind
 * @property {string} title
 * @property {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @property {Cell[][]} cells Sampled window of cells (bounded for embedding), aligned to `rect.r0/c0`.
 * @property {any} [meta]
 */

export {};
