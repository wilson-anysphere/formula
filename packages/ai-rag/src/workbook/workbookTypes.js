/**
 * Shared structural types (documentation-only; runtime is plain JS).
 *
 * @typedef {Object} Cell
 * @property {string|number|boolean|null|undefined} [v] Cell value
 * @property {string|null|undefined} [f] Formula string, e.g. "=SUM(A1:A10)"
 *
 * @typedef {Object} Sheet
 * @property {string} name
 * @property {Cell[][]} cells 2D array [row][col]
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
 * @property {Cell[][]} cells
 * @property {any} [meta]
 */

export {};
