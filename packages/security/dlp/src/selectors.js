import { CLASSIFICATION_LEVEL, DEFAULT_CLASSIFICATION, maxClassification } from "./classification.js";
import dlpCore from "./core.js";

const { normalizeRange, selectorKey } = dlpCore;

export { normalizeRange, selectorKey };

/**
 * A selector identifies the scope a classification applies to.
 *
 * Supported scopes:
 * - document
 * - sheet
 * - range (rectangular)
 * - column (by sheet column index; table columns should map here at runtime)
 * - cell
 */

export const CLASSIFICATION_SCOPE = Object.freeze({
  DOCUMENT: "document",
  SHEET: "sheet",
  RANGE: "range",
  COLUMN: "column",
  CELL: "cell",
});

/**
 * @param {number} colIndex 0-based
 */
export function colIndexToA1(colIndex) {
  if (!Number.isInteger(colIndex) || colIndex < 0) {
    throw new Error(`Invalid column index: ${colIndex}`);
  }
  let n = colIndex + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * @param {string} colLetters
 */
export function a1ToColIndex(colLetters) {
  const s = String(colLetters).toUpperCase();
  if (!/^[A-Z]+$/.test(s)) throw new Error(`Invalid column letters: ${colLetters}`);
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

/**
 * @param {{row:number,col:number}} cell 0-based row/col
 */
export function cellToA1(cell) {
  if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col)) {
    throw new Error(`Invalid cell: ${JSON.stringify(cell)}`);
  }
  return `${colIndexToA1(cell.col)}${cell.row + 1}`;
}

/**
 * @param {string} a1
 */
export function a1ToCell(a1) {
  const s = String(a1).trim().toUpperCase();
  const match = /^([A-Z]+)(\d+)$/.exec(s);
  if (!match) throw new Error(`Invalid A1 cell: ${a1}`);
  const col = a1ToColIndex(match[1]);
  const row = Number(match[2]) - 1;
  if (!Number.isInteger(row) || row < 0) throw new Error(`Invalid A1 row: ${a1}`);
  return { row, col };
}

/**
 * @param {{row:number,col:number}} cell
 * @param {{start:{row:number,col:number}, end:{row:number,col:number}}} range
 */
export function cellInRange(cell, range) {
  const r = normalizeRange(range);
  return (
    cell.row >= r.start.row &&
    cell.row <= r.end.row &&
    cell.col >= r.start.col &&
    cell.col <= r.end.col
  );
}

/**
 * @param {{start:{row:number,col:number}, end:{row:number,col:number}}} a
 * @param {{start:{row:number,col:number}, end:{row:number,col:number}}} b
 */
export function rangesIntersect(a, b) {
  const ra = normalizeRange(a);
  const rb = normalizeRange(b);
  const rowOverlap = ra.start.row <= rb.end.row && rb.start.row <= ra.end.row;
  const colOverlap = ra.start.col <= rb.end.col && rb.start.col <= ra.end.col;
  return rowOverlap && colOverlap;
}

/**
 * @param {any} selector
 */
/**
 * @param {any} selector
 * @param {{documentId:string, sheetId?:string, row?:number, col?:number}} cellRef
 */
export function selectorAppliesToCell(selector, cellRef) {
  if (!selector || typeof selector !== "object") return false;
  if (selector.documentId !== cellRef.documentId) return false;
  switch (selector.scope) {
    case CLASSIFICATION_SCOPE.DOCUMENT:
      return true;
    case CLASSIFICATION_SCOPE.SHEET:
      return selector.sheetId === cellRef.sheetId;
    case CLASSIFICATION_SCOPE.COLUMN:
      if (selector.sheetId !== cellRef.sheetId) return false;
      if (typeof selector.columnIndex === "number") return selector.columnIndex === cellRef.col;
      if (selector.tableId && selector.columnId && cellRef.tableId && cellRef.columnId) {
        return selector.tableId === cellRef.tableId && selector.columnId === cellRef.columnId;
      }
      return false;
    case CLASSIFICATION_SCOPE.RANGE:
      if (selector.sheetId !== cellRef.sheetId) return false;
      return cellInRange({ row: cellRef.row, col: cellRef.col }, selector.range);
    case CLASSIFICATION_SCOPE.CELL:
      return (
        selector.sheetId === cellRef.sheetId &&
        selector.row === cellRef.row &&
        selector.col === cellRef.col
      );
    default:
      return false;
  }
}

/**
 * Compute the effective classification for a single cell given a set of classification
 * records. The effective classification is the most restrictive classification that
 * applies to the cell.
 *
 * This "max" semantics intentionally prevents a less-restrictive classification at a
 * narrower scope from weakening a broader restriction (e.g., a Public cell inside a
 * Restricted range). This is the safest default for DLP enforcement.
 *
 * @param {{documentId:string, sheetId:string, row:number, col:number}} cellRef
 * @param {Array<{selector:any, classification:any}>} records
 */
export function effectiveCellClassification(cellRef, records) {
  let result = { ...DEFAULT_CLASSIFICATION };
  for (const record of records || []) {
    if (!record) continue;
    if (!selectorAppliesToCell(record.selector, cellRef)) continue;
    result = maxClassification(result, record.classification);
    if (result.level === CLASSIFICATION_LEVEL.RESTRICTED) return result;
  }
  return result;
}

/**
 * Returns the most restrictive classification that intersects the provided range.
 *
 * This is used for DLP actions that operate on a selection (clipboard/export). If any
 * restricted/confidential data is in scope, the operation should treat the whole
 * selection as that classification.
 *
 * @param {{documentId:string, sheetId:string, range:{start:{row:number,col:number}, end:{row:number,col:number}}}} rangeRef
 * @param {Array<{selector:any, classification:any}>} records
 */
export function effectiveRangeClassification(rangeRef, records) {
  const normalized = normalizeRange(rangeRef.range);
  let result = { ...DEFAULT_CLASSIFICATION };
  for (const record of records || []) {
    if (!record || !record.selector) continue;
    if (record.selector.documentId !== rangeRef.documentId) continue;

    switch (record.selector.scope) {
      case CLASSIFICATION_SCOPE.DOCUMENT:
        result = maxClassification(result, record.classification);
        break;
      case CLASSIFICATION_SCOPE.SHEET:
        if (record.selector.sheetId === rangeRef.sheetId) {
          result = maxClassification(result, record.classification);
        }
        break;
      case CLASSIFICATION_SCOPE.COLUMN:
        if (record.selector.sheetId !== rangeRef.sheetId) break;
        if (typeof record.selector.columnIndex !== "number") break;
        if (record.selector.columnIndex >= normalized.start.col && record.selector.columnIndex <= normalized.end.col) {
          result = maxClassification(result, record.classification);
        }
        break;
      case CLASSIFICATION_SCOPE.RANGE:
        if (record.selector.sheetId !== rangeRef.sheetId) break;
        if (rangesIntersect(record.selector.range, normalized)) {
          result = maxClassification(result, record.classification);
        }
        break;
      case CLASSIFICATION_SCOPE.CELL:
        if (record.selector.sheetId !== rangeRef.sheetId) break;
        if (cellInRange({ row: record.selector.row, col: record.selector.col }, normalized)) {
          result = maxClassification(result, record.classification);
        }
        break;
      default:
        break;
    }

    if (result.level === CLASSIFICATION_LEVEL.RESTRICTED) return result;
  }
  return result;
}

/**
 * Returns the most restrictive classification present anywhere in the document.
 *
 * This is intentionally conservative and is appropriate for whole-document operations
 * like external link sharing or exporting the full workbook.
 *
 * @param {string} documentId
 * @param {Array<{selector:any, classification:any}>} records
 */
export function effectiveDocumentClassification(documentId, records) {
  let result = { ...DEFAULT_CLASSIFICATION };
  for (const record of records || []) {
    if (!record || !record.selector) continue;
    if (record.selector.documentId !== documentId) continue;
    result = maxClassification(result, record.classification);
    if (result.level === CLASSIFICATION_LEVEL.RESTRICTED) return result;
  }
  return result;
}
