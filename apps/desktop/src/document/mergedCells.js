/**
 * Merged-cell helpers for the desktop DocumentController model.
 *
 * Excel semantics:
 * - Merged regions are anchored at their top-left cell.
 * - Values/formulas in non-anchor cells are discarded on merge.
 * - Unmerge only removes merge metadata; it does not restore discarded values.
 *
 * Ranges use inclusive coordinates (`endRow`/`endCol` are inclusive), matching
 * the desktop selection model and most DocumentController APIs.
 */

/**
 * @typedef {{ startRow: number; endRow: number; startCol: number; endCol: number }} MergedRange
 */

/**
 * @param {MergedRange} range
 * @returns {MergedRange}
 */
function normalizeMergedRange(range) {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

/**
 * @param {MergedRange} a
 * @param {MergedRange} b
 * @returns {boolean}
 */
function rangesIntersect(a, b) {
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol);
}

/**
 * @param {MergedRange[]} ranges
 * @returns {MergedRange[]}
 */
function sortRanges(ranges) {
  return ranges.slice().sort((a, b) => {
    if (a.startRow !== b.startRow) return a.startRow - b.startRow;
    if (a.startCol !== b.startCol) return a.startCol - b.startCol;
    if (a.endRow !== b.endRow) return a.endRow - b.endRow;
    return a.endCol - b.endCol;
  });
}

/**
 * Merge a rectangular range into a single merged cell.
 *
 * - Any existing merged regions intersecting `range` are unmerged first (Excel-like).
 * - Values/formulas in non-anchor cells are cleared (formatting is preserved).
 *
 * @param {import("./documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {MergedRange} range
 * @param {{ label?: string, mergeKey?: string }} [options]
 * @returns {boolean} Whether any merge metadata was added.
 */
export function mergeCells(doc, sheetId, range, options = {}) {
  const r = normalizeMergedRange(range);
  const isSingleCell = r.startRow === r.endRow && r.startCol === r.endCol;
  if (isSingleCell) return false;

  const existing = typeof doc.getMergedRanges === "function" ? doc.getMergedRanges(sheetId) : [];
  const remaining = existing.filter((m) => !rangesIntersect(m, r));
  const next = sortRanges([...remaining, r]);

  if (typeof doc.setMergedRanges === "function") {
    doc.setMergedRanges(sheetId, next, options);
  }

  // Clear values/formulas in non-anchor cells (preserve formatting).
  /** @type {Array<{ sheetId: string, row: number, col: number, value: any, formula: string | null }>} */
  const clearInputs = [];
  const anchorRow = r.startRow;
  const anchorCol = r.startCol;
  if (typeof doc.forEachCellInSheet === "function") {
    doc.forEachCellInSheet(sheetId, ({ row, col, cell }) => {
      if (row < r.startRow || row > r.endRow) return;
      if (col < r.startCol || col > r.endCol) return;
      if (row === anchorRow && col === anchorCol) return;
      if (cell?.value == null && cell?.formula == null) return;
      clearInputs.push({ sheetId, row, col, value: null, formula: null });
    });
  }
  if (clearInputs.length > 0 && typeof doc.setCellInputs === "function") {
    doc.setCellInputs(clearInputs, { label: options.label });
  }

  return true;
}

/**
 * Excel "Merge Across": merge each row segment independently.
 *
 * Example: A1:C3 => merges A1:C1, A2:C2, A3:C3.
 *
 * @param {import("./documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {MergedRange} range
 * @param {{ label?: string, mergeKey?: string }} [options]
 * @returns {boolean} Whether any merges were added.
 */
export function mergeAcross(doc, sheetId, range, options = {}) {
  const r = normalizeMergedRange(range);
  if (r.startCol === r.endCol) return false;

  const existing = typeof doc.getMergedRanges === "function" ? doc.getMergedRanges(sheetId) : [];
  // Any merges that intersect the selection are removed first (Excel-like).
  const remaining = existing.filter((m) => !rangesIntersect(m, r));

  /** @type {MergedRange[]} */
  const newMerges = [];
  for (let row = r.startRow; row <= r.endRow; row += 1) {
    newMerges.push({ startRow: row, endRow: row, startCol: r.startCol, endCol: r.endCol });
  }

  const next = sortRanges([...remaining, ...newMerges]);

  if (typeof doc.setMergedRanges === "function") {
    doc.setMergedRanges(sheetId, next, options);
  }

  // Clear values/formulas in non-anchor cells (preserve formatting). Each merged row is anchored
  // at its left-most cell.
  /** @type {Array<{ sheetId: string, row: number, col: number, value: any, formula: string | null }>} */
  const clearInputs = [];
  const anchorCol = r.startCol;
  if (typeof doc.forEachCellInSheet === "function") {
    doc.forEachCellInSheet(sheetId, ({ row, col, cell }) => {
      if (row < r.startRow || row > r.endRow) return;
      if (col < r.startCol || col > r.endCol) return;
      if (col === anchorCol) return;
      if (cell?.value == null && cell?.formula == null) return;
      clearInputs.push({ sheetId, row, col, value: null, formula: null });
    });
  }
  if (clearInputs.length > 0 && typeof doc.setCellInputs === "function") {
    doc.setCellInputs(clearInputs, { label: options.label });
  }

  return true;
}

/**
 * Merge + Center (Excel semantics): merge the range and set horizontal alignment to center
 * on the merged cell's anchor (top-left).
 *
 * @param {import("./documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {MergedRange} range
 * @param {{ label?: string, mergeKey?: string }} [options]
 * @returns {boolean}
 */
export function mergeCenter(doc, sheetId, range, options = {}) {
  const r = normalizeMergedRange(range);
  const didMerge = mergeCells(doc, sheetId, r, options);
  if (!didMerge) return false;

  // Apply alignment to the anchor cell only.
  if (typeof doc.setRangeFormat === "function") {
    doc.setRangeFormat(
      sheetId,
      { start: { row: r.startRow, col: r.startCol }, end: { row: r.startRow, col: r.startCol } },
      { alignment: { horizontal: "center" } },
      { label: options.label ?? "Merge & Center" },
    );
  }

  return true;
}

/**
 * Unmerge any merged regions that intersect `range`.
 *
 * @param {import("./documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {MergedRange} range
 * @param {{ label?: string, mergeKey?: string }} [options]
 * @returns {number} Number of merged regions removed.
 */
export function unmergeCells(doc, sheetId, range, options = {}) {
  const r = normalizeMergedRange(range);
  const existing = typeof doc.getMergedRanges === "function" ? doc.getMergedRanges(sheetId) : [];
  if (existing.length === 0) return 0;
  const remaining = existing.filter((m) => !rangesIntersect(m, r));
  const removed = existing.length - remaining.length;
  if (removed <= 0) return 0;

  if (typeof doc.setMergedRanges === "function") {
    doc.setMergedRanges(sheetId, remaining, options);
  }
  return removed;
}
