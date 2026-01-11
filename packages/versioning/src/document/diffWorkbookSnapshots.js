import { isDeepStrictEqual } from "node:util";

import { semanticDiff } from "../diff/semanticDiff.js";
import { workbookStateFromDocumentSnapshot } from "./workbookState.js";

/**
 * @typedef {{ id: string, name: string | null }} SheetMeta
 * @typedef {{ row: number, col: number }} CellRef
 * @typedef {{ oldLocation: CellRef, newLocation: CellRef, value: any, formula?: string | null }} MoveChange
 * @typedef {{ cell: CellRef }} CellChange
 *
 * @typedef {{
 *   added: CellChange[];
 *   removed: CellChange[];
 *   modified: CellChange[];
 *   moved: MoveChange[];
 *   formatOnly: CellChange[];
 * }} SheetDiff
 *
 * @typedef {{
 *   sheetId: string;
 *   sheetName: string | null;
 *   diff: SheetDiff;
 * }} SheetDiffEntry
 *
 * @typedef {{
 *   added: SheetMeta[];
 *   removed: SheetMeta[];
 *   renamed: { id: string, beforeName: string | null, afterName: string | null }[];
 * }} SheetsDiff
 *
 * @typedef {{
 *   added: { id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }[];
 *   removed: { id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }[];
 *   modified: { id: string, before: any, after: any }[];
 * }} CommentsDiff
 *
 * @typedef {{
 *   added: { key: string, value: any }[];
 *   removed: { key: string, value: any }[];
 *   modified: { key: string, before: any, after: any }[];
 * }} NamedRangesDiff
 *
 * @typedef {{
 *   sheets: SheetsDiff;
 *   cellsBySheet: SheetDiffEntry[];
 *   comments: CommentsDiff;
 *   namedRanges: NamedRangesDiff;
 * }} WorkbookDiff
 */

/**
 * @param {string} a
 * @param {string} b
 */
function compareStrings(a, b) {
  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

/**
 * @param {CellRef} a
 * @param {CellRef} b
 */
function compareCellRefs(a, b) {
  if (a.row !== b.row) return a.row - b.row;
  return a.col - b.col;
}

/**
 * @param {SheetDiff} diff
 * @returns {SheetDiff}
 */
function sortSheetDiff(diff) {
  return {
    added: [...diff.added].sort((a, b) => compareCellRefs(a.cell, b.cell)),
    removed: [...diff.removed].sort((a, b) => compareCellRefs(a.cell, b.cell)),
    modified: [...diff.modified].sort((a, b) => compareCellRefs(a.cell, b.cell)),
    formatOnly: [...diff.formatOnly].sort((a, b) => compareCellRefs(a.cell, b.cell)),
    moved: [...diff.moved].sort((a, b) => {
      const cmpOld = compareCellRefs(a.oldLocation, b.oldLocation);
      if (cmpOld !== 0) return cmpOld;
      return compareCellRefs(a.newLocation, b.newLocation);
    }),
  };
}

/**
 * Compute a workbook-level diff between two DocumentController snapshots.
 *
 * Unlike the sheet-only helpers, this compares the entire workbook and includes
 * collaboration-relevant metadata (when present in the snapshot JSON), such as
 * sheets, comments, and named ranges.
 *
 * @param {{ beforeSnapshot: Uint8Array, afterSnapshot: Uint8Array }} opts
 * @returns {WorkbookDiff}
 */
export function diffDocumentWorkbookSnapshots(opts) {
  const before = workbookStateFromDocumentSnapshot(opts.beforeSnapshot);
  const after = workbookStateFromDocumentSnapshot(opts.afterSnapshot);

  const beforeSheetsById = new Map(before.sheets.map((s) => [s.id, s]));
  const afterSheetsById = new Map(after.sheets.map((s) => [s.id, s]));

  /** @type {SheetsDiff} */
  const sheets = { added: [], removed: [], renamed: [] };
  for (const [id, sheet] of afterSheetsById) {
    if (!beforeSheetsById.has(id)) sheets.added.push(sheet);
  }
  for (const [id, sheet] of beforeSheetsById) {
    if (!afterSheetsById.has(id)) sheets.removed.push(sheet);
  }
  for (const [id, afterSheet] of afterSheetsById) {
    const beforeSheet = beforeSheetsById.get(id);
    if (!beforeSheet) continue;
    if ((beforeSheet.name ?? null) !== (afterSheet.name ?? null)) {
      sheets.renamed.push({ id, beforeName: beforeSheet.name ?? null, afterName: afterSheet.name ?? null });
    }
  }
  sheets.added.sort((a, b) => compareStrings(a.id, b.id));
  sheets.removed.sort((a, b) => compareStrings(a.id, b.id));
  sheets.renamed.sort((a, b) => compareStrings(a.id, b.id));

  const sheetIds = new Set([...before.cellsBySheet.keys(), ...after.cellsBySheet.keys()]);
  /** @type {SheetDiffEntry[]} */
  const cellsBySheet = [];
  for (const sheetId of Array.from(sheetIds).sort(compareStrings)) {
    const beforeSheetState = before.cellsBySheet.get(sheetId) ?? { cells: new Map() };
    const afterSheetState = after.cellsBySheet.get(sheetId) ?? { cells: new Map() };
    const diff = semanticDiff(beforeSheetState, afterSheetState);
    const sheetName = afterSheetsById.get(sheetId)?.name ?? beforeSheetsById.get(sheetId)?.name ?? null;
    cellsBySheet.push({ sheetId, sheetName, diff: sortSheetDiff(diff) });
  }

  /** @type {CommentsDiff} */
  const comments = { added: [], removed: [], modified: [] };
  for (const [id, comment] of after.comments) {
    if (!before.comments.has(id)) comments.added.push(comment);
  }
  for (const [id, comment] of before.comments) {
    if (!after.comments.has(id)) comments.removed.push(comment);
  }
  for (const [id, afterComment] of after.comments) {
    const beforeComment = before.comments.get(id);
    if (!beforeComment) continue;
    const changed =
      (beforeComment.cellRef ?? null) !== (afterComment.cellRef ?? null) ||
      (beforeComment.content ?? null) !== (afterComment.content ?? null) ||
      beforeComment.resolved !== afterComment.resolved ||
      beforeComment.repliesLength !== afterComment.repliesLength;
    if (changed) comments.modified.push({ id, before: beforeComment, after: afterComment });
  }

  comments.added.sort((a, b) => compareStrings(a.id, b.id));
  comments.removed.sort((a, b) => compareStrings(a.id, b.id));
  comments.modified.sort((a, b) => compareStrings(a.id, b.id));

  /** @type {NamedRangesDiff} */
  const namedRanges = { added: [], removed: [], modified: [] };
  for (const [key, value] of after.namedRanges) {
    if (!before.namedRanges.has(key)) namedRanges.added.push({ key, value });
  }
  for (const [key, value] of before.namedRanges) {
    if (!after.namedRanges.has(key)) namedRanges.removed.push({ key, value });
  }
  for (const [key, afterValue] of after.namedRanges) {
    const beforeValue = before.namedRanges.get(key);
    if (beforeValue === undefined) continue;
    if (!isDeepStrictEqual(beforeValue, afterValue)) {
      namedRanges.modified.push({ key, before: beforeValue, after: afterValue });
    }
  }

  namedRanges.added.sort((a, b) => compareStrings(a.key, b.key));
  namedRanges.removed.sort((a, b) => compareStrings(a.key, b.key));
  namedRanges.modified.sort((a, b) => compareStrings(a.key, b.key));

  return { sheets, cellsBySheet, comments, namedRanges };
}

