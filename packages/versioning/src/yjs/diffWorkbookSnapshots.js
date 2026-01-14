import { semanticDiff } from "../diff/semanticDiff.js";
import { deepEqual } from "../diff/deepEqual.js";
import { workbookStateFromYjsSnapshot } from "./workbookState.js";

/**
 * @typedef {"visible" | "hidden" | "veryHidden"} SheetVisibility
 * @typedef {{ frozenRows: number, frozenCols: number }} SheetViewMeta
 * @typedef {{ id: string, name: string | null, visibility: SheetVisibility, tabColor: string | null, view: SheetViewMeta }} SheetMeta
 * @typedef {{
 *   id: string,
 *   name: string | null,
 *   afterIndex: number,
 *   visibility: SheetVisibility,
 *   tabColor: string | null,
 *   view: SheetViewMeta,
 * }} AddedSheet
 * @typedef {{
 *   id: string,
 *   name: string | null,
 *   beforeIndex: number,
 *   visibility: SheetVisibility,
 *   tabColor: string | null,
 *   view: SheetViewMeta,
 * }} RemovedSheet
 * @typedef {{ id: string, beforeIndex: number, afterIndex: number }} MovedSheet
 * @typedef {{ id: string, field: string, before: any, after: any }} SheetMetaChange
 * @typedef {{ row: number, col: number }} CellRef
 * @typedef {{
 *   oldLocation: CellRef,
 *   newLocation: CellRef,
 *   value: any,
 *   formula?: string | null,
 *   encrypted?: boolean,
 *   keyId?: string | null,
 * }} MoveChange
 * @typedef {{
 *   cell: CellRef,
 *   oldValue?: any,
 *   newValue?: any,
 *   oldFormula?: string | null,
 *   newFormula?: string | null,
 *   oldEncrypted?: boolean,
 *   newEncrypted?: boolean,
 *   oldKeyId?: string | null,
 *   newKeyId?: string | null,
 * }} CellChange
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
 *   added: AddedSheet[];
 *   removed: RemovedSheet[];
 *   renamed: { id: string, beforeName: string | null, afterName: string | null }[];
 *   moved: MovedSheet[];
 *   metaChanged: SheetMetaChange[];
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
 *   added: { key: string, value: any }[];
 *   removed: { key: string, value: any }[];
 *   modified: { key: string, before: any, after: any }[];
 * }} MetadataDiff
 *
 * @typedef {{
 *   sheets: SheetsDiff;
 *   cellsBySheet: SheetDiffEntry[];
 *   comments: CommentsDiff;
 *   metadata: MetadataDiff;
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
 * @param {number[]} arr
 * @returns {number[]}
 */
function longestIncreasingSubsequenceIndices(arr) {
  /** @type {number[]} */
  const tails = [];
  /** @type {number[]} */
  const prev = new Array(arr.length).fill(-1);

  for (let i = 0; i < arr.length; i++) {
    const x = arr[i];
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (arr[tails[mid]] < x) lo = mid + 1;
      else hi = mid;
    }
    if (lo > 0) prev[i] = tails[lo - 1];
    if (lo === tails.length) tails.push(i);
    else tails[lo] = i;
  }

  /** @type {number[]} */
  const out = [];
  let k = tails.length ? tails[tails.length - 1] : -1;
  while (k >= 0) {
    out.push(k);
    k = prev[k];
  }
  out.reverse();
  return out;
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
 * Compute a workbook-level diff between two Yjs snapshots.
 *
 * This extends the existing cell-only diff by including collaboration-relevant
 * metadata like sheets, comments, and named ranges.
 *
 * @param {{ beforeSnapshot: Uint8Array, afterSnapshot: Uint8Array }} opts
 * @returns {WorkbookDiff}
 */
export function diffYjsWorkbookSnapshots(opts) {
  const before = workbookStateFromYjsSnapshot(opts.beforeSnapshot);
  const after = workbookStateFromYjsSnapshot(opts.afterSnapshot);

  const beforeSheetsById = new Map(before.sheets.map((s) => [s.id, s]));
  const afterSheetsById = new Map(after.sheets.map((s) => [s.id, s]));

  /** @type {SheetsDiff} */
  const sheets = { added: [], removed: [], renamed: [], moved: [], metaChanged: [] };

  const beforeOrder = Array.isArray(before.sheetOrder) && before.sheetOrder.length ? before.sheetOrder : before.sheets.map((s) => s.id);
  const afterOrder = Array.isArray(after.sheetOrder) && after.sheetOrder.length ? after.sheetOrder : after.sheets.map((s) => s.id);

  /** @type {Map<string, number>} */
  const beforeIndex = new Map();
  beforeOrder.forEach((id, idx) => {
    if (!beforeIndex.has(id)) beforeIndex.set(id, idx);
  });
  /** @type {Map<string, number>} */
  const afterIndex = new Map();
  afterOrder.forEach((id, idx) => {
    if (!afterIndex.has(id)) afterIndex.set(id, idx);
  });

  for (const [id, idx] of afterIndex) {
    if (!beforeSheetsById.has(id)) {
      const sheet = afterSheetsById.get(id) ?? { id, name: null, visibility: "visible", tabColor: null, view: { frozenRows: 0, frozenCols: 0 } };
      sheets.added.push({
        id,
        name: sheet.name ?? null,
        afterIndex: idx,
        visibility: sheet.visibility ?? "visible",
        tabColor: sheet.tabColor ?? null,
        view: sheet.view ?? { frozenRows: 0, frozenCols: 0 },
      });
    }
  }

  for (const [id, idx] of beforeIndex) {
    if (!afterSheetsById.has(id)) {
      const sheet = beforeSheetsById.get(id) ?? { id, name: null, visibility: "visible", tabColor: null, view: { frozenRows: 0, frozenCols: 0 } };
      sheets.removed.push({
        id,
        name: sheet.name ?? null,
        beforeIndex: idx,
        visibility: sheet.visibility ?? "visible",
        tabColor: sheet.tabColor ?? null,
        view: sheet.view ?? { frozenRows: 0, frozenCols: 0 },
      });
    }
  }

  for (const [id, afterSheet] of afterSheetsById) {
    const beforeSheet = beforeSheetsById.get(id);
    if (!beforeSheet) continue;
    if ((beforeSheet.name ?? null) !== (afterSheet.name ?? null)) {
      sheets.renamed.push({ id, beforeName: beforeSheet.name ?? null, afterName: afterSheet.name ?? null });
    }
  }

  // Sheet metadata changes (visibility/tabColor/frozen panes).
  for (const [id, afterSheet] of afterSheetsById) {
    const beforeSheet = beforeSheetsById.get(id);
    if (!beforeSheet) continue;

    if ((beforeSheet.visibility ?? null) !== (afterSheet.visibility ?? null)) {
      sheets.metaChanged.push({
        id,
        field: "visibility",
        before: beforeSheet.visibility ?? null,
        after: afterSheet.visibility ?? null,
      });
    }

    if (!deepEqual(beforeSheet.tabColor ?? null, afterSheet.tabColor ?? null)) {
      sheets.metaChanged.push({
        id,
        field: "tabColor",
        before: beforeSheet.tabColor ?? null,
        after: afterSheet.tabColor ?? null,
      });
    }

    const beforeView = beforeSheet.view ?? null;
    const afterView = afterSheet.view ?? null;
    const beforeFrozenRows = beforeView?.frozenRows ?? null;
    const afterFrozenRows = afterView?.frozenRows ?? null;
    if (beforeFrozenRows !== afterFrozenRows) {
      sheets.metaChanged.push({ id, field: "view.frozenRows", before: beforeFrozenRows, after: afterFrozenRows });
    }

    const beforeFrozenCols = beforeView?.frozenCols ?? null;
    const afterFrozenCols = afterView?.frozenCols ?? null;
    if (beforeFrozenCols !== afterFrozenCols) {
      sheets.metaChanged.push({ id, field: "view.frozenCols", before: beforeFrozenCols, after: afterFrozenCols });
    }
  }

  // Sheet reorder detection: find the minimal set of sheets whose relative order changed.
  const afterIds = new Set(afterOrder);
  const beforeIds = new Set(beforeOrder);
  const beforeCommon = beforeOrder.filter((id) => afterIds.has(id));
  const afterCommon = afterOrder.filter((id) => beforeIds.has(id));

  /** @type {Map<string, number>} */
  const afterCommonIndex = new Map();
  afterCommon.forEach((id, idx) => afterCommonIndex.set(id, idx));
  const seq = beforeCommon.map((id) => afterCommonIndex.get(id) ?? -1);
  const lisIdx = new Set(longestIncreasingSubsequenceIndices(seq));
  for (let i = 0; i < beforeCommon.length; i++) {
    if (lisIdx.has(i)) continue;
    const id = beforeCommon[i];
    const bi = beforeIndex.get(id);
    const ai = afterIndex.get(id);
    if (bi == null || ai == null) continue;
    sheets.moved.push({ id, beforeIndex: bi, afterIndex: ai });
  }

  sheets.added.sort((a, b) => compareStrings(a.id, b.id));
  sheets.removed.sort((a, b) => compareStrings(a.id, b.id));
  sheets.renamed.sort((a, b) => compareStrings(a.id, b.id));
  sheets.moved.sort((a, b) => compareStrings(a.id, b.id));
  sheets.metaChanged.sort((a, b) => compareStrings(a.id, b.id) || compareStrings(a.field, b.field));

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
    if (!before.comments.has(id)) {
      comments.added.push(comment);
    }
  }
  for (const [id, comment] of before.comments) {
    if (!after.comments.has(id)) {
      comments.removed.push(comment);
    }
  }
  for (const [id, afterComment] of after.comments) {
    const beforeComment = before.comments.get(id);
    if (!beforeComment) continue;
    const changed =
      (beforeComment.cellRef ?? null) !== (afterComment.cellRef ?? null) ||
      (beforeComment.content ?? null) !== (afterComment.content ?? null) ||
      beforeComment.resolved !== afterComment.resolved ||
      beforeComment.repliesLength !== afterComment.repliesLength;
    if (changed) {
      comments.modified.push({ id, before: beforeComment, after: afterComment });
    }
  }

  comments.added.sort((a, b) => compareStrings(a.id, b.id));
  comments.removed.sort((a, b) => compareStrings(a.id, b.id));
  comments.modified.sort((a, b) => compareStrings(a.id, b.id));

  /** @type {MetadataDiff} */
  const metadata = { added: [], removed: [], modified: [] };
  for (const [key, value] of after.metadata) {
    if (!before.metadata.has(key)) {
      metadata.added.push({ key, value });
    }
  }
  for (const [key, value] of before.metadata) {
    if (!after.metadata.has(key)) {
      metadata.removed.push({ key, value });
    }
  }
  for (const [key, afterValue] of after.metadata) {
    const beforeValue = before.metadata.get(key);
    if (beforeValue === undefined) continue;
    if (!deepEqual(beforeValue, afterValue)) {
      metadata.modified.push({ key, before: beforeValue, after: afterValue });
    }
  }
  metadata.added.sort((a, b) => compareStrings(a.key, b.key));
  metadata.removed.sort((a, b) => compareStrings(a.key, b.key));
  metadata.modified.sort((a, b) => compareStrings(a.key, b.key));

  /** @type {NamedRangesDiff} */
  const namedRanges = { added: [], removed: [], modified: [] };
  for (const [key, value] of after.namedRanges) {
    if (!before.namedRanges.has(key)) {
      namedRanges.added.push({ key, value });
    }
  }
  for (const [key, value] of before.namedRanges) {
    if (!after.namedRanges.has(key)) {
      namedRanges.removed.push({ key, value });
    }
  }
  for (const [key, afterValue] of after.namedRanges) {
    const beforeValue = before.namedRanges.get(key);
    if (beforeValue === undefined) continue;
    if (!deepEqual(beforeValue, afterValue)) {
      namedRanges.modified.push({ key, before: beforeValue, after: afterValue });
    }
  }

  namedRanges.added.sort((a, b) => compareStrings(a.key, b.key));
  namedRanges.removed.sort((a, b) => compareStrings(a.key, b.key));
  namedRanges.modified.sort((a, b) => compareStrings(a.key, b.key));

  return { sheets, cellsBySheet, comments, metadata, namedRanges };
}
