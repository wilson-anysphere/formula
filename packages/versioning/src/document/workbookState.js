import { cellKey } from "../diff/semanticDiff.js";

const decoder = new TextDecoder();

/**
 * @typedef {{ id: string, name: string | null }} SheetMeta
 * @typedef {{ id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }} CommentSummary
 *
 * @typedef {{
 *   sheets: SheetMeta[];
 *   namedRanges: Map<string, any>;
 *   comments: Map<string, CommentSummary>;
 *   cellsBySheet: Map<string, { cells: Map<string, any> }>;
 * }} WorkbookState
 */

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * @param {any} value
 * @returns {boolean}
 */
function coerceBool(value) {
  return Boolean(value);
}

/**
 * @param {any} value
 * @returns {number}
 */
function repliesLength(value) {
  if (Array.isArray(value)) return value.length;
  return 0;
}

/**
 * @param {any} value
 * @param {string} fallbackId
 * @returns {CommentSummary}
 */
function commentSummaryFromValue(value, fallbackId) {
  const id = coerceString(value?.id) ?? fallbackId;
  const cellRef = coerceString(value?.cellRef);
  const content = coerceString(value?.content);
  const resolved = coerceBool(value?.resolved);
  return { id, cellRef, content, resolved, repliesLength: repliesLength(value?.replies) };
}

/**
 * @param {Uint8Array} snapshot
 * @returns {WorkbookState}
 */
export function workbookStateFromDocumentSnapshot(snapshot) {
  let parsed;
  try {
    parsed = JSON.parse(decoder.decode(snapshot));
  } catch {
    throw new Error("Invalid document snapshot: not valid JSON");
  }

  const sheetsList = Array.isArray(parsed?.sheets) ? parsed.sheets : [];

  /** @type {SheetMeta[]} */
  const sheets = [];
  /** @type {Map<string, { cells: Map<string, any> }>} */
  const cellsBySheet = new Map();

  for (const sheet of sheetsList) {
    const id = coerceString(sheet?.id);
    if (!id) continue;
    const name = coerceString(sheet?.name);
    sheets.push({ id, name });

    /** @type {Map<string, any>} */
    const cells = new Map();
    const entries = Array.isArray(sheet?.cells) ? sheet.cells : [];
    for (const entry of entries) {
      const row = Number(entry?.row);
      const col = Number(entry?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      cells.set(cellKey(row, col), {
        value: entry?.value ?? null,
        formula: entry?.formula ?? null,
        format: entry?.format ?? null,
      });
    }

    cellsBySheet.set(id, { cells });
  }

  sheets.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  /** @type {Map<string, any>} */
  const namedRanges = new Map();
  const rawNamedRanges = parsed?.namedRanges;
  if (rawNamedRanges && typeof rawNamedRanges === "object") {
    if (Array.isArray(rawNamedRanges)) {
      for (const entry of rawNamedRanges) {
        const key = coerceString(entry?.name ?? entry?.id);
        if (!key) continue;
        namedRanges.set(key, structuredClone(entry));
      }
    } else {
      const keys = Object.keys(rawNamedRanges).sort();
      for (const key of keys) {
        namedRanges.set(key, structuredClone(rawNamedRanges[key]));
      }
    }
  }

  /** @type {Map<string, CommentSummary>} */
  const comments = new Map();
  const rawComments = parsed?.comments;
  if (rawComments && typeof rawComments === "object") {
    if (Array.isArray(rawComments)) {
      for (const entry of rawComments) {
        const id = coerceString(entry?.id);
        if (!id) continue;
        comments.set(id, commentSummaryFromValue(entry, id));
      }
    } else {
      const keys = Object.keys(rawComments).sort();
      for (const id of keys) {
        comments.set(id, commentSummaryFromValue(rawComments[id], id));
      }
    }
  }

  return { sheets, namedRanges, comments, cellsBySheet };
}

