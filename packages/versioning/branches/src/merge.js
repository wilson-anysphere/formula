import {
  cellContentEquivalent,
  cellsEqual,
  deepEqual,
  didContentChange,
  didFormatChange,
  normalizeCell
} from "./cell.js";
import { applyCellMovesToBaseSheet, applyCellMovesToSheet, detectCellMoves } from "./moves.js";

/**
 * @typedef {import("./types.js").Cell} Cell
 * @typedef {import("./types.js").CellMap} CellMap
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").MergeConflict} MergeConflict
 * @typedef {import("./types.js").MergeResult} MergeResult
 */

/**
 * @param {Record<string, any> | null} baseFormat
 * @param {Record<string, any> | null} oursFormat
 * @param {Record<string, any> | null} theirsFormat
 * @returns {{ merged: Record<string, any> | null, conflict: boolean }}
 */
function mergeFormats(baseFormat, oursFormat, theirsFormat) {
  const baseObj = baseFormat ?? null;
  const oursObj = oursFormat ?? null;
  const theirsObj = theirsFormat ?? null;

  if (deepEqual(oursObj, theirsObj)) return { merged: oursObj, conflict: false };
  if (deepEqual(baseObj, oursObj)) return { merged: theirsObj, conflict: false };
  if (deepEqual(baseObj, theirsObj)) return { merged: oursObj, conflict: false };

  const keys = new Set([
    ...Object.keys(baseObj ?? {}),
    ...Object.keys(oursObj ?? {}),
    ...Object.keys(theirsObj ?? {})
  ]);

  /** @type {Record<string, any>} */
  const merged = {};

  for (const key of keys) {
    const baseVal = baseObj?.[key];
    const oursVal = oursObj?.[key];
    const theirsVal = theirsObj?.[key];

    if (deepEqual(oursVal, theirsVal)) {
      if (oursVal !== undefined) merged[key] = oursVal;
      continue;
    }

    if (deepEqual(baseVal, oursVal)) {
      if (theirsVal !== undefined) merged[key] = theirsVal;
      continue;
    }

    if (deepEqual(baseVal, theirsVal)) {
      if (oursVal !== undefined) merged[key] = oursVal;
      continue;
    }

    return { merged: null, conflict: true };
  }

  return { merged: Object.keys(merged).length === 0 ? null : merged, conflict: false };
}

/**
 * @param {Cell | null} base
 * @param {Cell | null} ours
 * @param {Cell | null} theirs
 * @returns {{ merged: Cell | null, conflict: MergeConflict | null }}
 */
function mergeCell(base, ours, theirs) {
  const nBase = normalizeCell(base);
  const nOurs = normalizeCell(ours);
  const nTheirs = normalizeCell(theirs);

  if (cellsEqual(nOurs, nTheirs)) {
    return { merged: nOurs, conflict: null };
  }

  if (cellsEqual(nBase, nOurs)) return { merged: nTheirs, conflict: null };
  if (cellsEqual(nBase, nTheirs)) return { merged: nOurs, conflict: null };

  const contentChangedOurs = didContentChange(nBase, nOurs);
  const contentChangedTheirs = didContentChange(nBase, nTheirs);
  const formatChangedOurs = didFormatChange(nBase, nOurs);
  const formatChangedTheirs = didFormatChange(nBase, nTheirs);

  // Deletion vs edit is a specialized content conflict.
  if (
    nBase !== null &&
    ((nOurs === null && nTheirs !== null) || (nTheirs === null && nOurs !== null)) &&
    (contentChangedOurs || contentChangedTheirs)
  ) {
    return {
      merged: nOurs,
      conflict: {
        type: "cell",
        sheetId: "",
        cell: "",
        reason: "delete-vs-edit",
        base: nBase,
        ours: nOurs,
        theirs: nTheirs
      }
    };
  }

  /** @type {Cell | null} */
  let mergedContent = nBase;
  /** @type {MergeConflict | null} */
  let contentConflict = null;

  if (contentChangedOurs || contentChangedTheirs) {
    if (!contentChangedOurs) mergedContent = nTheirs;
    else if (!contentChangedTheirs) mergedContent = nOurs;
    else if (cellContentEquivalent(nOurs, nTheirs)) mergedContent = nOurs;
    else {
      contentConflict = {
        type: "cell",
        sheetId: "",
        cell: "",
        reason: "content",
        base: nBase,
        ours: nOurs,
        theirs: nTheirs
      };
    }
  }

  const baseFormat = nBase?.format ?? null;
  const oursFormat = nOurs?.format ?? null;
  const theirsFormat = nTheirs?.format ?? null;
  const mergedFormat = mergeFormats(baseFormat, oursFormat, theirsFormat);

  if (mergedFormat.conflict) {
    const conflict = {
      type: "cell",
      sheetId: "",
      cell: "",
      reason: "format",
      base: nBase,
      ours: nOurs,
      theirs: nTheirs
    };
    // If both content and format conflicted, preserve the content conflict
    // semantics (the user will still need to resolve per-cell).
    return { merged: nOurs, conflict: contentConflict ?? conflict };
  }

  if (contentConflict) {
    return { merged: nOurs, conflict: contentConflict };
  }

  if (mergedContent === null && mergedFormat.merged === null) return { merged: null, conflict: null };

  /** @type {Cell} */
  const mergedCell = {};

  if (mergedContent?.formula !== undefined) mergedCell.formula = mergedContent.formula;
  else if (mergedContent?.value !== undefined) mergedCell.value = mergedContent.value;

  if (mergedFormat.merged) mergedCell.format = mergedFormat.merged;

  if (Object.keys(mergedCell).length === 0) return { merged: null, conflict: null };

  return { merged: mergedCell, conflict: null };
}

/**
 * Performs a 3-way semantic merge of spreadsheet document states.
 *
 * @param {{ base: DocumentState, ours: DocumentState, theirs: DocumentState }} input
 * @returns {MergeResult}
 */
export function mergeDocumentStates({ base, ours, theirs }) {
  /** @type {MergeConflict[]} */
  const conflicts = [];

  /** @type {DocumentState} */
  const merged = { sheets: {} };

  const sheetIds = new Set([
    ...Object.keys(base.sheets ?? {}),
    ...Object.keys(ours.sheets ?? {}),
    ...Object.keys(theirs.sheets ?? {})
  ]);

  for (const sheetId of sheetIds) {
    const baseSheetOriginal = base.sheets[sheetId] ?? {};
    const oursSheetOriginal = ours.sheets[sheetId] ?? {};
    const theirsSheetOriginal = theirs.sheets[sheetId] ?? {};

    const oursMoves = detectCellMoves(baseSheetOriginal, oursSheetOriginal);
    const theirsMoves = detectCellMoves(baseSheetOriginal, theirsSheetOriginal);

    /** @type {Map<string, string>} */
    const combinedMoves = new Map();

    const movedFrom = new Set([...oursMoves.keys(), ...theirsMoves.keys()]);

    for (const from of movedFrom) {
      const oursTo = oursMoves.get(from);
      const theirsTo = theirsMoves.get(from);

      if (oursTo && theirsTo && oursTo !== theirsTo) {
        conflicts.push({
          type: "move",
          sheetId,
          cell: from,
          reason: "move-destination",
          base: normalizeCell(baseSheetOriginal[from]),
          ours: { to: oursTo },
          theirs: { to: theirsTo }
        });
        combinedMoves.set(from, oursTo);
        continue;
      }

      combinedMoves.set(from, oursTo ?? theirsTo);
    }

    const baseSheet = applyCellMovesToBaseSheet(baseSheetOriginal, combinedMoves);
    const oursSheet = applyCellMovesToSheet(baseSheetOriginal, oursSheetOriginal, combinedMoves);
    const theirsSheet = applyCellMovesToSheet(baseSheetOriginal, theirsSheetOriginal, combinedMoves);

    const cells = new Set([
      ...Object.keys(baseSheet),
      ...Object.keys(oursSheet),
      ...Object.keys(theirsSheet)
    ]);

    /** @type {CellMap} */
    const mergedSheet = {};

    for (const cellAddr of cells) {
      const baseCell = baseSheet[cellAddr];
      const oursCell = oursSheet[cellAddr];
      const theirsCell = theirsSheet[cellAddr];

      const mergedCellResult = mergeCell(baseCell, oursCell, theirsCell);

      if (mergedCellResult.conflict) {
        mergedCellResult.conflict.sheetId = sheetId;
        mergedCellResult.conflict.cell = cellAddr;
        conflicts.push(mergedCellResult.conflict);
      }

      if (mergedCellResult.merged !== null) {
        mergedSheet[cellAddr] = mergedCellResult.merged;
      }
    }

    merged.sheets[sheetId] = mergedSheet;
  }

  return { merged, conflicts };
}

/**
 * @typedef {{
 *   conflictIndex: number,
 *   choice: "ours" | "theirs" | "manual",
 *   manualCell?: Cell | null,
 *   manualMoveTo?: string
 * }} ConflictResolution
 */

/**
 * Applies conflict resolutions to a merge result, producing a final merged
 * state.
 *
 * @param {MergeResult} mergeResult
 * @param {ConflictResolution[]} resolutions
 * @returns {DocumentState}
 */
export function applyConflictResolutions(mergeResult, resolutions) {
  const merged = structuredClone(mergeResult.merged);

  for (const resolution of resolutions) {
    const conflict = mergeResult.conflicts[resolution.conflictIndex];
    if (!conflict) throw new Error(`Unknown conflict index ${resolution.conflictIndex}`);

    if (conflict.type === "cell") {
      const sheet = merged.sheets[conflict.sheetId] ?? {};
      merged.sheets[conflict.sheetId] = sheet;

      /** @type {Cell | null} */
      let finalCell = null;
      if (resolution.choice === "ours") finalCell = conflict.ours;
      else if (resolution.choice === "theirs") finalCell = conflict.theirs;
      else finalCell = resolution.manualCell ?? null;

      if (finalCell === null) delete sheet[conflict.cell];
      else sheet[conflict.cell] = finalCell;

      continue;
    }

    if (conflict.type === "move") {
      const sheet = merged.sheets[conflict.sheetId] ?? {};
      merged.sheets[conflict.sheetId] = sheet;

      const oursTo = conflict.ours?.to;
      const theirsTo = conflict.theirs?.to;
      const baseCell = conflict.base;

      let target;
      if (resolution.choice === "ours") target = oursTo;
      else if (resolution.choice === "theirs") target = theirsTo;
      else target = resolution.manualMoveTo;

      if (!target) throw new Error("Move conflict resolution requires a destination");

      // Clear both destination locations.
      if (oursTo) delete sheet[oursTo];
      if (theirsTo) delete sheet[theirsTo];
      // Ensure source is cleared.
      delete sheet[conflict.cell];

      if (baseCell) sheet[target] = baseCell;
    }
  }

  return merged;
}
