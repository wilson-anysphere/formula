import {
  cellContentEquivalent,
  cellsEqual,
  deepEqual,
  didContentChange,
  didFormatChange,
  normalizeCell
} from "./cell.js";
import { applyCellMovesToBaseSheet, applyCellMovesToSheet, detectCellMoves } from "./moves.js";
import { normalizeDocumentState } from "./state.js";

/**
 * @typedef {import("./types.js").Cell} Cell
 * @typedef {import("./types.js").CellMap} CellMap
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").MergeConflict} MergeConflict
 * @typedef {import("./types.js").MergeResult} MergeResult
 * @typedef {import("./types.js").SheetMeta} SheetMeta
 */

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

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
 * @param {Record<string, any>} baseRecord
 * @param {Record<string, any>} oursRecord
 * @param {Record<string, any>} theirsRecord
 * @param {{ conflictType: "metadata" | "namedRange" | "comment", keyField: string }} opts
 * @returns {{ merged: Record<string, any>, conflicts: MergeConflict[] }}
 */
function mergeKeyedRecords(baseRecord, oursRecord, theirsRecord, opts) {
  /** @type {MergeConflict[]} */
  const conflicts = [];

  /** @type {Record<string, any>} */
  const merged = {};

  const keys = new Set([
    ...Object.keys(baseRecord ?? {}),
    ...Object.keys(oursRecord ?? {}),
    ...Object.keys(theirsRecord ?? {}),
  ]);

  for (const key of keys) {
    const baseVal = baseRecord?.[key];
    const oursVal = oursRecord?.[key];
    const theirsVal = theirsRecord?.[key];

    const baseNorm = baseVal === undefined ? null : baseVal;
    const oursNorm = oursVal === undefined ? null : oursVal;
    const theirsNorm = theirsVal === undefined ? null : theirsVal;

    if (deepEqual(oursNorm, theirsNorm)) {
      if (oursVal !== undefined) merged[key] = structuredClone(oursVal);
      continue;
    }

    if (deepEqual(baseNorm, oursNorm)) {
      if (theirsVal !== undefined) merged[key] = structuredClone(theirsVal);
      continue;
    }

    if (deepEqual(baseNorm, theirsNorm)) {
      if (oursVal !== undefined) merged[key] = structuredClone(oursVal);
      continue;
    }

    conflicts.push({
      type: opts.conflictType,
      // @ts-expect-error - dynamic field name for conflict key/id.
      [opts.keyField]: key,
      base: baseNorm,
      ours: oursNorm,
      theirs: theirsNorm,
    });

    // Default to ours when unresolved.
    if (oursVal !== undefined) merged[key] = structuredClone(oursVal);
  }

  return { merged, conflicts };
}

/**
 * @param {string[]} order
 * @param {Record<string, SheetMeta>} metaById
 */
function normalizeOrder(order, metaById) {
  /** @type {string[]} */
  const out = [];
  const seen = new Set();
  for (const id of order) {
    if (typeof id !== "string" || id.length === 0) continue;
    if (seen.has(id)) continue;
    if (!metaById[id]) continue;
    out.push(id);
    seen.add(id);
  }
  for (const id of Object.keys(metaById)) {
    if (seen.has(id)) continue;
    out.push(id);
    seen.add(id);
  }
  return out;
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
 * Return the minimal set of sheet ids that changed relative order between `baseOrder`
 * and `afterOrder`.
 *
 * @param {string[]} baseOrder
 * @param {string[]} afterOrder
 * @returns {Set<string>}
 */
function movedSheetIds(baseOrder, afterOrder) {
  const afterIds = new Set(afterOrder);
  const baseIds = new Set(baseOrder);

  const baseCommon = baseOrder.filter((id) => afterIds.has(id));
  const afterCommon = afterOrder.filter((id) => baseIds.has(id));

  /** @type {Map<string, number>} */
  const afterIndex = new Map();
  afterCommon.forEach((id, idx) => {
    if (!afterIndex.has(id)) afterIndex.set(id, idx);
  });

  const seq = baseCommon.map((id) => afterIndex.get(id) ?? -1);
  const lisIdx = new Set(longestIncreasingSubsequenceIndices(seq));

  const moved = new Set();
  for (let i = 0; i < baseCommon.length; i += 1) {
    if (lisIdx.has(i)) continue;
    moved.add(baseCommon[i]);
  }
  return moved;
}

/**
 * Apply ordering constraints from `desiredOrder` for the given `movingIds` onto
 * `currentOrder`, preserving the relative order of all other ids.
 *
 * @param {string[]} currentOrder
 * @param {string[]} desiredOrder
 * @param {Iterable<string>} movingIds
 * @returns {{ order: string[], conflict: boolean }}
 */
function applyOrderConstraints(currentOrder, desiredOrder, movingIds) {
  const desiredIndex = new Map();
  desiredOrder.forEach((id, idx) => {
    if (!desiredIndex.has(id)) desiredIndex.set(id, idx);
  });

  const moving = Array.from(new Set(Array.from(movingIds))).filter((id) => typeof id === "string");
  moving.sort((a, b) => (desiredIndex.get(a) ?? Infinity) - (desiredIndex.get(b) ?? Infinity));

  const movingSet = new Set(moving);
  /** @type {string[]} */
  const out = currentOrder.filter((id) => !movingSet.has(id));

  for (const id of moving) {
    const idxId = desiredIndex.get(id);
    if (idxId === undefined) {
      out.push(id);
      continue;
    }

    let maxBefore = -1;
    let minAfter = out.length;

    for (let i = 0; i < out.length; i += 1) {
      const other = out[i];
      const idxOther = desiredIndex.get(other);
      if (idxOther === undefined) continue;
      if (idxOther < idxId) maxBefore = Math.max(maxBefore, i);
      if (idxOther > idxId) minAfter = Math.min(minAfter, i);
    }

    const insertAt = maxBefore + 1;
    if (insertAt > minAfter) {
      return { order: currentOrder.slice(), conflict: true };
    }

    out.splice(insertAt, 0, id);
  }

  return { order: out, conflict: false };
}

/**
 * Merge three sheet orderings. Prefers ours when conflicts occur.
 *
 * @param {{ base: string[], ours: string[], theirs: string[], metaById: Record<string, SheetMeta> }} input
 * @returns {{ order: string[], conflict: boolean }}
 */
function mergeSheetOrder({ base, ours, theirs, metaById }) {
  const sheetIds = new Set(Object.keys(metaById));

  const baseOrder = base.filter((id) => sheetIds.has(id));
  const oursOrder = ours.filter((id) => sheetIds.has(id));
  const theirsOrder = theirs.filter((id) => sheetIds.has(id));

  if (deepEqual(oursOrder, theirsOrder)) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: false };
  }

  // Fast paths: only one side changed the base ordering.
  if (deepEqual(baseOrder, oursOrder)) {
    return { order: normalizeOrder(theirsOrder, metaById), conflict: false };
  }
  if (deepEqual(baseOrder, theirsOrder)) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: false };
  }

  const movedOurs = movedSheetIds(baseOrder, oursOrder);
  const movedTheirs = movedSheetIds(baseOrder, theirsOrder);

  for (const id of movedOurs) {
    if (movedTheirs.has(id)) {
      return { order: normalizeOrder(oursOrder, metaById), conflict: true };
    }
  }

  let current = baseOrder.slice();
  const oursApplied = applyOrderConstraints(current, oursOrder, movedOurs);
  if (oursApplied.conflict) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: true };
  }
  current = oursApplied.order;

  const theirsApplied = applyOrderConstraints(current, theirsOrder, movedTheirs);
  if (theirsApplied.conflict) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: true };
  }
  current = theirsApplied.order;

  // Insert any sheets that didn't exist in base (added sheets).
  const baseIds = new Set(baseOrder);
  const oursAdded = oursOrder.filter((id) => !baseIds.has(id));
  const theirsAdded = theirsOrder.filter((id) => !baseIds.has(id) && !oursAdded.includes(id));

  const oursAddedApplied = applyOrderConstraints(current, oursOrder, oursAdded);
  if (oursAddedApplied.conflict) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: true };
  }
  current = oursAddedApplied.order;

  const theirsAddedApplied = applyOrderConstraints(current, theirsOrder, theirsAdded);
  if (theirsAddedApplied.conflict) {
    return { order: normalizeOrder(oursOrder, metaById), conflict: true };
  }
  current = theirsAddedApplied.order;

  return { order: normalizeOrder(current, metaById), conflict: false };
}

/**
 * @param {{ meta: SheetMeta, cells: CellMap } | null} a
 * @param {{ meta: SheetMeta, cells: CellMap } | null} b
 */
function sheetStateEqual(a, b) {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  if (!deepEqual(a.meta, b.meta)) return false;
  return deepEqual(a.cells, b.cells);
}

/**
 * Performs a 3-way semantic merge of spreadsheet document states.
 *
 * @param {{ base: DocumentState, ours: DocumentState, theirs: DocumentState }} input
 * @returns {MergeResult}
 */
export function mergeDocumentStates({ base, ours, theirs }) {
  const baseState = normalizeDocumentState(base);
  const oursState = normalizeDocumentState(ours);
  const theirsState = normalizeDocumentState(theirs);

  /** @type {MergeConflict[]} */
  const conflicts = [];

  // --- Workbook-level maps ---
  const metadata = mergeKeyedRecords(
    baseState.metadata ?? {},
    oursState.metadata ?? {},
    theirsState.metadata ?? {},
    { conflictType: "metadata", keyField: "key" },
  );
  conflicts.push(...metadata.conflicts);

  const namedRanges = mergeKeyedRecords(
    baseState.namedRanges ?? {},
    oursState.namedRanges ?? {},
    theirsState.namedRanges ?? {},
    { conflictType: "namedRange", keyField: "key" },
  );
  conflicts.push(...namedRanges.conflicts);

  const comments = mergeKeyedRecords(
    baseState.comments ?? {},
    oursState.comments ?? {},
    theirsState.comments ?? {},
    { conflictType: "comment", keyField: "id" },
  );
  conflicts.push(...comments.conflicts);

  // --- Sheets: presence + rename ---

  const allSheetIds = new Set([
    ...Object.keys(baseState.sheets.metaById ?? {}),
    ...Object.keys(oursState.sheets.metaById ?? {}),
    ...Object.keys(theirsState.sheets.metaById ?? {}),
  ]);

  /** @type {Record<string, SheetMeta>} */
  const metaById = {};

  /** @type {Map<string, "merge" | "ours" | "theirs">} */
  const cellStrategy = new Map();

  for (const sheetId of allSheetIds) {
    const baseMeta = baseState.sheets.metaById[sheetId] ?? null;
    const oursMeta = oursState.sheets.metaById[sheetId] ?? null;
    const theirsMeta = theirsState.sheets.metaById[sheetId] ?? null;

    const baseSheet = baseMeta ? baseState.cells[sheetId] ?? {} : {};
    const oursSheet = oursMeta ? oursState.cells[sheetId] ?? {} : {};
    const theirsSheet = theirsMeta ? theirsState.cells[sheetId] ?? {} : {};

    /** @type {{ meta: SheetMeta, cells: CellMap } | null} */
    const baseSheetState = baseMeta ? { meta: baseMeta, cells: baseSheet } : null;
    /** @type {{ meta: SheetMeta, cells: CellMap } | null} */
    const oursSheetState = oursMeta ? { meta: oursMeta, cells: oursSheet } : null;
    /** @type {{ meta: SheetMeta, cells: CellMap } | null} */
    const theirsSheetState = theirsMeta ? { meta: theirsMeta, cells: theirsSheet } : null;

    // Added sheets (base absent).
    if (!baseMeta) {
      if (oursMeta && !theirsMeta) {
        metaById[sheetId] = { id: sheetId, name: oursMeta.name ?? null };
        cellStrategy.set(sheetId, "ours");
        continue;
      }
      if (!oursMeta && theirsMeta) {
        metaById[sheetId] = { id: sheetId, name: theirsMeta.name ?? null };
        cellStrategy.set(sheetId, "theirs");
        continue;
      }
      if (oursMeta && theirsMeta) {
        if (!deepEqual(oursMeta.name ?? null, theirsMeta.name ?? null)) {
          conflicts.push({
            type: "sheet",
            reason: "rename",
            sheetId,
            base: null,
            ours: oursMeta.name ?? null,
            theirs: theirsMeta.name ?? null,
          });
        }
        metaById[sheetId] = { id: sheetId, name: oursMeta.name ?? null };
        cellStrategy.set(sheetId, "merge");
      }
      continue;
    }

    // Removed on both sides.
    if (!oursMeta && !theirsMeta) continue;

    // Delete vs keep.
    if (!oursMeta && theirsMeta) {
      // ours deleted.
      if (sheetStateEqual(baseSheetState, theirsSheetState)) {
        // theirs unchanged => safe delete.
        continue;
      }
      // Conflict: deletion vs edits.
      conflicts.push({
        type: "sheet",
        reason: "presence",
        sheetId,
        base: baseSheetState,
        ours: null,
        theirs: theirsSheetState,
      });
      // Prefer ours: keep deletion.
      continue;
    }

    if (oursMeta && !theirsMeta) {
      // theirs deleted.
      if (sheetStateEqual(baseSheetState, oursSheetState)) {
        // ours unchanged => safe delete (take theirs).
        continue;
      }
      conflicts.push({
        type: "sheet",
        reason: "presence",
        sheetId,
        base: baseSheetState,
        ours: oursSheetState,
        theirs: null,
      });
      // Prefer ours: keep sheet as-is.
      metaById[sheetId] = { id: sheetId, name: oursMeta.name ?? null };
      cellStrategy.set(sheetId, "ours");
      continue;
    }

    // Present on both sides: merge rename + cells.
    if (oursMeta && theirsMeta) {
      const baseName = baseMeta.name ?? null;
      const oursName = oursMeta.name ?? null;
      const theirsName = theirsMeta.name ?? null;

      let mergedName = oursName;
      if (deepEqual(oursName, theirsName)) {
        mergedName = oursName;
      } else if (deepEqual(baseName, oursName)) {
        mergedName = theirsName;
      } else if (deepEqual(baseName, theirsName)) {
        mergedName = oursName;
      } else {
        conflicts.push({
          type: "sheet",
          reason: "rename",
          sheetId,
          base: baseName,
          ours: oursName,
          theirs: theirsName,
        });
        mergedName = oursName;
      }

      metaById[sheetId] = { id: sheetId, name: mergedName };
      cellStrategy.set(sheetId, "merge");
    }
  }

  // --- Sheet order ---
  const orderMerge = mergeSheetOrder({
    base: baseState.sheets.order ?? [],
    ours: oursState.sheets.order ?? [],
    theirs: theirsState.sheets.order ?? [],
    metaById,
  });

  if (orderMerge.conflict) {
    conflicts.push({
      type: "sheet",
      reason: "order",
      base: baseState.sheets.order ?? [],
      ours: oursState.sheets.order ?? [],
      theirs: theirsState.sheets.order ?? [],
    });
  }

  /** @type {Record<string, CellMap>} */
  const cells = {};

  for (const sheetId of Object.keys(metaById)) {
    const strategy = cellStrategy.get(sheetId) ?? "ours";
    if (strategy === "ours") {
      cells[sheetId] = structuredClone(oursState.cells[sheetId] ?? {});
      continue;
    }
    if (strategy === "theirs") {
      cells[sheetId] = structuredClone(theirsState.cells[sheetId] ?? {});
      continue;
    }

    const baseSheetOriginal = baseState.cells[sheetId] ?? {};
    const oursSheetOriginal = oursState.cells[sheetId] ?? {};
    const theirsSheetOriginal = theirsState.cells[sheetId] ?? {};

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
          theirs: { to: theirsTo },
        });
        combinedMoves.set(from, oursTo);
        continue;
      }

      combinedMoves.set(from, oursTo ?? theirsTo);
    }

    const baseSheet = applyCellMovesToBaseSheet(baseSheetOriginal, combinedMoves);
    const oursSheet = applyCellMovesToSheet(baseSheetOriginal, oursSheetOriginal, combinedMoves);
    const theirsSheet = applyCellMovesToSheet(baseSheetOriginal, theirsSheetOriginal, combinedMoves);

    const cellAddrs = new Set([
      ...Object.keys(baseSheet),
      ...Object.keys(oursSheet),
      ...Object.keys(theirsSheet),
    ]);

    /** @type {CellMap} */
    const mergedSheet = {};

    for (const cellAddr of cellAddrs) {
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

    cells[sheetId] = mergedSheet;
  }

  /** @type {DocumentState} */
  const merged = {
    schemaVersion: 1,
    sheets: { order: orderMerge.order, metaById },
    cells,
    metadata: metadata.merged,
    namedRanges: namedRanges.merged,
    comments: comments.merged,
  };

  return { merged: normalizeDocumentState(merged), conflicts };
}

/**
 * @typedef {{
 *   conflictIndex: number,
 *   choice: "ours" | "theirs" | "manual",
 *   manualCell?: Cell | null,
 *   manualMoveTo?: string,
 *   manualSheetName?: string | null,
 *   manualSheetOrder?: string[],
 *   manualMetadataValue?: any,
 *   manualNamedRangeValue?: any,
 *   manualCommentValue?: any,
 *   manualSheetState?: any
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
  const merged = structuredClone(normalizeDocumentState(mergeResult.merged));

  for (const resolution of resolutions) {
    const conflict = mergeResult.conflicts[resolution.conflictIndex];
    if (!conflict) throw new Error(`Unknown conflict index ${resolution.conflictIndex}`);

    if (conflict.type === "cell") {
      const sheet = merged.cells[conflict.sheetId] ?? {};
      merged.cells[conflict.sheetId] = sheet;

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
      const sheet = merged.cells[conflict.sheetId] ?? {};
      merged.cells[conflict.sheetId] = sheet;

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
      continue;
    }

    if (conflict.type === "sheet") {
      if (conflict.reason === "rename") {
        const sheetId = conflict.sheetId;
        if (!sheetId) throw new Error("Sheet rename conflict requires sheetId");

        let name;
        if (resolution.choice === "ours") name = conflict.ours ?? null;
        else if (resolution.choice === "theirs") name = conflict.theirs ?? null;
        else name = resolution.manualSheetName ?? null;

        if (!merged.sheets.metaById[sheetId]) {
          merged.sheets.metaById[sheetId] = { id: sheetId, name: name ?? null };
        } else {
          merged.sheets.metaById[sheetId].name = name ?? null;
        }
        continue;
      }

      if (conflict.reason === "order") {
        let order;
        if (resolution.choice === "ours") order = conflict.ours;
        else if (resolution.choice === "theirs") order = conflict.theirs;
        else order = resolution.manualSheetOrder;

        if (!Array.isArray(order)) throw new Error("Sheet order conflict resolution requires an array order");
        merged.sheets.order = order.filter((id) => typeof id === "string");
        continue;
      }

      if (conflict.reason === "presence") {
        const sheetId = conflict.sheetId;
        if (!sheetId) throw new Error("Sheet presence conflict requires sheetId");

        let chosen;
        if (resolution.choice === "ours") chosen = conflict.ours;
        else if (resolution.choice === "theirs") chosen = conflict.theirs;
        else chosen = resolution.manualSheetState;

        if (chosen === null) {
          delete merged.sheets.metaById[sheetId];
          delete merged.cells[sheetId];
          merged.sheets.order = merged.sheets.order.filter((id) => id !== sheetId);
          continue;
        }

        if (!isRecord(chosen) || !isRecord(chosen.meta) || !isRecord(chosen.cells)) {
          throw new Error("Sheet presence conflict manual resolution requires { meta, cells } or null");
        }

        merged.sheets.metaById[sheetId] = {
          id: sheetId,
          name: chosen.meta.name == null ? null : String(chosen.meta.name),
        };
        merged.cells[sheetId] = structuredClone(chosen.cells);
        if (!merged.sheets.order.includes(sheetId)) merged.sheets.order.push(sheetId);
        continue;
      }
    }

    if (conflict.type === "metadata") {
      const key = conflict.key;
      let value;
      if (resolution.choice === "ours") value = conflict.ours;
      else if (resolution.choice === "theirs") value = conflict.theirs;
      else value = resolution.manualMetadataValue;

      if (value === null || value === undefined) delete merged.metadata[key];
      else merged.metadata[key] = value;
      continue;
    }

    if (conflict.type === "namedRange") {
      const key = conflict.key;
      let value;
      if (resolution.choice === "ours") value = conflict.ours;
      else if (resolution.choice === "theirs") value = conflict.theirs;
      else value = resolution.manualNamedRangeValue;

      if (value === null || value === undefined) delete merged.namedRanges[key];
      else merged.namedRanges[key] = value;
      continue;
    }

    if (conflict.type === "comment") {
      const id = conflict.id;
      let value;
      if (resolution.choice === "ours") value = conflict.ours;
      else if (resolution.choice === "theirs") value = conflict.theirs;
      else value = resolution.manualCommentValue;

      if (value === null || value === undefined) delete merged.comments[id];
      else merged.comments[id] = value;
      continue;
    }
  }

  return normalizeDocumentState(merged);
}
