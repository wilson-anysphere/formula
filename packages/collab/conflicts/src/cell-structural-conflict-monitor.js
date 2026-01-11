import * as Y from "yjs";
import { cellRefFromKey, numberToCol } from "./cell-ref.js";
 
/**
 * @typedef {object} NormalizedCell
 * @property {unknown} [value]
 * @property {string} [formula]
 * @property {Record<string, unknown> | null} [format]
 */
 
/**
 * @typedef {object} CellStructuralConflict
 * @property {string} id
 * @property {"move" | "cell"} type
 * @property {"move-destination" | "delete-vs-edit" | "content" | "format"} reason
 * @property {string} sheetId
 * @property {string} cell Address within the sheet (e.g. "A1"). For `type: "move"`, this is the source cell.
 * @property {string} cellKey Full Yjs key (e.g. "Sheet1:0:0") for the conflicting cell (source for moves).
 * @property {any} local
 * @property {any} remote
 * @property {string} remoteUserId
 * @property {number} detectedAt
 */
 
/**
 * @typedef {object} CellStructuralConflictResolution
 * @property {"ours" | "theirs" | "manual"} choice
 * @property {string} [to] Destination cellKey for resolving move conflicts.
 * @property {NormalizedCell | null} [cell] Cell value for manual cell resolutions.
 */
 
/**
 * Tracks structural cell conflicts (moves / delete-vs-edit) in a collaborative
 * Yjs spreadsheet document using causal state vectors.
 *
 * This monitor is deliberately conservative and focuses on single-cell moves
 * that can be inferred from a cut/paste transaction (delete at X + add at Y of
 * identical cell content).
 */
export class CellStructuralConflictMonitor {
  /**
   * @param {object} opts
   * @param {Y.Doc} opts.doc
   * @param {Y.Map<any>} [opts.cells]
   * @param {string} opts.localUserId
   * @param {any} [opts.origin]
   * @param {Set<any>} [opts.localOrigins]
   * @param {(conflict: CellStructuralConflict) => void} opts.onConflict
   */
  constructor(opts) {
    this.doc = opts.doc;
    this.cells = opts.cells ?? this.doc.getMap("cells");
    this.localUserId = opts.localUserId;
 
    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);
 
    this.onConflict = opts.onConflict;
 
    // Shared operation log used to exchange per-transaction causal metadata
    // (transaction.beforeState/afterState) between clients.
    this._ops = this.doc.getMap("cellStructuralOps");
 
    /** @type {Map<string, CellStructuralConflict>} */
    this._conflicts = new Map();
 
    /** @type {Map<string, any>} */
    this._opRecords = new Map();
 
    /** @type {Set<string>} */
    this._seenPairs = new Set();
 
    this._isApplyingResolution = false;
 
    this._onCellsDeepEvent = this._onCellsDeepEvent.bind(this);
    this.cells.observeDeep(this._onCellsDeepEvent);
 
    this._onOpsEvent = this._onOpsEvent.bind(this);
    this._ops.observe(this._onOpsEvent);
  }
 
  dispose() {
    this.cells.unobserveDeep(this._onCellsDeepEvent);
    this._ops.unobserve(this._onOpsEvent);
  }
 
  /** @returns {Array<CellStructuralConflict>} */
  listConflicts() {
    return Array.from(this._conflicts.values());
  }
 
  /**
   * Resolves a previously emitted conflict by writing the chosen state back
   * into the shared Yjs doc.
   *
   * @param {string} conflictId
   * @param {CellStructuralConflictResolution} resolution
   * @returns {boolean}
   */
  resolveConflict(conflictId, resolution) {
    const conflict = this._conflicts.get(conflictId);
    if (!conflict) return false;
 
    this._isApplyingResolution = true;
    try {
      if (conflict.type === "move") {
        const oursTo = conflict.local?.toCellKey ?? null;
        const theirsTo = conflict.remote?.toCellKey ?? null;
        const baseCell = conflict.local?.cell ?? conflict.remote?.cell ?? null;
 
        let target = null;
        if (resolution.choice === "ours") target = oursTo;
        else if (resolution.choice === "theirs") target = theirsTo;
        else target = resolution.to ?? null;
 
        if (!target) {
          throw new Error("Move conflict resolution requires a destination cellKey");
        }
 
        this.doc.transact(() => {
          // Clear source + both destinations.
          this.cells.delete(conflict.cellKey);
          if (oursTo) this.cells.delete(oursTo);
          if (theirsTo) this.cells.delete(theirsTo);
 
          if (baseCell) {
            this._writeCell(target, baseCell);
          }
        }, this.origin);
      } else {
        let finalCell = null;
        if (resolution.choice === "ours") finalCell = conflict.local?.after ?? null;
        else if (resolution.choice === "theirs") finalCell = conflict.remote?.after ?? null;
        else finalCell = resolution.cell ?? null;
 
        this.doc.transact(() => {
          if (finalCell === null) {
            this.cells.delete(conflict.cellKey);
          } else {
            this._writeCell(conflict.cellKey, finalCell);
          }
        }, this.origin);
      }
    } finally {
      this._isApplyingResolution = false;
    }
 
    this._conflicts.delete(conflictId);
    return true;
  }
 
  /**
   * @param {Array<any>} events
   * @param {Y.Transaction} transaction
   */
  _onCellsDeepEvent(events, transaction) {
    if (this._isApplyingResolution) return;
 
    const isLocal = this.localOrigins.has(transaction.origin);
    if (!isLocal) return;
 
    const tx = extractTransactionChanges(events, this.cells);
    if (!tx) return;
 
    const beforeState = encodeStateVector(transaction.beforeState);
    const afterState = encodeStateVector(transaction.afterState);
    const txId = crypto.randomUUID();
 
    /** @type {Array<any>} */
    const records = [];
 
    for (const move of tx.moves) {
      records.push({
        id: crypto.randomUUID(),
        txId,
        kind: "move",
        userId: this.localUserId,
        beforeState,
        afterState,
        touchedCells: tx.touchedCells,
        fromCellKey: move.fromCellKey,
        toCellKey: move.toCellKey,
        cell: move.cell,
        fingerprint: move.fingerprint
      });
    }
 
    for (const del of tx.deletes) {
      records.push({
        id: crypto.randomUUID(),
        txId,
        kind: "delete",
        userId: this.localUserId,
        beforeState,
        afterState,
        touchedCells: tx.touchedCells,
        cellKey: del.cellKey,
        before: del.before,
        after: null,
        fingerprint: del.fingerprint
      });
    }
 
    for (const edit of tx.edits) {
      records.push({
        id: crypto.randomUUID(),
        txId,
        kind: "edit",
        userId: this.localUserId,
        beforeState,
        afterState,
        touchedCells: tx.touchedCells,
        cellKey: edit.cellKey,
        before: edit.before,
        after: edit.after,
        beforeFingerprint: edit.beforeFingerprint,
        afterFingerprint: edit.afterFingerprint,
        contentChanged: edit.contentChanged,
        formatChanged: edit.formatChanged
      });
    }
 
    if (records.length === 0) return;
 
    // Store operation records in a shared log so other clients can compare
    // state vectors and detect true concurrency.
    this.doc.transact(() => {
      for (const record of records) {
        this._ops.set(record.id, record);
      }
    }, this.origin);
  }
 
  /**
   * @param {Y.YMapEvent<any>} event
   * @param {Y.Transaction} transaction
   */
  _onOpsEvent(event, transaction) {
    if (this._isApplyingResolution) return;
    if (!event?.changes?.keys) return;
 
    for (const [opId, change] of event.changes.keys.entries()) {
      if (change.action !== "add") continue;
      const record = this._ops.get(opId);
      if (!record) continue;
      this._ingestOpRecord(record);
    }
  }
 
  /**
   * @param {any} record
   */
  _ingestOpRecord(record) {
    if (!record || typeof record !== "object") return;
    const id = String(record.id ?? "");
    if (!id) return;
    if (this._opRecords.has(id)) return;
 
    this._opRecords.set(id, record);
 
    const ours = record.userId === this.localUserId;
 
    // Compare this operation against all known opposite-side operations.
    for (const other of this._opRecords.values()) {
      if (!other || typeof other !== "object") continue;
      if (other === record) continue;
      const otherIsOurs = other.userId === this.localUserId;
      if (otherIsOurs === ours) continue;
 
      const pairKey = makePairKey(id, String(other.id ?? ""));
      if (this._seenPairs.has(pairKey)) continue;
      this._seenPairs.add(pairKey);
 
      const oursOp = ours ? record : other;
      const theirsOp = ours ? other : record;
 
      this._handleOpPair(oursOp, theirsOp);
    }
  }
 
  /**
   * @param {any} oursOp
   * @param {any} theirsOp
   */
  _handleOpPair(oursOp, theirsOp) {
    if (!isCausallyConcurrent(oursOp, theirsOp)) return;
 
    // move-destination conflict (same source moved to different destinations).
    if (oursOp.kind === "move" && theirsOp.kind === "move") {
      if (oursOp.fromCellKey === theirsOp.fromCellKey && oursOp.toCellKey !== theirsOp.toCellKey) {
        this._emitConflict({
          type: "move",
          reason: "move-destination",
          sourceCellKey: oursOp.fromCellKey,
          local: oursOp,
          remote: theirsOp
        });
      }
      return;
    }
 
    // move vs edit auto-merge (rename-aware).
    if (oursOp.kind === "move" && theirsOp.kind === "edit") {
      this._maybeAutoMergeMoveEdit({ move: oursOp, edit: theirsOp });
      return;
    }
    if (oursOp.kind === "edit" && theirsOp.kind === "move") {
      this._maybeAutoMergeMoveEdit({ move: theirsOp, edit: oursOp });
      return;
    }
 
    // delete-vs-edit conflict.
    if (oursOp.kind === "delete" && theirsOp.kind === "edit" && oursOp.cellKey === theirsOp.cellKey) {
      this._emitConflict({
        type: "cell",
        reason: "delete-vs-edit",
        sourceCellKey: oursOp.cellKey,
        local: oursOp,
        remote: theirsOp
      });
      return;
    }
    if (oursOp.kind === "edit" && theirsOp.kind === "delete" && oursOp.cellKey === theirsOp.cellKey) {
      this._emitConflict({
        type: "cell",
        reason: "delete-vs-edit",
        sourceCellKey: oursOp.cellKey,
        local: oursOp,
        remote: theirsOp
      });
    }
  }
 
  /**
   * @param {{ move: any, edit: any }} input
   */
  _maybeAutoMergeMoveEdit(input) {
    const { move, edit } = input;
    if (!move?.fromCellKey || !move?.toCellKey) return;
    if (!edit?.cellKey) return;
    if (edit.cellKey !== move.fromCellKey) return;
 
    // Only relocate edits when the edit transaction didn't also touch the
    // destination cell (same semantics as BranchService rename-aware merges).
    const touched = Array.isArray(edit.touchedCells) ? edit.touchedCells : [];
    if (touched.includes(move.toCellKey)) {
      // Not safe to auto-merge when destination is also modified. Surface as a
      // delete-vs-edit conflict on the source cell (the move implies deletion
      // at the source).
      const moveDelete = {
        kind: "delete",
        userId: move.userId,
        beforeState: move.beforeState,
        afterState: move.afterState,
        touchedCells: move.touchedCells,
        cellKey: move.fromCellKey,
        before: move.cell ?? null,
        after: null
      };
 
      const oursIsEdit = edit.userId === this.localUserId;
 
      this._emitConflict({
        type: "cell",
        reason: "delete-vs-edit",
        sourceCellKey: move.fromCellKey,
        local: oursIsEdit ? edit : moveDelete,
        remote: oursIsEdit ? moveDelete : edit
      });
      return;
    }
 
    // Optional safety check: ensure the edit was applied to the same cell
    // content that was moved.
    if (move.fingerprint && edit.beforeFingerprint && move.fingerprint !== edit.beforeFingerprint) {
      return;
    }
 
    const desired = normalizeCell(edit.after);
 
    // If document already reflects the auto-merged state, do nothing.
    const currentSource = normalizeCell(this.cells.get(move.fromCellKey));
    const currentDest = normalizeCell(this.cells.get(move.toCellKey));
    if (cellsEqual(currentSource, null) && cellsEqual(currentDest, desired)) return;
 
    this._isApplyingResolution = true;
    try {
      this.doc.transact(() => {
        this.cells.delete(move.fromCellKey);
        if (desired === null) {
          this.cells.delete(move.toCellKey);
        } else {
          this._writeCell(move.toCellKey, desired);
        }
      }, this.origin);
    } finally {
      this._isApplyingResolution = false;
    }
  }
 
  /**
   * @param {{ type: "move" | "cell", reason: any, sourceCellKey: string, local: any, remote: any }} input
   */
  _emitConflict(input) {
    const ref = cellRefFromKey(input.sourceCellKey);
    const addr = `${numberToCol(ref.col)}${ref.row + 1}`;
 
    const conflict = /** @type {CellStructuralConflict} */ ({
      id: crypto.randomUUID(),
      type: input.type,
      reason: input.reason,
      sheetId: ref.sheetId,
      cell: addr,
      cellKey: input.sourceCellKey,
      local: simplifyOpForConflict(input.local),
      remote: simplifyOpForConflict(input.remote),
      remoteUserId: String(input.remote?.userId ?? ""),
      detectedAt: Date.now()
    });
 
    this._conflicts.set(conflict.id, conflict);
    this.onConflict(conflict);
  }
 
  /**
   * @param {string} cellKey
   * @param {NormalizedCell} cell
   */
  _writeCell(cellKey, cell) {
    let cellMap = /** @type {Y.Map<any> | undefined} */ (this.cells.get(cellKey));
    if (!isYMap(cellMap)) {
      cellMap = new Y.Map();
      this.cells.set(cellKey, cellMap);
    }
 
    const normalized = normalizeCell(cell);
    if (normalized === null) {
      this.cells.delete(cellKey);
      return;
    }
 
    if (normalized.formula != null) {
      cellMap.set("formula", normalized.formula);
      cellMap.set("value", null);
    } else {
      cellMap.delete("formula");
      cellMap.set("value", normalized.value ?? null);
    }
 
    cellMap.delete("style");
    if (normalized.format != null) {
      cellMap.set("format", normalized.format);
    } else {
      cellMap.delete("format");
    }
 
    cellMap.set("modified", Date.now());
    cellMap.set("modifiedBy", this.localUserId);
  }
 }
 
 /**
  * @param {string} a
  * @param {string} b
  */
 function makePairKey(a, b) {
   return a < b ? `${a}|${b}` : `${b}|${a}`;
 }
 
 /**
  * @param {any} op
  */
 function simplifyOpForConflict(op) {
   if (!op || typeof op !== "object") return null;
   if (op.kind === "move") {
     return {
       kind: "move",
       fromCellKey: op.fromCellKey,
       toCellKey: op.toCellKey,
       cell: normalizeCell(op.cell)
     };
   }
 
   if (op.kind === "delete" || op.kind === "edit") {
     return {
       kind: op.kind,
       cellKey: op.cellKey,
       before: normalizeCell(op.before),
       after: normalizeCell(op.after)
     };
   }
 
   return { ...op };
 }
 
 /**
  * @param {Map<number, number>} sv
  * @returns {Array<[number, number]>}
  */
 function encodeStateVector(sv) {
   const out = [];
   for (const [client, clock] of sv.entries()) {
     out.push([client, clock]);
   }
   out.sort((a, b) => a[0] - b[0]);
   return out;
 }
 
 /**
  * @param {Array<[number, number]> | Record<string, number> | null | undefined} encoded
  * @returns {Map<number, number>}
  */
 function decodeStateVector(encoded) {
   const out = new Map();
   if (Array.isArray(encoded)) {
     for (const entry of encoded) {
       if (!Array.isArray(entry) || entry.length < 2) continue;
       out.set(Number(entry[0]), Number(entry[1]));
     }
     return out;
   }
   if (encoded && typeof encoded === "object") {
     for (const [k, v] of Object.entries(encoded)) {
       out.set(Number(k), Number(v));
     }
   }
   return out;
 }
 
 /**
  * Returns true if state vector `a` dominates `b` (a >= b for all clients).
  * @param {Map<number, number>} a
  * @param {Map<number, number>} b
  */
 function dominatesStateVector(a, b) {
   for (const [client, clockB] of b.entries()) {
     const clockA = a.get(client) ?? 0;
     if (clockA < clockB) return false;
   }
   return true;
 }
 
 /**
  * @param {any} opA
  * @param {any} opB
  */
 function isCausallyConcurrent(opA, opB) {
   const aBefore = decodeStateVector(opA.beforeState);
   const aAfter = decodeStateVector(opA.afterState);
   const bBefore = decodeStateVector(opB.beforeState);
   const bAfter = decodeStateVector(opB.afterState);
 
   const aBeforeB = dominatesStateVector(bBefore, aAfter);
   const bBeforeA = dominatesStateVector(aBefore, bAfter);
   return !aBeforeB && !bBeforeA;
 }
 
 /**
  * @param {Array<any>} events
  * @param {Y.Map<any>} cells
  * @returns {{ moves: Array<{ fromCellKey: string, toCellKey: string, cell: NormalizedCell, fingerprint: string }>, deletes: Array<any>, edits: Array<any>, touchedCells: string[] } | null}
  */
 function extractTransactionChanges(events, cells) {
   /** @type {Map<string, { mapChange?: any, propChanges: Map<string, any> }>} */
   const touched = new Map();
 
   for (const event of events) {
     if (!event?.changes?.keys) continue;
 
     // Map-level changes on the "cells" map itself.
     if (event.target === cells) {
       for (const [key, change] of event.changes.keys.entries()) {
         if (typeof key !== "string") continue;
         const entry = touched.get(key) ?? { propChanges: new Map() };
         entry.mapChange = change;
         touched.set(key, entry);
       }
       continue;
     }
 
     const path = event.path ?? [];
     const cellKey = path[0];
     if (typeof cellKey !== "string") continue;
 
     const entry = touched.get(cellKey) ?? { propChanges: new Map() };
     touched.set(cellKey, entry);
 
     for (const [prop, change] of event.changes.keys.entries()) {
       if (prop !== "value" && prop !== "formula" && prop !== "format" && prop !== "style") continue;
       entry.propChanges.set(prop, change);
     }
   }
 
   if (touched.size === 0) return null;
 
   /** @type {Array<{ cellKey: string, before: NormalizedCell | null, after: NormalizedCell | null, beforeFingerprint: string | null, afterFingerprint: string | null }>} */
   const diffs = [];
 
   for (const [cellKey, entry] of touched.entries()) {
     const after = normalizeCell(cells.get(cellKey));
 
     /** @type {NormalizedCell | null} */
     let before = null;
 
     if (entry.mapChange?.action === "add") {
       before = null;
     } else if (entry.mapChange?.action === "delete" || entry.mapChange?.action === "update") {
       before = normalizeCell(entry.mapChange.oldValue);
     } else if (entry.propChanges.size > 0) {
       const seed = after
         ? { value: after.value ?? null, formula: after.formula ?? null, format: after.format ?? null }
         : { value: null, formula: null, format: null };
 
       for (const [prop, change] of entry.propChanges.entries()) {
         if (prop === "value") seed.value = change.oldValue ?? null;
         if (prop === "formula") seed.formula = change.oldValue ?? null;
         if (prop === "format" || prop === "style") seed.format = change.oldValue ?? null;
       }
 
       before = normalizeCell(seed);
     } else {
       // Only non-structural metadata changed (e.g. modified timestamp).
       continue;
     }
 
     const beforeFingerprint = cellFingerprint(before);
     const afterFingerprint = cellFingerprint(after);
 
     // Ignore pure metadata updates that don't affect the structural footprint.
     if (beforeFingerprint === afterFingerprint) continue;
 
     diffs.push({ cellKey, before, after, beforeFingerprint, afterFingerprint });
   }
 
   if (diffs.length === 0) return null;
 
   /** @type {Map<string, string[]>} */
   const additionsByFingerprint = new Map();
   for (const diff of diffs) {
     if (diff.before === null && diff.after !== null && diff.afterFingerprint) {
       const list = additionsByFingerprint.get(diff.afterFingerprint) ?? [];
       list.push(diff.cellKey);
       additionsByFingerprint.set(diff.afterFingerprint, list);
     }
   }
 
   /** @type {Set<string>} */
   const consumedAddrs = new Set();
   /** @type {Set<string>} */
   const consumedFrom = new Set();
 
   /** @type {Array<{ fromCellKey: string, toCellKey: string, cell: NormalizedCell, fingerprint: string }>} */
   const moves = [];
 
   for (const diff of diffs) {
     if (diff.before === null || diff.after !== null) continue;
     if (!diff.beforeFingerprint) continue;
     const candidates = additionsByFingerprint.get(diff.beforeFingerprint) ?? [];
     const target = candidates.find((k) => !consumedAddrs.has(k));
     if (!target) continue;
     moves.push({
       fromCellKey: diff.cellKey,
       toCellKey: target,
       cell: diff.before,
       fingerprint: diff.beforeFingerprint
     });
     consumedAddrs.add(target);
     consumedFrom.add(diff.cellKey);
   }
 
   /** @type {Array<any>} */
   const deletes = [];
   /** @type {Array<any>} */
   const edits = [];
 
   for (const diff of diffs) {
     if (consumedFrom.has(diff.cellKey) || consumedAddrs.has(diff.cellKey)) continue;
 
     if (diff.before !== null && diff.after === null) {
       deletes.push({ cellKey: diff.cellKey, before: diff.before, fingerprint: diff.beforeFingerprint });
       continue;
     }
 
     if (diff.before !== null && diff.after !== null) {
       const contentChanged = didContentChange(diff.before, diff.after);
       const formatChanged = didFormatChange(diff.before, diff.after);
       edits.push({
         cellKey: diff.cellKey,
         before: diff.before,
         after: diff.after,
         beforeFingerprint: diff.beforeFingerprint,
         afterFingerprint: diff.afterFingerprint,
         contentChanged,
         formatChanged
       });
     }
   }
 
   return {
     moves,
     deletes,
     edits,
     touchedCells: Array.from(touched.keys())
   };
 }
 
 /**
  * @param {any} cellData
  * @returns {NormalizedCell | null}
  */
 function normalizeCell(cellData) {
   if (cellData == null) return null;
 
   /** @type {any} */
   let value = null;
   /** @type {string | null} */
   let formula = null;
   /** @type {Record<string, unknown> | null} */
   let format = null;
 
   if (isYMap(cellData)) {
     value = cellData.get("value") ?? null;
     formula = cellData.get("formula") ?? null;
     format = cellData.get("format") ?? cellData.get("style") ?? null;
   } else if (typeof cellData === "object") {
     value = cellData.value ?? null;
     formula = cellData.formula ?? null;
     format = cellData.format ?? cellData.style ?? null;
   } else {
     value = cellData;
   }
 
   if (value && typeof value === "object" && value.t === "blank") {
     value = null;
   }
 
   const normalizedFormula = typeof formula === "string" ? formula.trim() : "";
   formula = normalizedFormula ? normalizedFormula : null;
 
   if (format && typeof format !== "object") {
     // Ensure format is JSON-friendly.
     format = null;
   }
 
   const hasValue = value !== null && value !== undefined && value !== "";
   const hasFormula = formula != null && formula !== "";
   const hasFormat = format != null;
 
   if (!hasValue && !hasFormula && !hasFormat) return null;
 
   /** @type {NormalizedCell} */
   const out = {};
   if (hasFormula) out.formula = formula;
   else if (hasValue) out.value = value;
   if (hasFormat) out.format = format;
   return out;
 }

 /**
  * Duck-type Y.Map detection to avoid `instanceof` pitfalls when multiple Yjs
  * module instances are present (e.g. pnpm workspaces + mixed ESM/CJS loaders).
  *
  * @param {any} value
  * @returns {value is Y.Map<any>}
  */
 function isYMap(value) {
   if (value instanceof Y.Map) return true;
   if (!value || typeof value !== "object") return false;
   const maybe = value;
   if (maybe.constructor?.name !== "YMap") return false;
   if (typeof maybe.get !== "function") return false;
   if (typeof maybe.set !== "function") return false;
   if (typeof maybe.delete !== "function") return false;
   return true;
 }
 
 /**
  * @param {NormalizedCell | null} cell
  * @returns {string | null}
  */
 function cellFingerprint(cell) {
   const normalized = normalizeCell(cell);
   if (normalized === null) return null;
 
   return stableStringify({
     value: normalized.value ?? null,
     formula: normalizeFormula(normalized.formula) ?? null,
     format: normalized.format ?? null
   });
 }
 
 /**
  * @param {string | null | undefined} formula
  */
 function normalizeFormula(formula) {
   const stripped = String(formula ?? "").trim().replace(/^\s*=\s*/, "");
   if (!stripped) return null;
   return stripped.replaceAll(/\s+/g, "").toUpperCase();
 }
 
 /**
  * Stable stringify for objects so we can build deterministic signatures.
  * @param {any} value
  * @returns {string}
  */
 function stableStringify(value) {
   if (value === null) return "null";
   const t = typeof value;
   if (t === "string") return JSON.stringify(value);
   if (t === "number" || t === "boolean") return JSON.stringify(value);
   if (t === "undefined") return "undefined";
   if (t === "bigint") return JSON.stringify(String(value));
   if (t === "function") return "\"[Function]\"";
   if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
   if (t === "object") {
     const keys = Object.keys(value).sort();
     return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
   }
   return JSON.stringify(value);
 }
 
 /**
  * @param {NormalizedCell | null} a
  * @param {NormalizedCell | null} b
  */
 function cellsEqual(a, b) {
   return stableStringify(a ?? null) === stableStringify(b ?? null);
 }
 
 /**
  * @param {NormalizedCell | null} a
  * @param {NormalizedCell | null} b
  */
 function didContentChange(a, b) {
   const na = normalizeCell(a);
   const nb = normalizeCell(b);
   return (
     stableStringify({ value: na?.value ?? null, formula: normalizeFormula(na?.formula) ?? null }) !==
     stableStringify({ value: nb?.value ?? null, formula: normalizeFormula(nb?.formula) ?? null })
   );
 }
 
 /**
  * @param {NormalizedCell | null} a
  * @param {NormalizedCell | null} b
  */
 function didFormatChange(a, b) {
   const na = normalizeCell(a);
   const nb = normalizeCell(b);
   return stableStringify(na?.format ?? null) !== stableStringify(nb?.format ?? null);
 }
 
