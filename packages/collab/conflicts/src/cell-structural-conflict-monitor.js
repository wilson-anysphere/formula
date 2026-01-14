import * as Y from "yjs";
import { getMapRoot, getYMap, yjsValueToJson } from "@formula/collab-yjs-utils";
import { cellRefFromKey, numberToCol } from "./cell-ref.js";
 
/**
 * @typedef {object} NormalizedCell
 * @property {unknown} [value]
 * @property {string} [formula]
 * @property {unknown} [enc]
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
 * @property {string} remoteUserId Best-effort id of the remote user. May be an empty string when unavailable.
 * @property {number} detectedAt
 */
 
/**
 * @typedef {object} CellStructuralConflictResolution
 * @property {"ours" | "theirs" | "manual"} choice
 * @property {string} [to] Destination cellKey for resolving move conflicts.
 * @property {NormalizedCell | null} [cell] Optional manual cell contents.
 *   - For `type: "cell"` conflicts, this is the final cell value.
 *   - For `type: "move"` conflicts, this overrides the moved cell content when
 *     `choice: "manual"` is used (in addition to selecting `to`).
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
   * @param {Set<any>} [opts.ignoredOrigins] Transaction origins to ignore entirely.
   * @param {(conflict: CellStructuralConflict) => void} opts.onConflict
   * @param {number} [opts.maxOpRecordsPerUser] Maximum number of structural op
   *   records to retain per user in the shared `cellStructuralOps` log.
   * @param {number | null} [opts.maxOpRecordAgeMs] Optional age-based pruning
   *   window for records in the shared `cellStructuralOps` log. When enabled,
   *   records older than `Date.now() - maxOpRecordAgeMs` may be deleted by any
   *   client (best-effort).
   *
   *   Pruning is additionally conservative relative to the local op log queue:
   *   records are only pruned when they are older than both the age cutoff and
   *   the oldest local op record (queue head). This avoids deleting records that
   *   may still be needed to compare against local ops that are in-flight (e.g.
   *   long offline periods).
   *
   *   Pruning is conservative: records are not deleted in the same op-log
   *   transaction they are added, so late-arriving/offline records have a chance
   *   to be ingested by other clients before being removed.
   *
   *   Pruning is incremental: very large logs may take multiple passes to fully
   *   clean up.
   *   Defaults to null (disabled).
   */
  constructor(opts) {
    this._maxOpRecordsPerUser = opts.maxOpRecordsPerUser ?? 2000;
    this._maxOpRecordAgeMs = opts.maxOpRecordAgeMs ?? null;
    this._lastAgePruneAt = 0;
    /** @type {ReturnType<typeof setTimeout> | null} */
    this._agePruneTimer = null;
    this._disposed = false;
    this.doc = opts.doc;
    this.cells = opts.cells ?? getMapRoot(this.doc, "cells");
    this.localUserId = opts.localUserId;
  
    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);
    this.ignoredOrigins = opts.ignoredOrigins ?? new Set();
  
    this.onConflict = opts.onConflict;
 
    // Shared operation log used to exchange per-transaction causal metadata
    // (transaction.beforeState/afterState) between clients.
    this._ops = getMapRoot(this.doc, "cellStructuralOps");
 
    /** @type {Map<string, CellStructuralConflict>} */
    this._conflicts = new Map();
 
    /** @type {Map<string, any>} */
    this._opRecords = new Map();

    /** @type {Array<{ id: string, createdAt: number }>} */
    this._localOpQueue = [];
    /** @type {Set<string>} */
    this._localOpIds = new Set();
    this._localOpQueueInitialized = false;

    this._isApplyingResolution = false;
 
    this._onCellsDeepEvent = this._onCellsDeepEvent.bind(this);
    this.cells.observeDeep(this._onCellsDeepEvent);
 
    this._onOpsEvent = this._onOpsEvent.bind(this);
    this._ops.observe(this._onOpsEvent);

    // If configured, opportunistically prune old operation records from existing
    // docs before ingesting them into local state. This keeps startup costs
    // bounded when opening long-lived documents.
    this._pruneOpLogByAge({ force: true });

    // Seed local state from any persisted op log entries so conflict detection
    // still works after a client restarts while offline (ops already exist in
    // the doc, so we won't observe "add" events for them).
    this._ensureLocalOpQueueInitialized();
  }
 
  dispose() {
    this._disposed = true;
    this.cells.unobserveDeep(this._onCellsDeepEvent);
    this._ops.unobserve(this._onOpsEvent);
    if (this._agePruneTimer) {
      clearTimeout(this._agePruneTimer);
      this._agePruneTimer = null;
    }
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
    let ok = true;
    try {
      if (conflict.type === "move") {
        const oursTo = conflict.local?.toCellKey ?? null;
        const theirsTo = conflict.remote?.toCellKey ?? null;
        const oursCell = conflict.local?.cell ?? null;
        const theirsCell = conflict.remote?.cell ?? null;
 
        let target = null;
        if (resolution.choice === "ours") target = oursTo;
        else if (resolution.choice === "theirs") target = theirsTo;
        else target = resolution.to ?? null;
 
        if (!target) return false;
 
        const chosenCell =
          resolution.choice === "manual" && "cell" in resolution
            ? resolution.cell
            : resolution.choice === "theirs"
              ? theirsCell ?? oursCell
              : oursCell ?? theirsCell;
 
        this.doc.transact(() => {
          // Clear source + both destinations.
          this._clearCell(conflict.cellKey);
          if (oursTo) this._clearCell(oursTo);
          if (theirsTo) this._clearCell(theirsTo);

          if (chosenCell) {
            this._writeCell(target, chosenCell);
          }
        }, this.origin);
      } else {
        let finalCell = null;
        if (resolution.choice === "ours") finalCell = conflict.local?.after ?? null;
        else if (resolution.choice === "theirs") finalCell = conflict.remote?.after ?? null;
        else finalCell = resolution.cell ?? null;
 
        this.doc.transact(() => {
          if (finalCell === null) {
            this._clearCell(conflict.cellKey);
          } else {
            this._writeCell(conflict.cellKey, finalCell);
          }
        }, this.origin);
      }
    } catch {
      ok = false;
    } finally {
      this._isApplyingResolution = false;
    }
 
    if (!ok) return false;
    this._conflicts.delete(conflictId);
    return true;
  }
 
  /**
   * @param {Array<any>} events
   * @param {Y.Transaction} transaction
   */
  _onCellsDeepEvent(events, transaction) {
    if (this._isApplyingResolution) return;

    if (this.ignoredOrigins?.has(transaction.origin)) return;
  
    const isLocal = this.localOrigins.has(transaction.origin);
    if (!isLocal) return;
 
    const tx = extractTransactionChanges(events, this.cells);
    if (!tx) return;

    const beforeState = encodeStateVector(transaction.beforeState);
    const txId = crypto.randomUUID();
    const createdAt = Date.now();
    /** @type {Array<any>} */
    const records = [];

    for (const move of tx.moves) {
      records.push({
        id: crypto.randomUUID(),
        txId,
        kind: "move",
        userId: this.localUserId,
        createdAt,
        beforeState,
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
        createdAt,
        beforeState,
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
        createdAt,
        beforeState,
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

    // Track local op ids eagerly so pruning can avoid scanning the entire shared
    // log on every write.
    this._ensureLocalOpQueueInitialized();
    const lastCreatedAt = this._localOpQueue.length ? this._localOpQueue[this._localOpQueue.length - 1].createdAt : -Infinity;
    let outOfOrder = false;
    for (const record of records) {
      if (this._localOpIds.has(record.id)) continue;
      this._localOpIds.add(record.id);
      this._localOpQueue.push({ id: record.id, createdAt: Number(record.createdAt ?? 0) });
      if (record.createdAt < lastCreatedAt) outOfOrder = true;
    }
    if (outOfOrder) {
      this._localOpQueue.sort((a, b) => a.createdAt - b.createdAt);
    }

    // `transaction.afterState` only advances when the transaction inserts new
    // structs. Pure deletions (DeleteSet-only transactions) leave the state
    // vector unchanged, which would make delete-vs-edit conflicts look causally
    // ordered when they were actually concurrent.
    //
    // In that case, we know we will insert at least one new struct when writing
    // our op log entries (one per record), so we can treat the op log insertion
    // as the "clock tick" that represents this mutation.
    //
    // For transactions that *do* insert structs, prefer the original
    // `transaction.afterState` so other clients only need to have seen the cell
    // change itself to establish causal ordering (they shouldn't also need to
    // have received the op log update).
    const afterStateVector = new Map(transaction.afterState);
    if (stateVectorsEqual(transaction.afterState, transaction.beforeState)) {
      const clientId = this.doc.clientID;
      afterStateVector.set(clientId, (afterStateVector.get(clientId) ?? 0) + records.length);
    }
    const afterState = encodeStateVector(afterStateVector);
    for (const record of records) {
      record.afterState = afterState;
    }

    // Store operation records in a shared log so other clients can compare
    // state vectors and detect true concurrency.
    this.doc.transact(() => {
      for (const record of records) {
        this._ops.set(record.id, record);
      }
    }, this.origin);

    this._pruneLocalOpLog();
  }
 
  /**
   * @param {Y.YMapEvent<any>} event
   * @param {Y.Transaction} transaction
   */
  _onOpsEvent(event, transaction) {
    if (this._isApplyingResolution) return;
    if (this.ignoredOrigins?.has(transaction.origin)) return;
    if (!event?.changes?.keys) return;
  
    /** @type {string[]} */
    const localDeletes = [];
    let sawAdd = false;
    /** @type {Set<string>} */
    const addedIds = new Set();
    for (const [opId, change] of event.changes.keys.entries()) {
      const id = String(opId);
      if (change.action === "delete") {
        this._opRecords.delete(id);
        if (this._localOpIds.has(id)) {
          this._localOpIds.delete(id);
          localDeletes.push(id);
        }
        continue;
      }
  
      if (change.action !== "add") continue;
      sawAdd = true;
      addedIds.add(id);
      const record = this._ops.get(id);
      if (!record) continue;
      this._ingestOpRecord(record, id);
    }

    if (localDeletes.length > 0 && this._localOpQueue.length > 0) {
      const toRemove = new Set(localDeletes);
      this._localOpQueue = this._localOpQueue.filter((entry) => !toRemove.has(entry.id));
    }

    if (sawAdd) {
      this._maybeScheduleDeferredAgePrune(addedIds);
      // Conservative safety: avoid pruning op records in the same transaction
      // they were added. This prevents us from deleting late-arriving (offline)
      // records immediately, which could cause other clients to miss the entry
      // before they've had a chance to ingest it and compare for conflicts.
      this._pruneOpLogByAge({ excludeIds: addedIds });
    }
  }
 
  /**
   * @param {any} record
   * @param {string} [opId]
   */
  _ingestOpRecord(record, opId) {
    if (!record || typeof record !== "object") return;
    const id = String(opId ?? record.id ?? "");
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
 
      const oursOp = ours ? record : other;
      const theirsOp = ours ? other : record;
 
      this._handleOpPair(oursOp, theirsOp);
    }
  }

  _pruneLocalOpLog() {
    const limit = Number(this._maxOpRecordsPerUser);
    if (!Number.isFinite(limit) || limit <= 0) return;

    this._ensureLocalOpQueueInitialized();
    if (this._localOpQueue.length <= limit) return;

    const toDelete = this._localOpQueue.slice(0, this._localOpQueue.length - limit);
    if (toDelete.length === 0) return;

    this._localOpQueue = this._localOpQueue.slice(this._localOpQueue.length - limit);
    for (const entry of toDelete) {
      this._localOpIds.delete(entry.id);
    }

    this.doc.transact(() => {
      for (const entry of toDelete) {
        this._ops.delete(entry.id);
      }
    }, this.origin);
  }

  /**
   * Best-effort pruning of the shared `cellStructuralOps` log by wall-clock age.
   *
   * This is intentionally conservative: it only considers records with a finite
   * numeric `createdAt` timestamp and is safe under concurrent clients (the log
   * is metadata-only).
   *
   * @param {{ force?: boolean, excludeIds?: Set<string> }} [opts]
   */
  _pruneOpLogByAge(opts = {}) {
    const maxAgeMs = this._maxOpRecordAgeMs;
    if (maxAgeMs == null) return;
    const ageMs = Number(maxAgeMs);
    if (!Number.isFinite(ageMs) || ageMs <= 0) return;

    const now = Date.now();

    // Throttle so we don't scan the shared log on every single local edit.
    const minIntervalMs = Math.max(5_000, Math.min(60_000, Math.floor(ageMs / 10)));
    if (!opts.force && now - this._lastAgePruneAt < minIntervalMs) return;
    this._lastAgePruneAt = now;

    const ageCutoff = now - ageMs;
    const localQueueHeadCreatedAt = this._getLocalOpQueueHeadCreatedAt();
    // Conservative policy: only prune entries that are older than the age cutoff
    // *and* older than our oldest local op record. This avoids deleting records
    // that might still be needed to compare against local operations that have
    // not yet been fully observed by other clients (best-effort).
    const cutoff = Math.min(ageCutoff, localQueueHeadCreatedAt);
  
    /** @type {string[]} */
    const toDelete = [];
    let hitDeleteLimit = false;
    // Upper bound on deletes per prune pass. Limits the size of the resulting
    // Yjs update and avoids long blocking scans/transactions when many expired
    // entries accumulate.
    const maxDeletesPerPass = 1_000;

    // Iterate manually (instead of `forEach`) so we can early-break once we have
    // enough deletions to make progress. This avoids scanning the entire map on
    // every prune pass when large logs accumulate.
    for (const rawId of this._ops.keys()) {
      const id = String(rawId);
      if (opts.excludeIds?.has(id)) continue;
      const record = this._ops.get(rawId);
      if (!record || typeof record !== "object") continue;
      const createdAt = Number(record.createdAt);
      if (!Number.isFinite(createdAt)) continue;
      if (createdAt < cutoff) {
        toDelete.push(id);
        if (toDelete.length >= maxDeletesPerPass) {
          hitDeleteLimit = true;
          break;
        }
      }
    }

    if (toDelete.length === 0) return;

    this.doc.transact(() => {
      for (const id of toDelete) {
        this._ops.delete(id);
      }
    }, this.origin);

    // If we hit our per-pass delete cap, schedule another prune pass so large
    // backlogs are eventually cleared without requiring new writes.
    if (hitDeleteLimit) {
      this._scheduleAgePruneTimer(minIntervalMs);
    }
  }

  _getLocalOpQueueHeadCreatedAt() {
    // Fast path: local queue already initialized and sorted.
    if (this._localOpQueueInitialized) {
      if (this._localOpQueue.length === 0) return Infinity;
      for (const entry of this._localOpQueue) {
        const t = Number(entry?.createdAt);
        if (Number.isFinite(t) && t > 0) return t;
      }
      // Local queue is initialized but contains no valid timestamps.
      return Infinity;
    }

    // Slow path (startup): scan existing op log for the oldest local record.
    let min = Infinity;
    for (const rawId of this._ops.keys()) {
      const record = this._ops.get(rawId);
      if (!record || typeof record !== "object") continue;
      if (record.userId !== this.localUserId) continue;
      const createdAt = Number(record.createdAt);
      if (!Number.isFinite(createdAt) || createdAt <= 0) continue;
      if (createdAt < min) min = createdAt;
    }
    return min;
  }

  /**
   * If the current op-log transaction added records that are already past the
   * age cutoff, schedule a deferred prune so the records don't stick around
   * indefinitely when the document becomes idle (and to give other clients a
   * chance to ingest the entries before deletion).
   *
   * @param {Set<string>} addedIds
   */
  _maybeScheduleDeferredAgePrune(addedIds) {
    const maxAgeMs = this._maxOpRecordAgeMs;
    if (maxAgeMs == null) return;
    const ageMs = Number(maxAgeMs);
    if (!Number.isFinite(ageMs) || ageMs <= 0) return;
    if (!addedIds || addedIds.size === 0) return;

    const now = Date.now();
    const ageCutoff = now - ageMs;
    const localQueueHeadCreatedAt = this._getLocalOpQueueHeadCreatedAt();
    const cutoff = Math.min(ageCutoff, localQueueHeadCreatedAt);

    let hasExpiredAdded = false;
    for (const id of addedIds) {
      const record = this._ops.get(id);
      if (!record || typeof record !== "object") continue;
      const createdAt = Number(record.createdAt);
      if (!Number.isFinite(createdAt)) continue;
      if (createdAt < cutoff) {
        hasExpiredAdded = true;
        break;
      }
    }
    if (!hasExpiredAdded) return;

    const minIntervalMs = Math.max(5_000, Math.min(60_000, Math.floor(ageMs / 10)));
    this._scheduleAgePruneTimer(minIntervalMs);
  }

  /**
   * @param {number} delayMs
   */
  _scheduleAgePruneTimer(delayMs) {
    // If a timer is already scheduled, keep it; a single deferred prune is
    // sufficient to eventually clean up any late-arriving expired records.
    if (this._disposed) return;
    if (this._agePruneTimer) return;

    this._agePruneTimer = setTimeout(() => {
      if (this._disposed) {
        this._agePruneTimer = null;
        return;
      }
      this._agePruneTimer = null;
      // Force to ensure we actually run even if the timer fires slightly early
      // (and therefore the normal throttle window hasn't fully elapsed yet).
      this._pruneOpLogByAge({ force: true });
    }, delayMs);

    // In Node tests, avoid keeping the process alive solely for best-effort
    // pruning. Browsers don't support `unref`.
    if (typeof this._agePruneTimer?.unref === "function") {
      this._agePruneTimer.unref();
    }
  }

  _ensureLocalOpQueueInitialized() {
    if (this._localOpQueueInitialized) return;
    this._localOpQueueInitialized = true;

    /** @type {Array<{ id: string, createdAt: number }>} */
    const ours = [];
    this._ops.forEach((record, id) => {
      if (!record || typeof record !== "object") return;
      this._ingestOpRecord(record, String(id));
      if (record.userId !== this.localUserId) return;
      ours.push({ id: String(id), createdAt: Number(record.createdAt ?? 0) });
    });
    ours.sort((a, b) => a.createdAt - b.createdAt);

    this._localOpQueue = ours;
    this._localOpIds = new Set(ours.map((entry) => entry.id));
  }
  
  /**
   * @param {any} oursOp
   * @param {any} theirsOp
   */
  _handleOpPair(oursOp, theirsOp) {
    if (!isCausallyConcurrent(oursOp, theirsOp)) return;
 
    // move-related conflicts.
    if (oursOp.kind === "move" && theirsOp.kind === "move") {
      if (oursOp.fromCellKey === theirsOp.fromCellKey) {
        if (oursOp.toCellKey !== theirsOp.toCellKey) {
          this._emitConflict({
            type: "move",
            reason: "move-destination",
            sourceCellKey: oursOp.fromCellKey,
            local: oursOp,
            remote: theirsOp
          });
        } else {
          const oursCell = normalizeCell(oursOp.cell);
          const theirsCell = normalizeCell(theirsOp.cell);
          if (cellsEqual(oursCell, theirsCell)) return;

          const reason = didContentChange(oursCell, theirsCell) ? "content" : "format";

          const oursAsEdit = {
            kind: "edit",
            userId: oursOp.userId,
            cellKey: oursOp.toCellKey,
            before: null,
            after: oursCell
          };
          const theirsAsEdit = {
            kind: "edit",
            userId: theirsOp.userId,
            cellKey: theirsOp.toCellKey,
            before: null,
            after: theirsCell
          };

          this._emitConflict({
            type: "cell",
            reason,
            sourceCellKey: oursOp.toCellKey,
            local: oursAsEdit,
            remote: theirsAsEdit
          });
        }
        return;
      }

      // Two different sources moved into the same destination.
      if (oursOp.toCellKey === theirsOp.toCellKey) {
        const oursCell = normalizeCell(oursOp.cell);
        const theirsCell = normalizeCell(theirsOp.cell);
        if (cellsEqual(oursCell, theirsCell)) return;

        const reason = didContentChange(oursCell, theirsCell) ? "content" : "format";

        const oursAsEdit = {
          kind: "edit",
          userId: oursOp.userId,
          cellKey: oursOp.toCellKey,
          before: null,
          after: oursCell
        };
        const theirsAsEdit = {
          kind: "edit",
          userId: theirsOp.userId,
          cellKey: theirsOp.toCellKey,
          before: null,
          after: theirsCell
        };

        this._emitConflict({
          type: "cell",
          reason,
          sourceCellKey: oursOp.toCellKey,
          local: oursAsEdit,
          remote: theirsAsEdit
        });
      }
      return;
    }

    // Move vs delete: treat as a delete-vs-edit conflict at the destination.
    if (oursOp.kind === "move" && theirsOp.kind === "delete") {
      if (theirsOp.cellKey === oursOp.toCellKey) {
        const oursAsEdit = {
          kind: "edit",
          userId: oursOp.userId,
          cellKey: oursOp.toCellKey,
          before: null,
          after: normalizeCell(oursOp.cell)
        };
        this._emitConflict({
          type: "cell",
          reason: "delete-vs-edit",
          sourceCellKey: oursOp.toCellKey,
          local: oursAsEdit,
          remote: theirsOp
        });
      } else if (
        theirsOp.cellKey === oursOp.fromCellKey &&
        (!oursOp.fingerprint || !theirsOp.fingerprint || oursOp.fingerprint === theirsOp.fingerprint)
      ) {
        const oursAsEdit = {
          kind: "edit",
          userId: oursOp.userId,
          cellKey: oursOp.toCellKey,
          before: null,
          after: normalizeCell(oursOp.cell)
        };
        const theirsAsDelete = {
          kind: "delete",
          userId: theirsOp.userId,
          cellKey: oursOp.toCellKey,
          before: normalizeCell(oursOp.cell),
          after: null
        };
        this._emitConflict({
          type: "cell",
          reason: "delete-vs-edit",
          sourceCellKey: oursOp.toCellKey,
          local: oursAsEdit,
          remote: theirsAsDelete
        });
      }
      return;
    }
    if (oursOp.kind === "delete" && theirsOp.kind === "move") {
      if (oursOp.cellKey === theirsOp.toCellKey) {
        const theirsAsEdit = {
          kind: "edit",
          userId: theirsOp.userId,
          cellKey: theirsOp.toCellKey,
          before: null,
          after: normalizeCell(theirsOp.cell)
        };
        this._emitConflict({
          type: "cell",
          reason: "delete-vs-edit",
          sourceCellKey: theirsOp.toCellKey,
          local: oursOp,
          remote: theirsAsEdit
        });
      } else if (
        oursOp.cellKey === theirsOp.fromCellKey &&
        (!oursOp.fingerprint || !theirsOp.fingerprint || oursOp.fingerprint === theirsOp.fingerprint)
      ) {
        const oursAsDelete = {
          kind: "delete",
          userId: oursOp.userId,
          cellKey: theirsOp.toCellKey,
          before: normalizeCell(theirsOp.cell),
          after: null
        };
        const theirsAsEdit = {
          kind: "edit",
          userId: theirsOp.userId,
          cellKey: theirsOp.toCellKey,
          before: null,
          after: normalizeCell(theirsOp.cell)
        };
        this._emitConflict({
          type: "cell",
          reason: "delete-vs-edit",
          sourceCellKey: theirsOp.toCellKey,
          local: oursAsDelete,
          remote: theirsAsEdit
        });
      }
      return;
    }
 
    // move vs edit: destination collisions surface conflicts; source edits are
    // eligible for rename-aware auto-merge.
    if (oursOp.kind === "move" && theirsOp.kind === "edit") {
      if (theirsOp.cellKey === oursOp.toCellKey) {
        const oursCell = normalizeCell(oursOp.cell);
        const theirsCell = normalizeCell(theirsOp.after);
        if (!cellsEqual(oursCell, theirsCell)) {
          const reason = didContentChange(oursCell, theirsCell) ? "content" : "format";
          const oursAsEdit = {
            kind: "edit",
            userId: oursOp.userId,
            cellKey: oursOp.toCellKey,
            before: null,
            after: oursCell
          };
          this._emitConflict({
            type: "cell",
            reason,
            sourceCellKey: oursOp.toCellKey,
            local: oursAsEdit,
            remote: theirsOp
          });
        }
        return;
      }
      this._maybeAutoMergeMoveEdit({ move: oursOp, edit: theirsOp });
      return;
    }
    if (oursOp.kind === "edit" && theirsOp.kind === "move") {
      if (oursOp.cellKey === theirsOp.toCellKey) {
        const oursCell = normalizeCell(oursOp.after);
        const theirsCell = normalizeCell(theirsOp.cell);
        if (!cellsEqual(oursCell, theirsCell)) {
          const reason = didContentChange(oursCell, theirsCell) ? "content" : "format";
          const theirsAsEdit = {
            kind: "edit",
            userId: theirsOp.userId,
            cellKey: theirsOp.toCellKey,
            before: null,
            after: theirsCell
          };
          this._emitConflict({
            type: "cell",
            reason,
            sourceCellKey: theirsOp.toCellKey,
            local: oursOp,
            remote: theirsAsEdit
          });
        }
        return;
      }
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
      // Not safe to auto-merge when destination is also modified. Let the
      // destination-level collision detection surface a conflict instead.
      return;
    }
 
    // Optional safety check: ensure the edit was applied to the same cell
    // content that was moved.
    if (move.fingerprint && move.fingerprint !== edit.beforeFingerprint) {
      return;
    }
 
    const desired = normalizeCell(edit.after);
 
    // If document already reflects the auto-merged state, do nothing.
    const currentSource = normalizeCell(this.cells.get(move.fromCellKey));
    const currentDest = normalizeCell(this.cells.get(move.toCellKey));
    if (cellsEqual(currentSource, null) && cellsEqual(currentDest, desired)) return;
 
    this._isApplyingResolution = true;
    try {
      try {
        this.doc.transact(() => {
          this._clearCell(move.fromCellKey);
          if (desired === null) {
            this._clearCell(move.toCellKey);
          } else {
            this._writeCell(move.toCellKey, desired);
          }
        }, this.origin);
      } catch {
        // If we can't apply the auto-merge safely (e.g. destination is encrypted
        // and we'd need to write plaintext), surface as a delete-vs-edit conflict
        // at the source cell.
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
      }
    } finally {
      this._isApplyingResolution = false;
    }
  }
 
  /**
   * @param {{ type: "move" | "cell", reason: any, sourceCellKey: string, local: any, remote: any }} input
   */
  _emitConflict(input) {
    let ref;
    try {
      ref = cellRefFromKey(input.sourceCellKey);
    } catch {
      // Ignore conflicts for malformed/unparseable cell keys. These can be introduced
      // by corrupted docs or malicious/buggy clients and should not crash observers.
      return;
    }
    const sheetId = String(ref?.sheetId ?? "");
    const row = Number(ref?.row);
    const col = Number(ref?.col);
    if (!sheetId) return;
    if (!Number.isInteger(row) || row < 0) return;
    if (!Number.isInteger(col) || col < 0) return;

    const addr = `${numberToCol(col)}${row + 1}`;
 
    const conflict = /** @type {CellStructuralConflict} */ ({
      id: crypto.randomUUID(),
      type: input.type,
      reason: input.reason,
      sheetId,
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
    const normalized = normalizeCell(cell);
    if (normalized === null) {
      this._clearCell(cellKey);
      return;
    }
 
    let cellMap = /** @type {Y.Map<any> | undefined} */ (this.cells.get(cellKey));
    if (!getYMap(cellMap)) {
      cellMap = new Y.Map();
      this.cells.set(cellKey, cellMap);
    }
 
    const existingEnc = cellMap.get("enc");
 
    // Never downgrade an encrypted cell to plaintext (CollabSession enforces the
    // same invariant). We allow format-only writes by preserving the existing
    // encrypted payload.
    if (existingEnc !== undefined && normalized.enc === undefined) {
      const wantsContent = normalized.formula != null || normalized.value != null;
      if (wantsContent) {
        throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
      }
      cellMap.set("enc", existingEnc);
      cellMap.delete("value");
      cellMap.delete("formula");
    } else if (normalized.enc !== undefined) {
      // Never downgrade a ciphertext payload to a null marker. If a caller passes
      // `{ enc: null }` while an existing non-null payload exists, preserve the
      // existing payload to avoid data loss.
      const nextEnc = normalized.enc === null && existingEnc !== undefined && existingEnc !== null ? existingEnc : normalized.enc;
      cellMap.set("enc", nextEnc);
      cellMap.delete("value");
      cellMap.delete("formula");
    } else if (normalized.formula != null) {
      cellMap.delete("enc");
      cellMap.set("formula", normalized.formula);
      cellMap.set("value", null);
    } else {
      cellMap.delete("enc");
      // Represent formula clears as an explicit marker rather than deleting the
      // key entirely. Map deletes don't create Yjs Items, which can erase causal
      // history and make delete-vs-overwrite ordering nondeterministic across
      // mixed client configurations.
      cellMap.set("formula", null);
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

  /**
   * Clear a cell while preserving its entry in the root `cells` Y.Map whenever
   * possible (writes explicit `value=null` / `formula=null` markers).
   *
   * This preserves delete-vs-overwrite causal history for downstream observers
   * (e.g. formula/value conflict monitors), since root map deletes do not create
   * Yjs Items and deep observers may ignore them.
   *
   * Encrypted cells cannot be cleared safely without an encryption key, so we
   * fall back to deleting the root entry in that case.
   *
   * @param {string} cellKey
   */
  _clearCell(cellKey) {
    const existing = this.cells.get(cellKey);
    if (existing === undefined) return;

    let cellMap = getYMap(existing);
    if (!cellMap) {
      cellMap = new Y.Map();
      this.cells.set(cellKey, cellMap);
    }

    const existingEnc = cellMap.get("enc");
    if (existingEnc !== undefined) {
      // Without keys we cannot safely overwrite ciphertext; deleting avoids
      // writing plaintext into an encrypted cell.
      this.cells.delete(cellKey);
      return;
    }

    cellMap.delete("enc");
    cellMap.set("value", null);
    cellMap.set("formula", null);
    cellMap.delete("format");
    cellMap.delete("style");
    cellMap.set("modified", Date.now());
    cellMap.set("modifiedBy", this.localUserId);
  }
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
 * @param {Map<number, number>} a
 * @param {Map<number, number>} b
 */
function stateVectorsEqual(a, b) {
  for (const [client, clock] of a.entries()) {
    if ((b.get(client) ?? 0) !== clock) return false;
  }
  for (const [client, clock] of b.entries()) {
    if ((a.get(client) ?? 0) !== clock) return false;
  }
  return true;
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
      if (prop !== "value" && prop !== "formula" && prop !== "enc" && prop !== "format" && prop !== "style") continue;
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
      // Reconstruct the "before" cell state by starting from the post-transaction view and then
      // patching in any `oldValue`s reported by Yjs for the keys that changed.
      //
      // Important: `enc` uses `undefined` (absent key) vs `null` (explicit marker) to distinguish
      // plaintext cells from encrypted/tainted ones. Do not default `enc` to `null`, otherwise
      // plaintext edits can look like encryption marker transitions and break move-vs-edit auto-merges.
      const seed = after
        ? { value: after.value ?? null, formula: after.formula ?? null, enc: after.enc, format: after.format ?? null }
        : { value: null, formula: null, enc: undefined, format: null };

      for (const [prop, change] of entry.propChanges.entries()) {
        if (prop === "value") seed.value = change.oldValue ?? null;
        if (prop === "formula") seed.formula = change.oldValue ?? null;
        if (prop === "enc") seed.enc = change.oldValue;
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

    // Treat additions and in-place edits as `edit` ops so we can detect
    // destination collisions against moves (e.g. move X->Y vs someone typing
    // into Y while offline).
    if (diff.after !== null) {
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
    // Only include cells whose structural footprint changed (ignore metadata-only
    // edits like modified timestamps).
    touchedCells: Array.from(new Set(diffs.map((d) => d.cellKey)))
  };
}

/**
 * Best-effort recovery of values from a deleted Y.Map.
 *
 * @param {any} map
 * @param {string} key
 * @returns {any}
 */
function readDeletedYMapValue(map, key) {
  // When a nested Y.Map is removed from its parent (e.g. `cells.delete(cellKey)`),
  // Yjs marks the nested map's items as deleted before deep observers run.
  // As a result, `map.get(key)` can return `undefined` even though the previous
  // value is still present on the deleted `Item` content.
  //
  // We only attempt to recover values when the map itself is deleted (its
  // integration item is deleted). For normal in-place edits, deleted entries
  // should remain `undefined`.
  if (!map || typeof map !== "object") return undefined;
  const item = map._item;
  if (!item || item.deleted !== true) return undefined;

  const entries = map._map;
  if (!(entries instanceof Map)) return undefined;
  const entry = entries.get(key);
  const content = entry?.content;
  if (content && Array.isArray(content.arr) && content.arr.length > 0) {
    return content.arr[0];
  }
  if (content && content.type !== undefined) {
    return content.type;
  }
  return undefined;
}

function readYMapValue(map, key) {
  if (!map || typeof map.get !== "function") return undefined;
  const direct = map.get(key);
  if (direct !== undefined) return direct;
  return readDeletedYMapValue(map, key);
}

function normalizeCell(cellData) {
  if (cellData == null) return null;

  /** @type {any} */
  let value = null;
  /** @type {string | null} */
  let formula = null;
  /** @type {any} */
  let enc = undefined;
  /** @type {Record<string, unknown> | null} */
  let format = null;

  const map = getYMap(cellData);
  if (map) {
    value = readYMapValue(map, "value") ?? null;
    formula = yjsValueToJson(readYMapValue(map, "formula") ?? null);
    if (formula != null) formula = String(formula);
    // Fail closed: preserve the distinction between an absent `enc` key (`undefined`)
    // and an explicit `enc: null` marker. Any defined `enc` value is treated as an
    // encryption marker so we never fall back to plaintext fields when it exists.
    enc = readYMapValue(map, "enc");
    format = (readYMapValue(map, "format") ?? readYMapValue(map, "style")) ?? null;
  } else if (typeof cellData === "object") {
    value = cellData.value ?? null;
    formula = yjsValueToJson(cellData.formula ?? null);
    if (formula != null) formula = String(formula);
    enc = Object.prototype.hasOwnProperty.call(cellData, "enc") ? cellData.enc : undefined;
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

  const hasEnc = enc !== undefined;
  const hasValue = value !== null && value !== undefined;
  const hasFormula = formula != null && formula !== "";
  const hasFormat = format != null;

  if (!hasEnc && !hasValue && !hasFormula && !hasFormat) return null;

  /** @type {NormalizedCell} */
  const out = {};
  if (hasEnc) out.enc = enc;
  else if (hasFormula) out.formula = formula;
  else if (hasValue) out.value = value;
  if (hasFormat) out.format = format;
  return out;
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
    enc: normalized.enc,
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
    stableStringify({
      value: na?.value ?? null,
      formula: normalizeFormula(na?.formula) ?? null,
      enc: na?.enc
    }) !==
    stableStringify({
      value: nb?.value ?? null,
      formula: normalizeFormula(nb?.formula) ?? null,
      enc: nb?.enc
    })
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
