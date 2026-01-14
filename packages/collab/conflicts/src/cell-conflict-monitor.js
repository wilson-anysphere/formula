import * as Y from "yjs";
import { getMapRoot } from "@formula/collab-yjs-utils";
import { cellRefFromKey } from "./cell-ref.js";

function safeCellRefFromKey(cellKey) {
  try {
    return cellRefFromKey(cellKey);
  } catch {
    return null;
  }
}

/**
 * @typedef {object} CellConflict
 * @property {string} id
 * @property {import("./cell-ref.js").CellRef} cell
 * @property {string} cellKey
 * @property {"value"} field
 * @property {any} localValue
 * @property {any} remoteValue
 * @property {string} remoteUserId Best-effort id of the remote user. May be an empty string when unavailable.
 * @property {number} detectedAt
 */

/**
 * Watches a Yjs spreadsheet document for true cell value conflicts (concurrent,
 * same-cell value edits), auto-resolves when the values are deeply equal, and
 * emits events for the rest.
 *
 * Expected document shape:
 * - doc.getMap("cells") -> Y.Map<cellKey, Y.Map>
 * - Each cell's Y.Map stores:
 *   - "value": any
 *   - "modified": number (timestamp)
 *   - "modifiedBy": string (user id)
 */
export class CellConflictMonitor {
  /**
   * @param {object} opts
   * @param {Y.Doc} opts.doc
   * @param {string} opts.localUserId
   * @param {Y.Map<any>} [opts.cells]
   * @param {object} [opts.origin] Origin token used for local transactions.
   * @param {Set<any>} [opts.localOrigins] Origins treated as local:
   *   - Conflicts are not emitted for these transactions.
   *   - Local edit tracking is updated from observed Yjs changes so later remote
   *     overwrites can be detected causally (even if callers don't use
   *     `setLocalValue`).
   * @param {Set<any>} [opts.ignoredOrigins] Transaction origins to ignore entirely.
   * @param {(conflict: CellConflict) => void} opts.onConflict
   */
  constructor(opts) {
    this.doc = opts.doc;
    this.cells = opts.cells ?? getMapRoot(this.doc, "cells");
    this.localUserId = opts.localUserId;

    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);
    this.ignoredOrigins = opts.ignoredOrigins ?? new Set();

    this.onConflict = opts.onConflict;

    /** @type {Map<string, { value: any, itemId: { client: number, clock: number } }>} */
    this._lastLocalEditByCellKey = new Map();

    /** @type {Map<string, CellConflict>} */
    this._conflicts = new Map();

    this._onDeepEvent = this._onDeepEvent.bind(this);
    this.cells.observeDeep(this._onDeepEvent);
  }

  dispose() {
    this.cells.unobserveDeep(this._onDeepEvent);
  }

  /** @returns {Array<CellConflict>} */
  listConflicts() {
    return Array.from(this._conflicts.values());
  }

  /**
   * Apply a value edit for the local user.
   *
   * @param {string} cellKey
   * @param {any} value
   */
  setLocalValue(cellKey, value) {
    const cell = this._ensureCell(cellKey);
    const nextValue = value ?? null;
    const ts = Date.now();
    const localClientId = this.doc.clientID;
    const startClock = Y.getState(this.doc.store, localClientId);

    this.doc.transact(() => {
      cell.set("value", nextValue);
      // Store a null marker rather than deleting so subsequent formula writes can
      // causally reference this clear via Item.origin. Yjs map deletes do not
      // create a new Item, which makes delete-vs-overwrite concurrency ambiguous
      // across mixed client configurations (e.g. some clients running formula
      // conflict monitors and some not).
      cell.set("formula", null);
      cell.set("modified", ts);
      cell.set("modifiedBy", this.localUserId);
    }, this.origin);

    this._lastLocalEditByCellKey.set(cellKey, { value: nextValue, itemId: { client: localClientId, clock: startClock } });
  }

  /**
   * Resolve a previously emitted value conflict by writing the chosen value back
   * into the shared Yjs doc.
   *
   * @param {string} conflictId
   * @param {any} chosenValue
   * @returns {boolean}
   */
  resolveConflict(conflictId, chosenValue) {
    const conflict = this._conflicts.get(conflictId);
    if (!conflict) return false;

    const normalizedChosen = chosenValue ?? null;
    const cell = /** @type {any} */ (this.cells.get(conflict.cellKey));
    const currentValue = cell?.get?.("value") ?? null;
    // Only apply a write if the chosen value differs from the current doc
    // state. This keeps choosing an already-applied value as a no-op while
    // still allowing resolution to re-apply if the cell changed since the
    // conflict was detected.
    if (!valuesDeeplyEqual(normalizedChosen, currentValue)) this.setLocalValue(conflict.cellKey, normalizedChosen);
    this._conflicts.delete(conflictId);
    return true;
  }

  /**
   * @param {Array<any>} events
   * @param {Y.Transaction} transaction
   */
  _onDeepEvent(events, transaction) {
    if (this.ignoredOrigins?.has(transaction.origin)) return;
    for (const event of events) {
      if (!event?.changes?.keys) continue;
      const path = event.path ?? [];
      const cellKey = path[0];
      if (typeof cellKey !== "string") continue;

      const change = event.changes.keys.get("value");
      if (!change) continue;

      const cellMap = /** @type {Y.Map<any>} */ (event.target);
      const modifiedByChange = event.changes.keys.get("modifiedBy");
      const oldValue = change.oldValue ?? null;
      const newValue = cellMap.get("value") ?? null;
      const currentModifiedBy = (cellMap.get("modifiedBy") ?? "").toString();
      // `modifiedBy` is best-effort metadata. Some writers may not update it.
      // If it didn't change in this transaction, we can't reliably attribute the overwrite.
      const remoteUserId = modifiedByChange ? currentModifiedBy : "";
      const oldModifiedBy = modifiedByChange ? (modifiedByChange.oldValue ?? "").toString() : currentModifiedBy;
      const action = change.action;
      const itemId = getItemId(cellMap, "value");
      const newItemOriginId = getItemOriginId(cellMap, "value");
      const itemLeftId = getItemLeftId(cellMap, "value");
      const currentFormula = (cellMap.get("formula") ?? "").toString();

      this._handleValueChange({
        cellKey,
        oldValue,
        newValue,
        action,
        remoteUserId,
        oldModifiedBy,
        origin: transaction.origin,
        itemId,
        newItemOriginId,
        itemLeftId,
        currentFormula
      });
    }
  }

  /**
   * @param {object} input
   * @param {string} input.cellKey
   * @param {any} input.oldValue
   * @param {any} input.newValue
   * @param {"add" | "update" | "delete"} input.action
   * @param {string} input.remoteUserId
   * @param {string} [input.oldModifiedBy]
   * @param {any} input.origin
   * @param {{ client: number, clock: number } | null} input.itemId
   * @param {{ client: number, clock: number } | null} input.newItemOriginId
   * @param {{ client: number, clock: number } | null} [input.itemLeftId]
   * @param {string} [input.currentFormula]
   */
  _handleValueChange(input) {
    const {
      cellKey,
      oldValue,
      newValue,
      action,
      remoteUserId,
      oldModifiedBy = "",
      origin,
      itemId,
      newItemOriginId,
      itemLeftId = null,
      currentFormula = ""
    } = input;

    const isLocal = this.localOrigins.has(origin);
    if (isLocal) {
      // Track local-origin edits even when callers don't use `setLocalValue`
      // (e.g. binder-style writes). This enables deterministic causal conflict
      // detection even when `modifiedBy` metadata is absent/legacy.
      //
      // Value changes that accompany formula writes are represented as value=null
      // markers and should not be tracked as local value edits.
      if (currentFormula.trim()) return;
      if (itemId) this._lastLocalEditByCellKey.set(cellKey, { value: newValue, itemId });
      return;
    }

    const lastLocal = this._lastLocalEditByCellKey.get(cellKey);
    if (!lastLocal) {
      // Fallback for monitor restarts (e.g. app reload): infer "this was my local
      // value" from the previous `modifiedBy` value and the overwritten value
      // itself.
      if (oldModifiedBy !== this.localUserId) return;
      // Deletes don't create new Items, so we can't reliably distinguish sequential
      // deletes from concurrent clears without in-memory item ids.
      if (action === "delete") return;
      // Value changes that accompany formula writes are represented as value=null
      // markers; don't treat those as value conflicts.
      if (currentFormula.trim()) return;

      // Sequential overwrite (remote saw our write) - ignore.
      if (itemLeftId && idsEqual(newItemOriginId, itemLeftId)) return;

      // Auto-resolve when the values are deep-equal.
      if (valuesDeeplyEqual(newValue, oldValue)) return;

      const cell = safeCellRefFromKey(cellKey);
      if (!cell) return;
      const conflict = /** @type {CellConflict} */ ({
        id: crypto.randomUUID(),
        cell,
        cellKey,
        field: "value",
        localValue: oldValue,
        remoteValue: newValue,
        remoteUserId,
        detectedAt: Date.now()
      });

      this._conflicts.set(conflict.id, conflict);
      this.onConflict(conflict);
      return;
    }

    // Did this remote update overwrite the last value we wrote locally?
    if (!valuesDeeplyEqual(oldValue, lastLocal.value)) return;

    // Sequential delete: remote explicitly deleted the exact item we wrote.
    // Map deletes don't create a new Item, so we can't use origin ids like we do for overwrites.
    if (action === "delete" && idsEqual(itemId, lastLocal.itemId)) {
      this._lastLocalEditByCellKey.delete(cellKey);
      return;
    }

    // Sequential overwrite (remote saw our write) - ignore.
    if (idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalEditByCellKey.delete(cellKey);
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    this._lastLocalEditByCellKey.delete(cellKey);

    // Auto-resolve when the values are deep-equal.
    if (valuesDeeplyEqual(newValue, lastLocal.value)) {
      return;
    }

    const cell = safeCellRefFromKey(cellKey);
    if (!cell) return;
    const conflict = /** @type {CellConflict} */ ({
      id: crypto.randomUUID(),
      cell,
      cellKey,
      field: "value",
      localValue: lastLocal.value,
      remoteValue: newValue,
      remoteUserId,
      detectedAt: Date.now()
    });

    this._conflicts.set(conflict.id, conflict);
    this.onConflict(conflict);
  }

  /**
   * @param {string} cellKey
   * @returns {Y.Map<any>}
   */
  _ensureCell(cellKey) {
    let cell = /** @type {Y.Map<any>|undefined} */ (this.cells.get(cellKey));
    if (!cell) {
      cell = new Y.Map();
      this.cells.set(cellKey, cell);
    }
    return cell;
  }
}

/**
 * Extract the original `origin` id for the currently visible value of a Y.Map key.
 *
 * @param {Y.Map<any>} ymap
 * @param {string} key
 * @returns {{ client: number, clock: number } | null}
 */
function getItemOriginId(ymap, key) {
  // @ts-ignore - accessing Yjs internals
  const item = ymap?._map?.get?.(key);
  if (!item) return null;
  const origin = item.origin;
  if (!origin || typeof origin !== "object") return null;
  const client = origin.client;
  const clock = origin.clock;
  if (typeof client !== "number" || typeof clock !== "number") return null;
  return { client, clock };
}

/**
 * Extract the item id for the currently visible (or most recent tombstoned) value of a Y.Map key.
 *
 * @param {Y.Map<any>} ymap
 * @param {string} key
 * @returns {{ client: number, clock: number } | null}
 */
function getItemId(ymap, key) {
  // @ts-ignore - accessing Yjs internals
  const item = ymap?._map?.get?.(key);
  if (!item) return null;
  const id = item.id;
  if (!id || typeof id !== "object") return null;
  const client = id.client;
  const clock = id.clock;
  if (typeof client !== "number" || typeof clock !== "number") return null;
  return { client, clock };
}

/**
 * Extract the id for the Item immediately to the left of the currently visible
 * value of a Y.Map key.
 *
 * @param {Y.Map<any>} ymap
 * @param {string} key
 * @returns {{ client: number, clock: number } | null}
 */
function getItemLeftId(ymap, key) {
  // @ts-ignore - accessing Yjs internals
  const item = ymap?._map?.get?.(key);
  if (!item) return null;

  /** @type {any} */
  const left = item.left;
  if (!left) return null;

  const id = left.lastId ?? left.id;
  if (!id || typeof id !== "object") return null;
  const client = id.client;
  const clock = id.clock;
  if (typeof client !== "number" || typeof clock !== "number") return null;
  return { client, clock };
}

/**
 * @param {{ client: number, clock: number } | null | undefined} a
 * @param {{ client: number, clock: number } | null | undefined} b
 */
function idsEqual(a, b) {
  if (!a || !b) return false;
  return a.client === b.client && a.clock === b.clock;
}

/**
 * Deep equality suitable for plain JSON-ish cell values.
 *
 * @param {any} a
 * @param {any} b
 * @param {Map<any, any>} [seen]
 */
function valuesDeeplyEqual(a, b, seen = new Map()) {
  if (Object.is(a, b)) return true;
  if (typeof a !== typeof b) return false;
  if (a === null || b === null) return false;

  const type = typeof a;
  if (type !== "object") return false;

  if (a instanceof Date && b instanceof Date) {
    return a.getTime() === b.getTime();
  }

  if (a instanceof Uint8Array && b instanceof Uint8Array) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  }

  const prior = seen.get(a);
  if (prior !== undefined) {
    return prior === b;
  }
  seen.set(a, b);

  const isArrayA = Array.isArray(a);
  const isArrayB = Array.isArray(b);
  if (isArrayA !== isArrayB) return false;

  if (isArrayA) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (!valuesDeeplyEqual(a[i], b[i], seen)) return false;
    }
    return true;
  }

  const keysA = Object.keys(a);
  const keysB = Object.keys(b);
  if (keysA.length !== keysB.length) return false;
  keysA.sort();
  keysB.sort();
  for (let i = 0; i < keysA.length; i += 1) {
    if (keysA[i] !== keysB[i]) return false;
  }
  for (const key of keysA) {
    // eslint-disable-next-line no-prototype-builtins
    if (!Object.prototype.hasOwnProperty.call(b, key)) return false;
    if (!valuesDeeplyEqual(a[key], b[key], seen)) return false;
  }
  return true;
}
