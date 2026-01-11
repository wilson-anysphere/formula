import * as Y from "yjs";
import { resolveFormulaConflict } from "./formula-conflict-resolver.js";
import { cellRefFromKey } from "./cell-ref.js";
import { tryEvaluateFormula } from "./formula-eval.js";

/**
 * @typedef {object} FormulaConflictBase
 * @property {string} id
 * @property {import("./cell-ref.js").CellRef} cell
 * @property {string} cellKey
 * @property {string} remoteUserId
 * @property {number} detectedAt
 */

/**
 * @typedef {FormulaConflictBase & {
 *   kind: "formula",
 *   localFormula: string,
 *   remoteFormula: string,
 *   localPreview?: any,
 *   remotePreview?: any
 * }} FormulaTextConflict
 */

/**
 * @typedef {FormulaConflictBase & {
 *   kind: "value",
 *   localValue: any,
 *   remoteValue: any
 * }} ValueConflict
 */

/**
 * @typedef {FormulaTextConflict | ValueConflict} FormulaConflict
 */

/**
 * Watches a Yjs spreadsheet document for true formula conflicts (concurrent,
 * same-cell edits), auto-resolves the easy cases, and emits events for the rest.
 *
 * Formula conflicts are always detected. Value conflicts are optionally detected
 * when `mode: "formula+value"` / `includeValueConflicts` is enabled.
 *
 * Concurrency detection is causal and based on Yjs Item `origin` ids. If the
 * overwriting entry has an origin that points at our local entry id, then the
 * overwriter had integrated our write (sequential overwrite, not a conflict).
 */
export class FormulaConflictMonitor {
  /**
   * @param {object} opts
   * @param {Y.Doc} opts.doc
   * @param {string} opts.localUserId
   * @param {Y.Map<any>} [opts.cells]
   * @param {object} [opts.origin] Origin token used for local transactions.
   * @param {Set<any>} [opts.localOrigins] Origins treated as local (for ignoring).
   * @param {(conflict: FormulaConflict) => void} opts.onConflict
   * @param {(ref: { sheetId: string, row: number, col: number }) => any} [opts.getCellValue]
   * @param {number} [opts.concurrencyWindowMs] Deprecated (ignored). Former wall-clock heuristic.
   * @param {"formula" | "formula+value"} [opts.mode]
   * @param {boolean} [opts.includeValueConflicts] Backwards-compatible alias for `mode: "formula+value"`.
   */
  constructor(opts) {
    this.doc = opts.doc;
    this.cells = opts.cells ?? this.doc.getMap("cells");
    this.localUserId = opts.localUserId;

    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);

    this.onConflict = opts.onConflict;
    this.getCellValue = opts.getCellValue ?? null;

    /** @type {"formula" | "formula+value"} */
    this.mode = opts.mode ?? (opts.includeValueConflicts ? "formula+value" : "formula");
    this.includeValueConflicts = this.mode === "formula+value";

    /** @type {Map<string, { formula: string, itemId: { client: number, clock: number } | null }>} */
    this._lastLocalFormulaEditByCellKey = new Map();

    /** @type {Map<string, { value: any, itemId: { client: number, clock: number } }>} */
    this._lastLocalValueEditByCellKey = new Map();

    /** @type {Map<string, FormulaConflict>} */
    this._conflicts = new Map();

    this._onDeepEvent = this._onDeepEvent.bind(this);
    this.cells.observeDeep(this._onDeepEvent);
  }

  dispose() {
    this.cells.unobserveDeep(this._onDeepEvent);
  }

  /** @returns {Array<FormulaConflict>} */
  listConflicts() {
    return Array.from(this._conflicts.values());
  }

  /**
   * Apply a formula edit for the local user (this is the API we'd call from UI).
   *
   * @param {string} cellKey
   * @param {string} formula
   */
  setLocalFormula(cellKey, formula) {
    const cell = this._ensureCell(cellKey);
    const nextFormula = formula.trim();
    const ts = Date.now();
    const localClientId = this.doc.clientID;
    const startClock = Y.getState(this.doc.store, localClientId);

    this.doc.transact(() => {
      if (nextFormula) {
        cell.set("formula", nextFormula);
      } else {
        // Store a null marker rather than deleting the key so subsequent writes
        // can causally reference this deletion via Item.origin. Yjs map deletes
        // do not create a new Item, which makes delete-vs-overwrite concurrency
        // ambiguous without an explicit marker.
        cell.set("formula", null);
      }
      // Formula cells should not store a synced "value" (it's computed locally).
      // Clearing the value marks the cell dirty for the local formula engine to recompute.
      cell.set("value", null);
      cell.set("modified", ts);
      cell.set("modifiedBy", this.localUserId);
    }, this.origin);

    // Track locally so we can detect "remote overwrote my just-written formula".
    this._lastLocalFormulaEditByCellKey.set(cellKey, {
      formula: nextFormula,
      itemId: { client: localClientId, clock: startClock }
    });
  }

  /**
   * Apply a value edit for the local user.
   *
   * Values are only tracked for conflict detection when `mode: "formula+value"`
   * is enabled.
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
      cell.delete("formula");
      cell.set("modified", ts);
      cell.set("modifiedBy", this.localUserId);
    }, this.origin);

    if (this.includeValueConflicts) {
      this._lastLocalValueEditByCellKey.set(cellKey, { value: nextValue, itemId: { client: localClientId, clock: startClock } });
    }
  }

  /**
   * Resolves a previously emitted conflict by writing the chosen state back
   * into the shared Yjs doc.
   *
   * @param {string} conflictId
   * @param {any} chosen
   * @returns {boolean}
   */
  resolveConflict(conflictId, chosen) {
    const conflict = this._conflicts.get(conflictId);
    if (!conflict) return false;

    if (conflict.kind === "value") {
      this.setLocalValue(conflict.cellKey, chosen);
    } else {
      this.setLocalFormula(conflict.cellKey, String(chosen ?? ""));
    }
    this._conflicts.delete(conflictId);
    return true;
  }

  /**
   * @param {Array<any>} events
   * @param {Y.Transaction} transaction
   */
  _onDeepEvent(events, transaction) {
    for (const event of events) {
      // We only care about map key changes on the cell-level Y.Map.
      if (!event?.changes?.keys) continue;
      const path = event.path ?? [];
      const cellKey = path[0];
      if (typeof cellKey !== "string") continue;

      const cellMap = /** @type {Y.Map<any>} */ (event.target);
      const remoteUserId = (cellMap.get("modifiedBy") ?? "").toString();

      const formulaChange = event.changes.keys.get("formula");
      if (formulaChange) {
        const oldFormula = (formulaChange.oldValue ?? "").toString();
        const newFormula = (cellMap.get("formula") ?? "").toString();
        const action = formulaChange.action;
        const itemId = getItemId(cellMap, "formula");
        const newItemOriginId = getItemOriginId(cellMap, "formula");

        this._handleFormulaChange({
          cellKey,
          oldFormula,
          newFormula,
          action,
          remoteUserId,
          origin: transaction.origin,
          itemId,
          newItemOriginId
        });
      }

      if (this.includeValueConflicts) {
        const valueChange = event.changes.keys.get("value");
        if (valueChange) {
          const oldValue = valueChange.oldValue ?? null;
          const newValue = cellMap.get("value") ?? null;
          const newItemOriginId = getItemOriginId(cellMap, "value");

          this._handleValueChange({
            cellKey,
            oldValue,
            newValue,
            remoteUserId,
            origin: transaction.origin,
            newItemOriginId
          });
        }
      }
    }
  }

  /**
   * @param {object} input
   * @param {string} input.cellKey
   * @param {string} input.oldFormula
   * @param {string} input.newFormula
   * @param {"add" | "update" | "delete"} input.action
   * @param {string} input.remoteUserId
   * @param {any} input.origin
   * @param {{ client: number, clock: number } | null} input.itemId
   * @param {{ client: number, clock: number } | null} input.newItemOriginId
   */
  _handleFormulaChange(input) {
    const { cellKey, oldFormula, newFormula, action, remoteUserId, origin, itemId, newItemOriginId } = input;

    const isLocal = this.localOrigins.has(origin);
    if (isLocal) return;

    const lastLocal = this._lastLocalFormulaEditByCellKey.get(cellKey);
    if (!lastLocal) return;

    // Did this remote update overwrite the last formula we wrote locally?
    if (!formulasRoughlyEqual(oldFormula, lastLocal.formula)) return;

    // Sequential delete: remote explicitly deleted the exact item we wrote.
    // Map deletes don't create a new Item, so we can't use origin ids like we do for overwrites.
    if (action === "delete" && lastLocal.itemId && idsEqual(itemId, lastLocal.itemId)) {
      this._lastLocalFormulaEditByCellKey.delete(cellKey);
      return;
    }

    // Sequential overwrite (remote saw our write) - ignore.
    if (lastLocal.itemId && idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalFormulaEditByCellKey.delete(cellKey);
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    this._lastLocalFormulaEditByCellKey.delete(cellKey);

    const localFormula = oldFormula.trim();
    const remoteFormula = newFormula.trim();

    const decision = resolveFormulaConflict({
      localFormula,
      remoteFormula
    });

    if (decision.kind === "equivalent" || decision.kind === "prefer-remote") {
      // Remote formula is already applied in the doc.
      return;
    }

    if (decision.kind === "prefer-local") {
      // Re-apply the local extension on top of the remote write (sequentially).
      this.setLocalFormula(cellKey, localFormula);
      return;
    }

    const cell = cellRefFromKey(cellKey);

    const conflict = /** @type {FormulaConflict} */ ({
      id: crypto.randomUUID(),
      kind: "formula",
      cell,
      cellKey,
      localFormula,
      remoteFormula,
      remoteUserId,
      detectedAt: Date.now()
    });

    if (this.getCellValue) {
      const localPreview = tryEvaluateFormula(localFormula, {
        getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
      });
      const remotePreview = tryEvaluateFormula(remoteFormula, {
        getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
      });

      conflict.localPreview = localPreview.ok ? localPreview.value : null;
      conflict.remotePreview = remotePreview.ok ? remotePreview.value : null;
    }

    this._conflicts.set(conflict.id, conflict);
    this.onConflict(conflict);
  }

  /**
   * @param {object} input
   * @param {string} input.cellKey
   * @param {any} input.oldValue
   * @param {any} input.newValue
   * @param {string} input.remoteUserId
   * @param {any} input.origin
   * @param {{ client: number, clock: number } | null} input.newItemOriginId
   */
  _handleValueChange(input) {
    const { cellKey, oldValue, newValue, remoteUserId, origin, newItemOriginId } = input;

    const isLocal = this.localOrigins.has(origin);
    if (isLocal) return;

    const lastLocal = this._lastLocalValueEditByCellKey.get(cellKey);
    if (!lastLocal) return;

    if (!valuesDeeplyEqual(oldValue, lastLocal.value)) return;

    // Sequential overwrite (remote saw our write) - ignore.
    if (idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalValueEditByCellKey.delete(cellKey);
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    this._lastLocalValueEditByCellKey.delete(cellKey);

    // Auto-resolve when the values are deep-equal.
    if (valuesDeeplyEqual(newValue, lastLocal.value)) return;

    const cell = cellRefFromKey(cellKey);
    const conflict = /** @type {FormulaConflict} */ ({
      id: crypto.randomUUID(),
      kind: "value",
      cell,
      cellKey,
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
 * @param {string} a
 * @param {string} b
 */
function formulasRoughlyEqual(a, b) {
  return normalizeFormulaText(a) === normalizeFormulaText(b);
}

/**
 * @param {string} formula
 */
function normalizeFormulaText(formula) {
  const stripped = String(formula ?? "").trim().replace(/^\s*=\s*/, "");
  return stripped.replaceAll(/\s+/g, "").toUpperCase();
}

/**
 * Extract the original `origin` id for the currently visible value of a Y.Map key.
 *
 * @param {Y.Map<any>} ymap
 * @param {string} key
 * @returns {{ client: number, clock: number } | null}
 */
function getItemOriginId(ymap, key) {
  // Yjs stores key/value entries as internal `Item` structs accessible from `._map`.
  // There is no public API for retrieving causal ids for map entries today.
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
