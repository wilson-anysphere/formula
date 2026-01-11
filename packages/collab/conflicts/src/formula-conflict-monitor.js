import * as Y from "yjs";
import { resolveFormulaConflict } from "./formula-conflict-resolver.js";
import { cellRefFromKey } from "./cell-ref.js";
import { tryEvaluateFormula } from "./formula-eval.js";

/**
 * @typedef {object} FormulaConflict
 * @property {string} id
 * @property {import("./cell-ref.js").CellRef} cell
  * @property {string} cellKey
 * @property {string} localFormula
 * @property {string} remoteFormula
 * @property {string} remoteUserId
 * @property {number} detectedAt
 * @property {any} [localPreview]
 * @property {any} [remotePreview]
 */

/**
 * Watches a Yjs spreadsheet document for true formula conflicts (concurrent,
 * same-cell edits), auto-resolves the easy cases, and emits events for the rest.
 *
 * The document shape this monitor expects (matching docs/06-collaboration.md):
 * - doc.getMap("cells") -> Y.Map<cellKey, Y.Map>
 * - Each cell's Y.Map stores:
 *   - "formula": string (optional)
 *   - "modified": number (timestamp)
 *   - "modifiedBy": string (user id)
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
    */
  constructor(opts) {
    this.doc = opts.doc;
    this.cells = opts.cells ?? this.doc.getMap("cells");
    this.localUserId = opts.localUserId;

    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);

    this.onConflict = opts.onConflict;
    this.getCellValue = opts.getCellValue ?? null;

    /** @type {Map<string, { formula: string, itemId: { client: number, clock: number } | null }>} */
    this._lastLocalEditByCellKey = new Map();

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
        cell.delete("formula");
      }
      cell.set("modified", ts);
      cell.set("modifiedBy", this.localUserId);
    }, this.origin);

    // Track locally so we can detect "remote overwrote my just-written formula".
    this._lastLocalEditByCellKey.set(cellKey, {
      formula: nextFormula,
      itemId: nextFormula ? { client: localClientId, clock: startClock } : null
    });
  }

  /**
   * Resolves a previously emitted conflict by writing the chosen formula back
   * into the shared Yjs doc.
   *
   * @param {string} conflictId
   * @param {string} chosenFormula
   * @returns {boolean}
   */
  resolveConflict(conflictId, chosenFormula) {
    const conflict = this._conflicts.get(conflictId);
    if (!conflict) return false;

    this.setLocalFormula(conflict.cellKey, chosenFormula);
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

      const change = event.changes.keys.get("formula");
      if (!change) continue;

      const cellMap = /** @type {Y.Map<any>} */ (event.target);
      const oldFormula = (change.oldValue ?? "").toString();
      const newFormula = (cellMap.get("formula") ?? "").toString();
      const remoteUserId = (cellMap.get("modifiedBy") ?? "").toString();
      const newItemOriginId = getItemOriginId(cellMap, "formula");

      this._handleFormulaChange({
        cellKey,
        oldFormula,
        newFormula,
        remoteUserId,
        origin: transaction.origin,
        newItemOriginId
      });
    }
  }

  /**
   * @param {object} input
   * @param {string} input.cellKey
   * @param {string} input.oldFormula
   * @param {string} input.newFormula
   * @param {string} input.remoteUserId
   * @param {any} input.origin
    * @param {{ client: number, clock: number } | null} input.newItemOriginId
   */
  _handleFormulaChange(input) {
    const { cellKey, oldFormula, newFormula, remoteUserId, origin, newItemOriginId } = input;

    const isLocal = this.localOrigins.has(origin);

    if (isLocal) {
      return;
    }

    const lastLocal = this._lastLocalEditByCellKey.get(cellKey);
    if (!lastLocal) return;

    // Did this remote update overwrite the last formula we wrote locally?
    if (!formulasRoughlyEqual(oldFormula, lastLocal.formula)) return;

    // Concurrency detection:
    // If the remote writer was causally aware of our local formula write, the remote Yjs Item's
    // `origin` will point at our local Item id (meaning this overwrite is sequential, not a
    // concurrent/offline conflict).
    if (lastLocal.itemId && idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalEditByCellKey.delete(cellKey);
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    this._lastLocalEditByCellKey.delete(cellKey);

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

    // True conflict - surface UI.
    const cell = cellRefFromKey(cellKey);

    const conflict = /** @type {FormulaConflict} */ ({
      id: crypto.randomUUID(),
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
 * @param {{ client: number, clock: number } | null | undefined} a
 * @param {{ client: number, clock: number } | null | undefined} b
 */
function idsEqual(a, b) {
  if (!a || !b) return false;
  return a.client === b.client && a.clock === b.clock;
}
