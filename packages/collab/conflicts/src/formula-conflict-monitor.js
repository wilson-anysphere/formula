import * as Y from "yjs";
import { getMapRoot } from "@formula/collab-yjs-utils";
import { normalizeFormula } from "../../../versioning/src/formula/normalize.js";
import { resolveFormulaConflict } from "./formula-conflict-resolver.js";
import { cellRefFromKey } from "./cell-ref.js";
import { tryEvaluateFormula } from "./formula-eval.js";

function safeCellRefFromKey(cellKey) {
  try {
    const ref = cellRefFromKey(cellKey);
    const sheetId = ref?.sheetId;
    const row = ref?.row;
    const col = ref?.col;
    if (typeof sheetId !== "string" || sheetId === "") return null;
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return ref;
  } catch {
    return null;
  }
}

/**
 * @typedef {object} FormulaConflictBase
 * @property {string} id
 * @property {import("./cell-ref.js").CellRef} cell
 * @property {string} cellKey
 * @property {string} remoteUserId Best-effort id of the remote user. May be an empty string when unavailable.
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
 * @typedef {{ type: "formula", formula: string, preview?: any } | { type: "value", value: any }} CellContentChoice
 */

/**
 * @typedef {FormulaConflictBase & {
 *   kind: "content",
 *   local: CellContentChoice,
 *   remote: CellContentChoice
 * }} ContentConflict
 */

/**
 * @typedef {FormulaTextConflict | ValueConflict | ContentConflict} FormulaConflict
 */

/**
 * @typedef {{ client: number, clock: number }} ItemId
 */

/**
 * @typedef {{ kind: "formula", formula: string, formulaItemId: ItemId, valueItemId: ItemId }
 *   | { kind: "value", value: any, valueItemId: ItemId, formulaItemId: ItemId }} LocalContentEdit
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
 *
 * Note: To detect delete-vs-overwrite conflicts deterministically, local formula
 * clears are written as `cell.set("formula", null)` (not `cell.delete("formula")`)
 * so Yjs creates an Item that later overwrites can reference via `origin`.
 *
  * Likewise, value writes clear formulas via `cell.set("formula", null)` (not
  * `cell.delete("formula")`) so later formula writes can causally reference the
  * clear via `Item.origin`. This is required even when running in formula-only
  * mode so other collaborators (who may have `mode: "formula+value"` enabled)
  * can deterministically reason about concurrent formula-vs-value edits.
 */
export class FormulaConflictMonitor {
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
   *     `setLocalFormula` / `setLocalValue`).
   * @param {Set<any>} [opts.ignoredOrigins] Transaction origins to ignore entirely.
   * @param {(conflict: FormulaConflict) => void} opts.onConflict
   * @param {(ref: { sheetId: string, row: number, col: number }) => any} [opts.getCellValue]
   * @param {number} [opts.concurrencyWindowMs] Deprecated/ignored. Former wall-clock heuristic.
   * @deprecated
   * @param {"formula" | "formula+value"} [opts.mode]
    * @param {boolean} [opts.includeValueConflicts] Backwards-compatible alias for `mode: "formula+value"`.
    */
  constructor(opts) {
    this.doc = opts.doc;
    this.cells = opts.cells ?? getMapRoot(this.doc, "cells");
    this.localUserId = opts.localUserId;

    this.origin = opts.origin ?? { type: "local" };
    this.localOrigins = opts.localOrigins ?? new Set([this.origin]);
    this.ignoredOrigins = opts.ignoredOrigins ?? new Set();

    this.onConflict = opts.onConflict;
    this.getCellValue = opts.getCellValue ?? null;

    /** @type {"formula" | "formula+value"} */
    this.mode = opts.mode ?? (opts.includeValueConflicts ? "formula+value" : "formula");
    this.includeValueConflicts = this.mode === "formula+value";

    /** @type {Map<string, { formula: string, itemId: { client: number, clock: number } | null }>} */
    this._lastLocalFormulaEditByCellKey = new Map();

    /** @type {Map<string, { value: any, itemId: { client: number, clock: number } }>} */
    this._lastLocalValueEditByCellKey = new Map();

    /** @type {Map<string, LocalContentEdit>} */
    this._lastLocalContentEditByCellKey = new Map();

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
   * Writes `value=null` alongside the formula since formula cells don't sync a
   * computed value; this marks the cell dirty for local recalculation.
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

    this._lastLocalContentEditByCellKey.set(cellKey, {
      kind: "formula",
      formula: nextFormula,
      formulaItemId: { client: localClientId, clock: startClock },
      valueItemId: { client: localClientId, clock: startClock + 1 }
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
      // Store a null marker rather than deleting so subsequent formula writes can
      // causally reference this clear via Item.origin. Yjs map deletes do not
      // create a new Item, which makes delete-vs-overwrite concurrency ambiguous
      // across mixed client configurations.
      cell.set("formula", null);
      cell.set("modified", ts);
      cell.set("modifiedBy", this.localUserId);
    }, this.origin);

    if (this.includeValueConflicts) {
      this._lastLocalValueEditByCellKey.set(cellKey, { value: nextValue, itemId: { client: localClientId, clock: startClock } });

      // Track the full "content" change so we can detect formula-vs-value conflicts.
      this._lastLocalContentEditByCellKey.set(cellKey, {
        kind: "value",
        value: nextValue,
        valueItemId: { client: localClientId, clock: startClock },
        formulaItemId: { client: localClientId, clock: startClock + 1 }
      });
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

    const cell = /** @type {any} */ (this.cells.get(conflict.cellKey));
    const currentFormula = (cell?.get?.("formula") ?? "").toString();
    const currentValue = cell?.get?.("value") ?? null;

    if (conflict.kind === "content") {
      const choice = /** @type {any} */ (chosen);
      if (!choice || typeof choice !== "object") return false;

      if (choice.type === "formula") {
        const chosenFormula = String(choice.formula ?? "");
        const formulaAlready = formulasRoughlyEqual(currentFormula, chosenFormula);
        const valueAlreadyCleared = currentValue === null;
        if (!(formulaAlready && valueAlreadyCleared)) this.setLocalFormula(conflict.cellKey, chosenFormula);
      } else if (choice.type === "value") {
        const chosenValue = choice.value ?? null;
        const valueAlready = valuesDeeplyEqual(currentValue, chosenValue);
        const formulaAlreadyCleared = formulasRoughlyEqual(currentFormula, "");
        if (!(valueAlready && formulaAlreadyCleared)) this.setLocalValue(conflict.cellKey, chosenValue);
      } else {
        return false;
      }
      this._conflicts.delete(conflictId);
      return true;
    }

    if (conflict.kind === "value") {
      const normalizedChosen = chosen ?? null;
      // Only apply a write if the chosen value differs from the current doc
      // state. This keeps choosing an already-applied value as a no-op while
      // still allowing resolution to re-apply if the cell changed since the
      // conflict was detected.
      if (!valuesDeeplyEqual(normalizedChosen, currentValue)) this.setLocalValue(conflict.cellKey, normalizedChosen);
    } else {
      const chosenFormula = String(chosen ?? "");
      // Apply the chosen formula if it differs from the current doc state, or if
      // the cell still holds a stale literal value alongside the formula.
      const formulaAlready = formulasRoughlyEqual(chosenFormula, currentFormula);
      const valueAlreadyCleared = currentValue === null;
      if (!(formulaAlready && valueAlreadyCleared)) this.setLocalFormula(conflict.cellKey, chosenFormula);
    }
    this._conflicts.delete(conflictId);
    return true;
  }

  /**
   * @param {Array<any>} events
   * @param {Y.Transaction} transaction
   */
  _onDeepEvent(events, transaction) {
    if (this.ignoredOrigins?.has(transaction.origin)) return;
    const isLocalTransaction = this.localOrigins.has(transaction.origin);
    for (const event of events) {
      // We only care about map key changes on the cell-level Y.Map.
      if (!event?.changes?.keys) continue;
      const path = event.path ?? [];
      const cellKey = path[0];
      if (typeof cellKey !== "string") continue;

      const cellMap = /** @type {Y.Map<any>} */ (event.target);
      const modifiedByChange = event.changes.keys.get("modifiedBy");
      const currentModifiedBy = (cellMap.get("modifiedBy") ?? "").toString();
      // `modifiedBy` is best-effort metadata. Some writers may not update it.
      // If it didn't change in this transaction, we can't reliably attribute the overwrite.
      const remoteUserId = modifiedByChange ? currentModifiedBy : "";
      const oldModifiedBy = modifiedByChange ? (modifiedByChange.oldValue ?? "").toString() : currentModifiedBy;
      const valueChange = event.changes.keys.get("value");
      const formulaChange = event.changes.keys.get("formula");
 
      // Local-origin transactions should not emit conflicts, but they *should* be
      // recorded so we can later detect true offline concurrent overwrites via
      // causality (even when the writer didn't call `setLocalFormula` /
      // `setLocalValue`, e.g. binder-style writes).
      if (isLocalTransaction) {
        this._trackLocalOriginChange({
          cellKey,
          cellMap,
          formulaChange,
          valueChange
        });
        continue;
      }
 
      if (formulaChange) {
        const oldFormula = (formulaChange.oldValue ?? "").toString();
        const newFormula = (cellMap.get("formula") ?? "").toString();
        const action = formulaChange.action;
        const itemId = getItemId(cellMap, "formula");
        const newItemOriginId = getItemOriginId(cellMap, "formula");
        const itemLeftId = getItemLeftId(cellMap, "formula");
        const hasValueChange = Boolean(valueChange);
        const oldValue = valueChange?.oldValue ?? null;
        const currentValue = cellMap.get("value") ?? null;

        this._handleFormulaChange({
          cellKey,
          oldFormula,
          newFormula,
          action,
          remoteUserId,
          oldModifiedBy,
          origin: transaction.origin,
          itemId,
          newItemOriginId,
          itemLeftId,
          hasValueChange,
          oldValue,
          currentValue
        });
      }

      if (this.includeValueConflicts) {
        if (valueChange) {
          const oldValue = valueChange.oldValue ?? null;
          const newValue = cellMap.get("value") ?? null;
          const action = valueChange.action;
          const itemId = getItemId(cellMap, "value");
          const newItemOriginId = getItemOriginId(cellMap, "value");
          const itemLeftId = getItemLeftId(cellMap, "value");
          const currentFormula = (cellMap.get("formula") ?? "").toString();
          // When the formula key changes in the same transaction, use its oldValue to
          // reconstruct what the cell formula looked like before this remote overwrite.
          // Otherwise, the formula key is unchanged, so the current value is also the old value.
          const oldFormula = formulaChange ? (formulaChange.oldValue ?? "").toString() : currentFormula;

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
            currentFormula,
            oldFormula
          });
        }
      }
    }
  }

  /**
   * Record a local-origin edit (as observed via Yjs events) into the in-memory
   * local edit logs used for causal conflict detection.
   *
   * This enables correct conflict detection even when the local edit was written
   * by another subsystem using the correct local transaction origin (e.g. binder
   * writes) instead of calling `setLocalFormula` / `setLocalValue`.
   *
   * @param {object} input
   * @param {string} input.cellKey
   * @param {Y.Map<any>} input.cellMap
   * @param {any} [input.formulaChange]
   * @param {any} [input.valueChange]
   */
  _trackLocalOriginChange(input) {
    const { cellKey, cellMap, formulaChange, valueChange } = input;

    const hasFormulaChange = Boolean(formulaChange);
    const hasValueChange = Boolean(valueChange);
    if (!hasFormulaChange && !hasValueChange) return;

    // Normalize post-transaction visible state. Yjs stores formula clears as
    // `null` (or key deletes); track those as an empty string to match the
    // `setLocalFormula` API semantics.
    const nextFormula = (cellMap.get("formula") ?? "").toString().trim();
    const nextValue = cellMap.get("value") ?? null;

    const formulaItemId = hasFormulaChange ? getItemId(cellMap, "formula") : null;
    const valueItemId = hasValueChange ? getItemId(cellMap, "value") : null;

    // In formula-only mode, we still need to track formula edits (including clears)
    // even when they occur alongside value writes (e.g. value writes clearing
    // formulas via `formula=null` markers). Value edits themselves are only
    // tracked when value conflict detection is enabled.
    if (!this.includeValueConflicts) {
      if (hasFormulaChange) {
        this._lastLocalFormulaEditByCellKey.set(cellKey, { formula: nextFormula, itemId: formulaItemId });
      }
      return;
    }

    /** @type {"formula" | "value" | null} */
    let kind = null;
    if (hasFormulaChange && !hasValueChange) kind = "formula";
    else if (!hasFormulaChange && hasValueChange) kind = "value";
    else if (hasFormulaChange && hasValueChange) {
      // Prefer the post-transaction visible content shape.
      if (nextFormula) {
        kind = "formula";
      } else if (nextValue !== null) {
        kind = "value";
      } else {
        // Both formula and value cleared. Disambiguate based on what the cell
        // previously contained and, when possible, the write ordering.
        //
        // In the common case:
        // - `setLocalFormula("")` writes formula first, then value.
        // - `setLocalValue(null)` writes value first, then formula.
        //
        // Binder-style writers may set `formula=null` before `value=null`, so
        // ordering alone is not sufficient when the prior formula was empty.
        const priorFormula = (formulaChange?.oldValue ?? "").toString().trim();
        if (!priorFormula) {
          // Treat clears on value-cells as value edits (also matches binder-style ordering).
          kind = "value";
        } else if (formulaChange?.action === "delete") {
          // Some legacy/edge writers may still clear formulas via key deletion.
          // Treat those clears as part of the value edit so we don't accidentally
          // record them as local formula edits.
          kind = "value";
        } else if (formulaItemId && valueItemId && formulaItemId.client === valueItemId.client) {
          // Infer whether this was a formula edit or value edit based on which key
          // was written first in the transaction.
          kind = valueItemId.clock < formulaItemId.clock ? "value" : "formula";
        } else {
          // Fall back to treating it as a formula clear when we can't infer order.
          kind = "formula";
        }
      }
    }

    if (kind === "formula") {
      if (!hasFormulaChange) return;
      this._lastLocalFormulaEditByCellKey.set(cellKey, { formula: nextFormula, itemId: formulaItemId });
      if (formulaItemId && valueItemId) {
        this._lastLocalContentEditByCellKey.set(cellKey, {
          kind: "formula",
          formula: nextFormula,
          formulaItemId,
          valueItemId
        });
      }
      return;
    }

    if (kind === "value") {
      // Value edits are only tracked when value conflict detection is enabled.
      if (!this.includeValueConflicts) return;
      if (!hasValueChange) return;
      if (valueItemId) this._lastLocalValueEditByCellKey.set(cellKey, { value: nextValue, itemId: valueItemId });

      if (valueItemId && formulaItemId) {
        this._lastLocalContentEditByCellKey.set(cellKey, { kind: "value", value: nextValue, valueItemId, formulaItemId });
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
   * @param {string} [input.oldModifiedBy]
   * @param {any} input.origin
   * @param {{ client: number, clock: number } | null} input.itemId
   * @param {{ client: number, clock: number } | null} input.newItemOriginId
   * @param {{ client: number, clock: number } | null} [input.itemLeftId]
   * @param {boolean} [input.hasValueChange]
   * @param {any} [input.oldValue]
   * @param {any} [input.currentValue]
   */
  _handleFormulaChange(input) {
    const {
      cellKey,
      oldFormula,
      newFormula,
      action,
      remoteUserId,
      oldModifiedBy = "",
      origin,
      itemId,
      newItemOriginId,
      itemLeftId = null,
      hasValueChange = false,
      oldValue = null,
      currentValue = null
    } = input;

    // When value-conflict mode is enabled, local value writes clear formulas via
    // `formula=null` (creating an Item id). If a remote user concurrently writes
    // a formula, surface a content conflict that presents the local value vs the
    // remote formula, rather than a confusing "value null vs value" conflict.
    if (this.includeValueConflicts) {
      const lastContent = this._lastLocalContentEditByCellKey.get(cellKey);
      if (lastContent?.kind === "value") {
        const remoteFormula = newFormula.trim();
        if (remoteFormula) {
          // Did this remote update overwrite the formula clear marker we wrote as part
          // of the value edit?
          if (!formulasRoughlyEqual(oldFormula, "")) return;

          // Sequential overwrite (remote saw our clear) - ignore.
          if (idsEqual(newItemOriginId, lastContent.formulaItemId)) {
            this._lastLocalContentEditByCellKey.delete(cellKey);
            this._lastLocalValueEditByCellKey.delete(cellKey);
            return;
          }

          // If the remote deleted the exact marker item we wrote (legacy clients may
          // still use key deletion), treat it as sequential and do not surface a conflict.
          if (action === "delete" && idsEqual(itemId, lastContent.formulaItemId)) {
            this._lastLocalContentEditByCellKey.delete(cellKey);
            this._lastLocalValueEditByCellKey.delete(cellKey);
            return;
          }

          // We no longer consider that local edit "pending" for conflict detection.
          this._lastLocalContentEditByCellKey.delete(cellKey);
          this._lastLocalValueEditByCellKey.delete(cellKey);

          const cell = safeCellRefFromKey(cellKey);
          if (!cell) return;

          const conflict = /** @type {FormulaConflict} */ ({
            id: crypto.randomUUID(),
            kind: "content",
            cell,
            cellKey,
            local: { type: "value", value: lastContent.value },
            remote: { type: "formula", formula: remoteFormula },
            remoteUserId,
            detectedAt: Date.now()
          });

          if (this.getCellValue) {
            const remotePreview = tryEvaluateFormula(remoteFormula, {
              getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
            });
            if (conflict.remote.type === "formula") {
              conflict.remote.preview = remotePreview.ok ? remotePreview.value : null;
            }
          }

          this._conflicts.set(conflict.id, conflict);
          this.onConflict(conflict);
          return;
        }
      }

      // Fallback when the monitor restarts (or local edit tracking is otherwise lost):
      // if this user last modified the cell by writing a literal value and a remote
      // user concurrently writes a formula that overwrites our `formula=null` marker,
      // surface a content conflict even though we don't have `_lastLocalContentEditByCellKey`.
      //
      // We intentionally only surface this as a content conflict when the local value
      // is non-null. When the local value is null, it's ambiguous whether the user
      // cleared the cell via a value edit or via a formula clear, so we let formula
      // conflict handling cover it instead.
      if (lastContent?.kind !== "value") {
        const remoteFormula = newFormula.trim();
        const localValue = hasValueChange ? oldValue : currentValue;
        if (
          remoteFormula &&
          localValue !== null &&
          oldModifiedBy === this.localUserId &&
          formulasRoughlyEqual(oldFormula, "")
        ) {
          // Sequential overwrite (remote saw our clear) - ignore.
          if (itemLeftId && idsEqual(newItemOriginId, itemLeftId)) return;

          const cell = safeCellRefFromKey(cellKey);
          if (!cell) return;
          const conflict = /** @type {FormulaConflict} */ ({
            id: crypto.randomUUID(),
            kind: "content",
            cell,
            cellKey,
            local: { type: "value", value: localValue },
            remote: { type: "formula", formula: remoteFormula },
            remoteUserId,
            detectedAt: Date.now()
          });

          if (this.getCellValue) {
            const remotePreview = tryEvaluateFormula(remoteFormula, {
              getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
            });
            if (conflict.remote.type === "formula") {
              conflict.remote.preview = remotePreview.ok ? remotePreview.value : null;
            }
          }

          this._conflicts.set(conflict.id, conflict);
          this.onConflict(conflict);
          return;
        }
      }
    }

    const lastLocal = this._lastLocalFormulaEditByCellKey.get(cellKey);
    if (!lastLocal) {
      // When the conflict monitor is re-created (e.g. app reload) it loses the
      // in-memory record of what formulas were written by this user. Use the
      // cell-level `modifiedBy` metadata as a best-effort fallback so we can
      // still detect true offline/concurrent overwrites of a cell last edited
      // by this user.
      //
      // Note: this is intentionally conservative - it won't fire if a client
      // doesn't write `modifiedBy`.
      if (oldModifiedBy !== this.localUserId) return;

      // Deletes never create new Items, so a delete that removes this formula
      // is necessarily sequential (remote must have seen the exact item id).
      // We only care about concurrent overwrites here.
      if (action === "delete") return;
    }

    // Did this remote update overwrite the last formula we wrote locally?
    if (lastLocal && !formulasRoughlyEqual(oldFormula, lastLocal.formula)) return;

    // Sequential delete: remote explicitly deleted the exact item we wrote.
    // Map deletes don't create a new Item, so we can't use origin ids like we do for overwrites.
    if (action === "delete" && lastLocal?.itemId && idsEqual(itemId, lastLocal.itemId)) {
      this._lastLocalFormulaEditByCellKey.delete(cellKey);
      this._lastLocalContentEditByCellKey.delete(cellKey);
      return;
    }

    // Sequential overwrite (remote saw our write) - ignore.
    if (lastLocal?.itemId && idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalFormulaEditByCellKey.delete(cellKey);
      this._lastLocalContentEditByCellKey.delete(cellKey);
      return;
    }

    // Sequential overwrite fallback (local tracking lost): compare the overwrite's
    // origin id to the currently-left item for this key (the overwritten value).
    if (!lastLocal && itemLeftId && idsEqual(newItemOriginId, itemLeftId)) {
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    if (lastLocal) this._lastLocalFormulaEditByCellKey.delete(cellKey);
    const lastContent = lastLocal ? this._lastLocalContentEditByCellKey.get(cellKey) : null;
    if (lastLocal) this._lastLocalContentEditByCellKey.delete(cellKey);

    const localFormula = oldFormula.trim();
    const remoteFormula = newFormula.trim();

    // Symmetric content conflict:
    // local formula write vs remote value write, where the remote value wins and
    // clears the formula (formula=null marker) while also writing a literal value.
    if (this.includeValueConflicts && hasValueChange && lastContent?.kind === "formula" && !remoteFormula && currentValue !== null) {
      const cell = safeCellRefFromKey(cellKey);
      if (!cell) return;
      const conflict = /** @type {FormulaConflict} */ ({
        id: crypto.randomUUID(),
        kind: "content",
        cell,
        cellKey,
        local: { type: "formula", formula: localFormula },
        remote: { type: "value", value: currentValue },
        remoteUserId,
        detectedAt: Date.now()
      });

      if (this.getCellValue) {
        const localPreview = tryEvaluateFormula(localFormula, {
          getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
        });
        if (conflict.local.type === "formula") {
          conflict.local.preview = localPreview.ok ? localPreview.value : null;
        }
      }

      this._conflicts.set(conflict.id, conflict);
      this.onConflict(conflict);
      return;
    }

    // Fallback when the monitor restarts (or local edit tracking is otherwise lost):
    // local formula write vs remote value write where the remote value wins and clears the
    // formula (formula=null marker). Without the in-memory local edit log we can still
    // surface a content conflict based on the previous `modifiedBy` value.
    if (
      this.includeValueConflicts &&
      hasValueChange &&
      !lastContent &&
      !remoteFormula &&
      currentValue !== null &&
      oldModifiedBy === this.localUserId &&
      !formulasRoughlyEqual(localFormula, "")
    ) {
      const cell = safeCellRefFromKey(cellKey);
      if (!cell) return;
      const conflict = /** @type {FormulaConflict} */ ({
        id: crypto.randomUUID(),
        kind: "content",
        cell,
        cellKey,
        local: { type: "formula", formula: localFormula },
        remote: { type: "value", value: currentValue },
        remoteUserId,
        detectedAt: Date.now()
      });

      if (this.getCellValue) {
        const localPreview = tryEvaluateFormula(localFormula, {
          getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
        });
        if (conflict.local.type === "formula") {
          conflict.local.preview = localPreview.ok ? localPreview.value : null;
        }
      }

      this._conflicts.set(conflict.id, conflict);
      this.onConflict(conflict);
      return;
    }

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

    const cell = safeCellRefFromKey(cellKey);
    if (!cell) return;

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
   * @param {"add" | "update" | "delete"} input.action
   * @param {string} input.remoteUserId
   * @param {string} [input.oldModifiedBy]
   * @param {any} input.origin
   * @param {{ client: number, clock: number } | null} input.itemId
   * @param {{ client: number, clock: number } | null} input.newItemOriginId
   * @param {{ client: number, clock: number } | null} [input.itemLeftId]
   * @param {string} [input.currentFormula]
   * @param {string} [input.oldFormula]
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
      currentFormula = "",
      oldFormula = ""
    } = input;

    const isLocal = this.localOrigins.has(origin);
    if (isLocal) return;

    // Content conflict (legacy/split-key): local formula write (value=null) vs a
    // remote value write that overwrote the value key without clearing the formula key.
    // This can happen with legacy clients that clear formulas via key deletion:
    // deleting a missing `formula` key is a no-op and won't conflict with a concurrent
    // formula insert, leaving both `formula` and `value` present.
    if (this.includeValueConflicts) {
      const lastContent = this._lastLocalContentEditByCellKey.get(cellKey);
      if (lastContent?.kind === "formula" && valuesDeeplyEqual(oldValue, null) && newValue !== null) {
        // Sequential delete: remote explicitly deleted the exact item we wrote.
        if (action === "delete" && idsEqual(itemId, lastContent.valueItemId)) {
          this._lastLocalContentEditByCellKey.delete(cellKey);
          this._lastLocalFormulaEditByCellKey.delete(cellKey);
          return;
        }

        // Sequential overwrite (remote saw our write) - ignore.
        if (idsEqual(newItemOriginId, lastContent.valueItemId)) {
          this._lastLocalContentEditByCellKey.delete(cellKey);
          this._lastLocalFormulaEditByCellKey.delete(cellKey);
          return;
        }

        // We no longer consider that local edit "pending" for conflict detection.
        this._lastLocalContentEditByCellKey.delete(cellKey);
        this._lastLocalFormulaEditByCellKey.delete(cellKey);

        const cell = safeCellRefFromKey(cellKey);
        if (!cell) return;
        const conflict = /** @type {FormulaConflict} */ ({
          id: crypto.randomUUID(),
          kind: "content",
          cell,
          cellKey,
          local: { type: "formula", formula: lastContent.formula },
          remote: { type: "value", value: newValue },
          remoteUserId,
          detectedAt: Date.now()
        });

        if (this.getCellValue) {
          const localPreview = tryEvaluateFormula(lastContent.formula, {
            getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
          });
          if (conflict.local.type === "formula") {
            conflict.local.preview = localPreview.ok ? localPreview.value : null;
          }
        }

        this._conflicts.set(conflict.id, conflict);
        this.onConflict(conflict);
        return;
      }

      // Fallback when the monitor restarts (or local edit tracking is otherwise lost):
      // local formula write vs legacy remote value write that overwrote the value key
      // without clearing the formula key (e.g. remote attempted `delete("formula")`
      // but the key was missing, leaving both `formula` and `value` present).
      if (
        !lastContent &&
        oldModifiedBy === this.localUserId &&
        valuesDeeplyEqual(oldValue, null) &&
        newValue !== null &&
        currentFormula.trim()
      ) {
        // Sequential overwrite (remote saw our value=null marker) - ignore.
        if (itemLeftId && idsEqual(newItemOriginId, itemLeftId)) return;

        const cell = safeCellRefFromKey(cellKey);
        if (!cell) return;
        const localFormula = currentFormula.trim();
        const conflict = /** @type {FormulaConflict} */ ({
          id: crypto.randomUUID(),
          kind: "content",
          cell,
          cellKey,
          local: { type: "formula", formula: localFormula },
          remote: { type: "value", value: newValue },
          remoteUserId,
          detectedAt: Date.now()
        });

        if (this.getCellValue) {
          const localPreview = tryEvaluateFormula(localFormula, {
            getCellValue: ({ col, row }) => this.getCellValue({ sheetId: cell.sheetId, col, row })
          });
          if (conflict.local.type === "formula") {
            conflict.local.preview = localPreview.ok ? localPreview.value : null;
          }
        }

        this._conflicts.set(conflict.id, conflict);
        this.onConflict(conflict);
        return;
      }

      // If the cell currently has a formula, value edits are either formula-side markers
      // (`value=null`) or legacy/split-key content conflicts (handled above). Don't emit a
      // standalone value conflict in that case.
      if (currentFormula.trim()) return;
    }

    const lastLocal = this._lastLocalValueEditByCellKey.get(cellKey);
    if (!lastLocal) {
      // Restart fallback: infer "this was my local value" from the previous `modifiedBy`
      // and the overwritten value itself.
      if (oldModifiedBy !== this.localUserId) return;
      if (action === "delete") return;
      // If the overwritten state had a formula, this is a content conflict
      // (formula vs value), not a value-vs-value conflict.
      if (oldFormula.trim()) return;

      // Sequential overwrite (remote saw our write) - ignore.
      if (itemLeftId && idsEqual(newItemOriginId, itemLeftId)) return;

      // Auto-resolve when the values are deep-equal.
      if (valuesDeeplyEqual(newValue, oldValue)) return;

      const cell = safeCellRefFromKey(cellKey);
      if (!cell) return;
      const conflict = /** @type {FormulaConflict} */ ({
        id: crypto.randomUUID(),
        kind: "value",
        cell,
        cellKey,
        localValue: oldValue,
        remoteValue: newValue,
        remoteUserId,
        detectedAt: Date.now()
      });

      this._conflicts.set(conflict.id, conflict);
      this.onConflict(conflict);
      return;
    }

    if (!valuesDeeplyEqual(oldValue, lastLocal.value)) return;

    // Sequential delete: remote explicitly deleted the exact item we wrote.
    // Map deletes don't create a new Item, so we can't use origin ids like we do for overwrites.
    if (action === "delete" && idsEqual(itemId, lastLocal.itemId)) {
      this._lastLocalValueEditByCellKey.delete(cellKey);
      this._lastLocalContentEditByCellKey.delete(cellKey);
      return;
    }

    // Sequential overwrite (remote saw our write) - ignore.
    if (idsEqual(newItemOriginId, lastLocal.itemId)) {
      this._lastLocalValueEditByCellKey.delete(cellKey);
      this._lastLocalContentEditByCellKey.delete(cellKey);
      return;
    }

    // We no longer consider that local edit "pending" for conflict detection.
    this._lastLocalValueEditByCellKey.delete(cellKey);
    this._lastLocalContentEditByCellKey.delete(cellKey);

    // Auto-resolve when the values are deep-equal.
    if (valuesDeeplyEqual(newValue, lastLocal.value)) return;

    const cell = safeCellRefFromKey(cellKey);
    if (!cell) return;
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
  return normalizeFormula(a) === normalizeFormula(b);
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
 * Extract the id for the Item immediately to the left of the currently visible
 * value of a Y.Map key.
 *
 * This can be used to recover the overwritten item id even when local edit
 * tracking has been lost (e.g. after recreating a monitor instance), since Yjs
 * integrates concurrent overwrites into a linked list ordered by Item ids.
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
