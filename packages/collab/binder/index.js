import * as Y from "yjs";

import { makeCellKey, normalizeCellKey, parseCellKey } from "../session/src/cell-key.js";

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

function getYMapCell(cellData) {
  // y-websocket currently pulls in the CJS build of Yjs, which means callers using
  // ESM `import * as Y from "yjs"` can observe Yjs types that fail `instanceof`
  // checks (same version, different module instance).
  //
  // Use a lightweight duck-type check instead of `instanceof Y.Map` so the
  // binder can interop with y-websocket providers in pnpm workspaces.
  if (!cellData || typeof cellData !== "object") return null;
  // eslint-disable-next-line no-prototype-builtins
  if (cellData.constructor?.name !== "YMap") return null;
  if (typeof cellData.get !== "function") return null;
  if (typeof cellData.set !== "function") return null;
  if (typeof cellData.delete !== "function") return null;
  return cellData;
}

function normalizeFormula(value) {
  if (value == null) return null;
  const trimmed = String(value).trimStart();
  if (trimmed === "") return null;
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   formatKey: string | undefined,
 * }} NormalizedCell
 */

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   format: any | undefined,
 *   formatKey: string | undefined,
 * }} ParsedYjsCell
 */

/**
 * @param {Y.Map<any>} cell
 * @returns {ParsedYjsCell}
 */
function readCellFromYjs(cell) {
  const formula = normalizeFormula(cell.get("formula") ?? null);
  let format = undefined;
  let formatKey = undefined;
  if (typeof cell.has === "function" ? cell.has("format") : cell.get("format") !== undefined) {
    format = cell.get("format") ?? null;
    formatKey = stableStringify(format);
  }
  if (formula) {
    return { value: null, formula, format, formatKey };
  }
  return { value: cell.get("value") ?? null, formula: null, format, formatKey };
}

function sameNormalizedCell(a, b) {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.formula === b.formula && Object.is(a.value, b.value) && a.formatKey === b.formatKey;
}

/**
 * Bind a Yjs spreadsheet document (the `cells` root type) to a desktop `DocumentController`.
 *
 * This binder syncs cell `value`, `formula`, and `format` (cell styles).
 * Other workbook metadata is expected to be handled by future bindings.
 *
 * @param {{
 *   ydoc: import("yjs").Doc,
 *   documentController: import("../../../apps/desktop/src/document/documentController.js").DocumentController,
 *   undoService?: { transact?: (fn: () => void) => void, origin?: any, localOrigins?: Set<any> } | null,
 *   defaultSheetId?: string,
 *   userId?: string | null,
 * }} options
 */
export function bindYjsToDocumentController(options) {
  const {
    ydoc,
    documentController,
    undoService = null,
    defaultSheetId = "Sheet1",
    userId = null,
  } = options ?? {};

  if (!ydoc) throw new Error("bindYjsToDocumentController requires { ydoc }");
  if (!documentController) throw new Error("bindYjsToDocumentController requires { documentController }");

  const cells = ydoc.getMap("cells");

  // Stable origin token for local DocumentController -> Yjs transactions when
  // we don't have a dedicated undo service wrapper.
  const binderOrigin = undoService?.origin ?? { type: "document-controller:binder" };

  /**
   * Yjs "origin" values that correspond to local DocumentController-driven writes.
   * We ignore these in the Yjs -> DocumentController observer to avoid echo work.
   *
   * Important: we intentionally do NOT ignore Y.UndoManager (used for undo/redo)
   * origins so local undo/redo still updates the DocumentController.
   * @type {Set<any>}
   */
  const localOrigins = new Set([binderOrigin]);

  if (undoService?.origin) localOrigins.add(undoService.origin);

  // Some undo services expose a `localOrigins` set which includes both the
  // origin token and the UndoManager instance. We only want to treat the origin
  // token as "local" for echo suppression.
  const maybeLocalOrigins = undoService?.localOrigins;
  if (maybeLocalOrigins && typeof maybeLocalOrigins[Symbol.iterator] === "function") {
    for (const origin of maybeLocalOrigins) {
      if (isUndoManager(origin)) continue;
      localOrigins.add(origin);
    }
  }

  /** @type {Map<string, NormalizedCell>} */
  let cache = new Map();

  let applyingRemote = false;

  /**
   * Track raw Yjs keys that correspond to a canonical `${sheetId}:${row}:${col}` key.
   *
   * This lets us apply DocumentController -> Yjs mutations without needing to scan
   * the full cells map even when the doc contains historical key encodings
   * (`${sheetId}:${row},${col}` or `r{row}c{col}`).
   *
   * @type {Map<string, Set<string>>}
   */
  const yjsKeysByCell = new Map();

  /**
   * @param {string} canonicalKey
   * @param {string} rawKey
   */
  function trackRawKey(canonicalKey, rawKey) {
    let set = yjsKeysByCell.get(canonicalKey);
    if (!set) {
      set = new Set();
      yjsKeysByCell.set(canonicalKey, set);
    }
    set.add(rawKey);
  }

  /**
   * @param {string} canonicalKey
   * @param {string} rawKey
   */
  function untrackRawKey(canonicalKey, rawKey) {
    const set = yjsKeysByCell.get(canonicalKey);
    if (!set) return;
    set.delete(rawKey);
    if (set.size === 0) yjsKeysByCell.delete(canonicalKey);
  }

  /**
   * @param {string} rawKey
   * @returns {string | null}
   */
  function canonicalKeyFromRawKey(rawKey) {
    return normalizeCellKey(rawKey, { defaultSheetId });
  }

  /**
   * @param {string} canonicalKey
   * @returns {ParsedYjsCell | null}
   */
  function readCanonicalCellFromYjs(canonicalKey) {
    const rawKeys = yjsKeysByCell.get(canonicalKey);
    let candidates;
    if (rawKeys && rawKeys.size > 0) {
      if (rawKeys.has(canonicalKey)) {
        candidates = [canonicalKey, ...Array.from(rawKeys).filter((k) => k !== canonicalKey)];
      } else {
        candidates = rawKeys;
      }
    } else {
      candidates = [canonicalKey];
    }

    for (const rawKey of candidates) {
      const cellData = cells.get(rawKey);
      const cell = getYMapCell(cellData);
      if (!cell) continue;
      const parsed = readCellFromYjs(cell);
      const hasData =
        parsed.value != null || parsed.formula != null || parsed.formatKey !== undefined;
      if (!hasData) continue;
      return parsed;
    }

    return null;
  }

  /**
   * Apply a set of canonical cell keys from Yjs into DocumentController, batching
   * into a single `applyExternalDeltas` call.
   *
   * @param {Set<string>} changedCanonicalKeys
   */
  function applyYjsChangesToDocumentController(changedCanonicalKeys) {
    if (!changedCanonicalKeys || changedCanonicalKeys.size === 0) return;

    /** @type {any[]} */
    const deltas = [];
    /** @type {Map<string, NormalizedCell | null>} */
    const nextByKey = new Map();

    for (const canonicalKey of changedCanonicalKeys) {
      const parsed = parseCellKey(canonicalKey, { defaultSheetId });
      if (!parsed) continue;

      const before = documentController.getCell(parsed.sheetId, { row: parsed.row, col: parsed.col });
      const prev = cache.get(canonicalKey) ?? null;

      const curr = readCanonicalCellFromYjs(canonicalKey);

      const currValue = curr?.formula ? null : (curr?.value ?? null);
      const currFormula = curr?.formula ?? null;

      let styleId = before.styleId;
      if (curr?.formatKey !== undefined) {
        const format = curr.format ?? null;
        styleId = format == null ? 0 : documentController.styleTable.intern(format);
      } else if (prev?.formatKey !== undefined) {
        // `format` key removed. Treat as explicit clear even though the key is now absent.
        styleId = 0;
      }

      const next =
        curr && (currValue != null || currFormula != null || curr.formatKey !== undefined)
          ? { value: currValue, formula: currFormula, formatKey: curr.formatKey }
          : null;
      nextByKey.set(canonicalKey, next);

      if (sameNormalizedCell(prev, next)) continue;

      const after = { value: currValue, formula: currFormula, styleId };
      if (
        (before.value ?? null) === (after.value ?? null) &&
        (before.formula ?? null) === (after.formula ?? null) &&
        before.styleId === after.styleId
      ) {
        continue;
      }

      deltas.push({
        sheetId: parsed.sheetId,
        row: parsed.row,
        col: parsed.col,
        before,
        after,
      });
    }

    const updateCache = () => {
      for (const [canonicalKey, next] of nextByKey.entries()) {
        if (!next) cache.delete(canonicalKey);
        else cache.set(canonicalKey, next);
      }
    };

    if (deltas.length === 0) {
      updateCache();
      return;
    }

    applyingRemote = true;
    try {
      if (typeof documentController.applyExternalDeltas === "function") {
        documentController.applyExternalDeltas(deltas);
      } else {
        // Fallback for older DocumentController versions: apply via user mutations without feedback.
        for (const delta of deltas) {
          if (delta.after.formula != null) {
            documentController.setCellFormula(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.formula);
          } else {
            documentController.setCellValue(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.value);
          }
        }
      }
    } finally {
      applyingRemote = false;
      updateCache();
    }
  }

  /**
   * One-time initial hydration of existing Yjs state into DocumentController.
   * Allowed to scan the full `cells` map exactly once at bind time.
   */
  function hydrateFromYjs() {
    /** @type {Set<string>} */
    const changedKeys = new Set();
    cells.forEach((_cellData, rawKey) => {
      const canonicalKey = canonicalKeyFromRawKey(rawKey);
      if (!canonicalKey) return;
      trackRawKey(canonicalKey, rawKey);
      changedKeys.add(canonicalKey);
    });

    applyYjsChangesToDocumentController(changedKeys);
  }

  /**
   * @param {any[]} events
   * @param {any} transaction
   */
  const handleCellsDeepChange = (events, transaction) => {
    if (!events || events.length === 0) return;
    const origin = transaction?.origin ?? null;
    if (localOrigins.has(origin)) return;

    /** @type {Set<string>} */
    const changed = new Set();

    for (const event of events) {
      // Nested map updates: event.path[0] is the root cells map key.
      const path = event?.path;
      if (Array.isArray(path) && path.length > 0) {
        const rawKey = path[0];
        if (typeof rawKey !== "string") continue;

        const changes = event?.changes?.keys;
        if (changes && !(changes.has("value") || changes.has("formula") || changes.has("format"))) continue;

        const canonicalKey = canonicalKeyFromRawKey(rawKey);
        if (!canonicalKey) continue;
        trackRawKey(canonicalKey, rawKey);
        changed.add(canonicalKey);
        continue;
      }

      // Root map changes: keys added/updated/removed.
      const changes = event?.changes?.keys;
      if (!changes) continue;
      for (const [rawKey, change] of changes.entries()) {
        if (typeof rawKey !== "string") continue;
        const canonicalKey = canonicalKeyFromRawKey(rawKey);
        if (!canonicalKey) continue;

        if (change?.action === "delete") {
          untrackRawKey(canonicalKey, rawKey);
        } else {
          trackRawKey(canonicalKey, rawKey);
        }

        changed.add(canonicalKey);
      }
    }

    applyYjsChangesToDocumentController(changed);
  };

  /**
   * @param {any} payload
   */
  const handleDocumentChange = (payload) => {
    if (applyingRemote) return;
    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    if (deltas.length === 0) return;

    const apply = () => {
      for (const delta of deltas) {
        const canonicalKey = makeCellKey({ sheetId: delta.sheetId, row: delta.row, col: delta.col });

        const value = delta.after?.value ?? null;
        const formula = normalizeFormula(delta.after?.formula ?? null);
        const styleId = Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0;
        const format = styleId === 0 ? null : documentController.styleTable.get(styleId);
        const formatKey = styleId === 0 ? undefined : stableStringify(format);

        const rawKeys = yjsKeysByCell.get(canonicalKey);
        const targets = rawKeys && rawKeys.size > 0 ? Array.from(rawKeys) : [canonicalKey];

        if (value == null && formula == null && styleId === 0) {
          for (const rawKey of targets) {
            cells.delete(rawKey);
            untrackRawKey(canonicalKey, rawKey);
          }
          cache.delete(canonicalKey);
          continue;
        }

        for (const rawKey of targets) {
          let cellData = cells.get(rawKey);
          let cell = getYMapCell(cellData);
          if (!cell) {
            cell = new Y.Map();
            cells.set(rawKey, cell);
          }

          if (formula != null) {
            cell.set("formula", formula);
            cell.set("value", null);
          } else {
            cell.delete("formula");
            cell.set("value", value);
          }

          if (styleId === 0) {
            cell.delete("format");
          } else {
            cell.set("format", format);
          }

          cell.set("modified", Date.now());
          if (userId) cell.set("modifiedBy", userId);

          trackRawKey(canonicalKey, rawKey);
        }

        const normalizedValue = formula != null ? null : value;
        const normalizedFormula = formula != null ? formula : null;
        if (normalizedValue != null || normalizedFormula != null || formatKey !== undefined) {
          cache.set(canonicalKey, { value: normalizedValue, formula: normalizedFormula, formatKey });
        } else {
          cache.delete(canonicalKey);
        }
      }
    };

    if (typeof undoService?.transact === "function") {
      undoService.transact(apply);
    } else {
      ydoc.transact(apply, binderOrigin);
    }
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  cells.observeDeep(handleCellsDeepChange);

  // Initial hydration (and for cases where the provider has already applied some state).
  hydrateFromYjs();

  return {
    destroy() {
      unsubscribe?.();
      cells.unobserveDeep(handleCellsDeepChange);
    },
  };
}

function isUndoManager(value) {
  if (!value || typeof value !== "object") return false;
  const maybe = value;
  if (maybe.constructor?.name === "UndoManager") return true;
  return typeof maybe.undo === "function" && typeof maybe.redo === "function" && maybe.trackedOrigins instanceof Set;
}

