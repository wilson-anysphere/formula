import * as Y from "yjs";

import { makeCellKey, parseCellKey } from "../session/src/index.ts";

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
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? value : null;
}

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   formatKey: string | null | undefined,
 * }} NormalizedCell
 */

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   format: any | undefined,
 *   formatKey: string | null | undefined,
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

/**
 * Bind a Yjs spreadsheet document (the `cells` root type) to a desktop `DocumentController`.
 *
 * This binder syncs cell `value`, `formula`, and `format` (cell styles).
 * Other workbook metadata is expected to be handled by future bindings.
 *
 * @param {{
 *   ydoc: import("yjs").Doc,
 *   documentController: import("../../../apps/desktop/src/document/documentController.js").DocumentController,
 *   undoService?: { transact?: (fn: () => void) => void, origin?: any } | null,
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
  const localOrigin = undoService?.origin ?? { type: "document-controller:binder" };

  /** @type {Map<string, NormalizedCell>} */
  let cache = new Map();

  let applyingRemote = false;

  /**
   * @param {string[]} cellKeys
   * @returns {any[]}
   */
  function computeExternalDeltas(cellKeys) {
    /** @type {any[]} */
    const deltas = [];

    for (const cellKey of cellKeys) {
      const parsed = parseCellKey(cellKey, { defaultSheetId });
      if (!parsed) continue;

      const before = documentController.getCell(parsed.sheetId, { row: parsed.row, col: parsed.col });

      const prev = cache.get(cellKey) ?? null;

      const cellData = cells.get(cellKey);
      const cell = getYMapCell(cellData);
      const curr = cell ? readCellFromYjs(cell) : null;

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

      const after = {
        value: currValue,
        formula: currFormula,
        styleId,
      };

      if (
        (before.value ?? null) === (after.value ?? null) &&
        (before.formula ?? null) === (after.formula ?? null) &&
        before.styleId === after.styleId
      ) {
        // Even if the Yjs cell changed (e.g. modified timestamp), avoid
        // generating a no-op external delta.
      } else {
        deltas.push({
          sheetId: parsed.sheetId,
          row: parsed.row,
          col: parsed.col,
          before,
          after,
        });
      }

      // Update cache after computing the delta. Include format-only cells.
      if (curr && (currValue != null || currFormula != null || curr.formatKey !== undefined)) {
        cache.set(cellKey, { value: currValue, formula: currFormula, formatKey: curr.formatKey });
      } else {
        cache.delete(cellKey);
      }
    }

    return deltas;
  }

  /**
   * Apply Yjs changes for the provided cell keys into the DocumentController.
   *
   * This avoids rescanning the entire Yjs `cells` map on every update by relying on
   * Yjs observeDeep events to supply the changed keys.
   *
   * @param {string[]} cellKeys
   */
  function applyYjsToDocumentController(cellKeys) {
    const deltas = computeExternalDeltas(cellKeys);
    if (deltas.length === 0) return;

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
    }
  }

  /**
   * @param {any} payload
   */
  const handleDocumentChange = (payload) => {
    if (applyingRemote) return;
    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    if (deltas.length === 0) return;

    const apply = () => {
      for (const delta of deltas) {
        const cellKey = makeCellKey({ sheetId: delta.sheetId, row: delta.row, col: delta.col });

        const value = delta.after?.value ?? null;
        const formula = normalizeFormula(delta.after?.formula ?? null);
        const styleId = Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0;
        const format = styleId === 0 ? null : documentController.styleTable.get(styleId);
        const formatKey = styleId === 0 ? undefined : stableStringify(format);

        if (value == null && formula == null && styleId === 0) {
          cells.delete(cellKey);
          cache.delete(cellKey);
          continue;
        }

        let cellData = cells.get(cellKey);
        let cell = getYMapCell(cellData);
        if (!cell) {
          cell = new Y.Map();
          cells.set(cellKey, cell);
        }

        if (formula != null) {
          cell.set("formula", formula);
          cell.set("value", null);
          cache.set(cellKey, { value: null, formula, formatKey });
        } else {
          cell.delete("formula");
          cell.set("value", value);
          cache.set(cellKey, { value, formula: null, formatKey });
        }

        if (styleId === 0) {
          cell.delete("format");
        } else {
          cell.set("format", format);
        }

        cell.set("modified", Date.now());
        if (userId) cell.set("modifiedBy", userId);
      }
    };

    if (typeof undoService?.transact === "function") {
      undoService.transact(apply);
    } else {
      ydoc.transact(apply, localOrigin);
    }
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  /**
   * Observe deep Yjs changes so we can apply only the touched cell keys to the
   * DocumentController, rather than rescanning the entire map on every update.
   *
   * @param {any[]} events
   */
  const handleYjsCellsChange = (events, transaction) => {
    if (!Array.isArray(events) || events.length === 0) return;
    const origin = transaction?.origin ?? events[0]?.transaction?.origin ?? null;
    // Ignore transactions that originated from the DocumentController -> Yjs path.
    if (origin === localOrigin) return;

    /** @type {Set<string>} */
    const changedKeys = new Set();

    for (const event of events) {
      if (!event) continue;

      // Root map changes (cell added/removed/replaced).
      if (event.target === cells) {
        const changes = event.changes?.keys;
        if (changes && typeof changes.forEach === "function") {
          changes.forEach((_change, key) => {
            if (typeof key === "string") changedKeys.add(key);
          });
        }
        continue;
      }

      // Nested cell changes (value/formula/format/etc).
      const path = event.path;
      if (Array.isArray(path) && typeof path[0] === "string") {
        changedKeys.add(path[0]);
      }
    }

    if (changedKeys.size === 0) return;
    applyYjsToDocumentController(Array.from(changedKeys));
  };

  cells.observeDeep(handleYjsCellsChange);

  // Initial hydration (and for cases where the provider has already applied some state).
  const allKeys = [];
  cells.forEach((_cellData, key) => {
    if (typeof key === "string") allKeys.push(key);
  });
  if (allKeys.length > 0) {
    applyYjsToDocumentController(allKeys);
  }

  return {
    destroy() {
      unsubscribe?.();
      cells.unobserveDeep(handleYjsCellsChange);
    },
  };
}
