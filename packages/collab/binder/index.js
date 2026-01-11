import * as Y from "yjs";

import { makeCellKey, parseCellKey } from "../session/src/index.ts";

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
 * @typedef {{ value: any, formula: string | null }} NormalizedCell
 */

/**
 * @param {Y.Map<any>} cell
 * @returns {NormalizedCell}
 */
function readCellFromYjs(cell) {
  const formula = normalizeFormula(cell.get("formula") ?? null);
  if (formula) {
    return { value: null, formula };
  }
  return { value: cell.get("value") ?? null, formula: null };
}

/**
 * Bind a Yjs spreadsheet document (the `cells` root type) to a desktop `DocumentController`.
 *
 * This binder is intentionally lightweight: it syncs only cell `value` + `formula`.
 * Formatting and other workbook metadata are expected to be handled by future bindings.
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
  const localOrigin = undoService?.origin ?? null;

  /** @type {Map<string, NormalizedCell>} */
  let cache = new Map();

  let applyingRemote = false;

  /**
   * @returns {Map<string, NormalizedCell>}
   */
  function snapshotYjsCells() {
    /** @type {Map<string, NormalizedCell>} */
    const next = new Map();

    cells.forEach((cellData, key) => {
      const parsed = parseCellKey(key, { defaultSheetId });
      if (!parsed) return;

      const cell = getYMapCell(cellData);
      if (!cell) return;

      const normalized = readCellFromYjs(cell);
      if (normalized.value == null && normalized.formula == null) return;
      next.set(makeCellKey(parsed), normalized);
    });

    return next;
  }

  function applyYjsToDocumentController() {
    const next = snapshotYjsCells();

    /** @type {any[]} */
    const deltas = [];
    const allKeys = new Set([...cache.keys(), ...next.keys()]);

    for (const cellKey of allKeys) {
      const prev = cache.get(cellKey) ?? null;
      const curr = next.get(cellKey) ?? null;

      if (
        prev &&
        curr &&
        prev.formula === curr.formula &&
        Object.is(prev.value, curr.value)
      ) {
        continue;
      }

      const parsed = parseCellKey(cellKey, { defaultSheetId });
      if (!parsed) continue;

      const before = documentController.getCell(parsed.sheetId, { row: parsed.row, col: parsed.col });
      const after = {
        value: curr?.formula ? null : (curr?.value ?? null),
        formula: curr?.formula ?? null,
        styleId: before.styleId,
      };

      if (
        (before.value ?? null) === (after.value ?? null) &&
        (before.formula ?? null) === (after.formula ?? null)
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

    if (deltas.length === 0) {
      cache = next;
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
      cache = next;
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

        if (value == null && formula == null) {
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
          cache.set(cellKey, { value: null, formula });
        } else {
          cell.delete("formula");
          cell.set("value", value);
          cache.set(cellKey, { value, formula: null });
        }

        cell.set("modified", Date.now());
        if (userId) cell.set("modifiedBy", userId);
      }
    };

    if (typeof undoService?.transact === "function") {
      undoService.transact(apply);
    } else {
      ydoc.transact(apply, localOrigin ?? "document-controller:binder");
    }
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  const handleDocUpdate = (_update, origin) => {
    // Ignore local-origin transactions (those were initiated by DocumentController edits).
    if (localOrigin && origin === localOrigin) return;
    applyYjsToDocumentController();
  };

  ydoc.on("update", handleDocUpdate);

  // Initial hydration (and for cases where the provider has already applied some state).
  applyYjsToDocumentController();

  return {
    destroy() {
      unsubscribe?.();
      ydoc.off("update", handleDocUpdate);
    },
  };
}
