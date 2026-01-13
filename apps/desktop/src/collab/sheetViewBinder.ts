import * as Y from "yjs";
import { getYMap, getYText } from "@formula/collab-yjs-utils";

import type { CollabSession } from "@formula/collab-session";
import type { DocumentController } from "../document/documentController.js";

export type SheetViewState = {
  frozenRows: number;
  frozenCols: number;
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
};

export type SheetViewDelta = {
  sheetId: string;
  before: SheetViewState;
  after: SheetViewState;
};

export type SheetViewBinder = { destroy: () => void };

const VIEW_KEYS = new Set(["view", "frozenRows", "frozenCols", "colWidths", "rowHeights"]);

function isRecord(value: unknown): value is Record<string, any> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function readYMapOrObject(value: any, key: string): any {
  const map = getYMap(value);
  if (map) return map.get(key);
  if (isRecord(value)) return value[key];
  return undefined;
}

function coerceString(value: unknown): string | null {
  if (typeof value === "string") return value;
  const text = getYText(value);
  if (text) return text.toString();
  if (value == null) return null;
  return String(value);
}

function normalizeFrozenCount(value: unknown): number {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

function normalizeAxisSize(value: unknown): number | null {
  const num = Number(value);
  if (!Number.isFinite(num)) return null;
  if (num <= 0) return null;
  return num;
}

function readAxisOverrides(raw: unknown): Record<string, number> | undefined {
  if (!raw) return undefined;

  const out: Record<string, number> = {};

  const map = getYMap(raw);
  if (map) {
    map.forEach((value: unknown, key: string) => {
      const idx = Number(key);
      if (!Number.isInteger(idx) || idx < 0) return;
      const normalized = normalizeAxisSize(value);
      if (normalized == null) return;
      out[String(idx)] = normalized;
    });
    return Object.keys(out).length > 0 ? out : undefined;
  }

  if (Array.isArray(raw)) {
    for (const entry of raw) {
      const index = Array.isArray(entry) ? entry[0] : (entry as any)?.index;
      const size = Array.isArray(entry) ? entry[1] : (entry as any)?.size;
      const idx = Number(index);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const normalized = normalizeAxisSize(size);
      if (normalized == null) continue;
      out[String(idx)] = normalized;
    }
    return Object.keys(out).length > 0 ? out : undefined;
  }

  if (isRecord(raw)) {
    for (const [key, value] of Object.entries(raw)) {
      const idx = Number(key);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const normalized = normalizeAxisSize(value);
      if (normalized == null) continue;
      out[String(idx)] = normalized;
    }
    return Object.keys(out).length > 0 ? out : undefined;
  }

  return undefined;
}

function sheetViewEquals(a: SheetViewState, b: SheetViewState): boolean {
  if (a === b) return true;

  const axisEquals = (left?: Record<string, number>, right?: Record<string, number>): boolean => {
    if (left === right) return true;
    const leftKeys = left ? Object.keys(left) : [];
    const rightKeys = right ? Object.keys(right) : [];
    if (leftKeys.length !== rightKeys.length) return false;
    leftKeys.sort((x, y) => Number(x) - Number(y));
    rightKeys.sort((x, y) => Number(x) - Number(y));
    for (let i = 0; i < leftKeys.length; i += 1) {
      const key = leftKeys[i]!;
      if (key !== rightKeys[i]) return false;
      const lv = left![key]!;
      const rv = right![key]!;
      if (Math.abs(lv - rv) > 1e-6) return false;
    }
    return true;
  };

  return (
    a.frozenRows === b.frozenRows &&
    a.frozenCols === b.frozenCols &&
    axisEquals(a.colWidths, b.colWidths) &&
    axisEquals(a.rowHeights, b.rowHeights)
  );
}

function getSheetIdFromSheetMap(sheet: any): string | null {
  return coerceString(sheet?.get?.("id") ?? sheet?.id);
}

function readSheetViewFromSheetMap(sheet: any): SheetViewState {
  const viewRaw = sheet?.get?.("view");

  const frozenRows =
    viewRaw !== undefined ? normalizeFrozenCount(readYMapOrObject(viewRaw, "frozenRows")) : normalizeFrozenCount(sheet?.get?.("frozenRows"));
  const frozenCols =
    viewRaw !== undefined ? normalizeFrozenCount(readYMapOrObject(viewRaw, "frozenCols")) : normalizeFrozenCount(sheet?.get?.("frozenCols"));

  const colWidths =
    viewRaw !== undefined ? readAxisOverrides(readYMapOrObject(viewRaw, "colWidths")) : readAxisOverrides(sheet?.get?.("colWidths"));
  const rowHeights =
    viewRaw !== undefined ? readAxisOverrides(readYMapOrObject(viewRaw, "rowHeights")) : readAxisOverrides(sheet?.get?.("rowHeights"));

  const out: SheetViewState = { frozenRows, frozenCols };
  if (colWidths) out.colWidths = colWidths;
  if (rowHeights) out.rowHeights = rowHeights;
  return out;
}

function ensureNestedYMap(parent: any, key: string): Y.Map<any> {
  const existing = parent?.get?.(key);
  const existingMap = getYMap(existing);
  if (existingMap) return existingMap;

  const next = new Y.Map();

  // Best-effort: if the existing value was a plain object, preserve entries.
  if (isRecord(existing)) {
    for (const [k, v] of Object.entries(existing)) {
      next.set(k, v);
    }
  }

  parent?.set?.(key, next);
  return next;
}

function applyAxisDelta(map: Y.Map<any>, before?: Record<string, number>, after?: Record<string, number>): void {
  const keys = new Set<string>();
  for (const k of Object.keys(before ?? {})) keys.add(k);
  for (const k of Object.keys(after ?? {})) keys.add(k);

  for (const key of keys) {
    const next = after?.[key];
    if (next == null) {
      if (before && key in before) map.delete(key);
      continue;
    }
    const prev = before?.[key];
    if (prev == null || Math.abs(prev - next) > 1e-6) {
      map.set(key, next);
    }
  }
}

/**
 * Bind per-sheet "view" metadata (frozen panes + axis size overrides) between a CollabSession's
 * `sheets` root and a desktop DocumentController.
 */
export function bindSheetViewToCollabSession(options: {
  session: CollabSession;
  documentController: DocumentController;
  /**
   * Optional stable Yjs transaction origin used for DocumentController -> Yjs writes.
   *
   * When omitted, a new per-binder origin token is created.
   */
  origin?: any;
}): SheetViewBinder {
  const { session, documentController } = options ?? ({} as any);
  if (!session) throw new Error("bindSheetViewToCollabSession requires { session }");
  if (!documentController) throw new Error("bindSheetViewToCollabSession requires { documentController }");

  const binderOrigin = options?.origin ?? { type: "document-controller:sheet-view-binder" };
  session.localOrigins?.add?.(binderOrigin);

  let destroyed = false;
  let applyingRemote = false;

  const findSheetMap = (sheetId: string): any | null => {
    const arr = session.sheets?.toArray?.() ?? [];
    for (const entry of arr) {
      const id = getSheetIdFromSheetMap(entry as any);
      if (id === sheetId) return entry as any;
    }
    return null;
  };

  const applyYjsToDocumentController = (sheetIds: Iterable<string>): void => {
    const deltas: SheetViewDelta[] = [];

    for (const sheetId of sheetIds) {
      const sheet = findSheetMap(sheetId);
      if (!sheet) continue;

      const after = readSheetViewFromSheetMap(sheet);
      const before = documentController.getSheetView(sheetId) as SheetViewState;
      if (sheetViewEquals(before, after)) continue;

      deltas.push({ sheetId, before, after });
    }

    if (deltas.length === 0) return;

    applyingRemote = true;
    try {
      if (typeof (documentController as any).applyExternalSheetViewDeltas === "function") {
        (documentController as any).applyExternalSheetViewDeltas(deltas, { source: "collab" });
      }
    } finally {
      applyingRemote = false;
    }
  };

  const hydrateFromYjs = () => {
    const ids: string[] = [];
    const arr = session.sheets?.toArray?.() ?? [];
    for (const entry of arr) {
      const id = getSheetIdFromSheetMap(entry as any);
      if (id) ids.push(id);
    }
    applyYjsToDocumentController(ids);
  };

  const handleDocumentChange = (payload: any) => {
    if (destroyed) return;
    if (applyingRemote) return;

    // In read-only collab sessions (viewer/commenter), avoid persisting local view metadata
    // back into the shared Yjs document.
    if (session.isReadOnly()) return;

    const sheetViewDeltas: SheetViewDelta[] = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
    if (sheetViewDeltas.length === 0) return;

    session.doc?.transact?.(
      () => {
        for (const delta of sheetViewDeltas) {
          const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
          if (!sheetId) continue;

          const sheet = findSheetMap(sheetId);
          if (!sheet) continue;

          const after = delta?.after as SheetViewState | undefined;
          const before = delta?.before as SheetViewState | undefined;
          if (!after) continue;

          // Mirror top-level frozen rows/cols for backwards compatibility.
          if (sheet.get("frozenRows") !== after.frozenRows) sheet.set("frozenRows", after.frozenRows);
          if (sheet.get("frozenCols") !== after.frozenCols) sheet.set("frozenCols", after.frozenCols);

          const viewMap = ensureNestedYMap(sheet, "view");

          if (viewMap.get("frozenRows") !== after.frozenRows) viewMap.set("frozenRows", after.frozenRows);
          if (viewMap.get("frozenCols") !== after.frozenCols) viewMap.set("frozenCols", after.frozenCols);

          const colWidthsMap = ensureNestedYMap(viewMap, "colWidths");
          const rowHeightsMap = ensureNestedYMap(viewMap, "rowHeights");

          applyAxisDelta(colWidthsMap, before?.colWidths, after.colWidths);
          applyAxisDelta(rowHeightsMap, before?.rowHeights, after.rowHeights);
        }
      },
      binderOrigin,
    );
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  const handleSheetsDeepChange = (events: any[], transaction: Y.Transaction) => {
    if (destroyed) return;
    if (!events || events.length === 0) return;

    const origin = transaction?.origin ?? null;
    if (origin === binderOrigin) return;

    let shouldHydrateAll = false;
    const changedSheetIds = new Set<string>();

    for (const event of events) {
      // If the root sheets array itself changed, hydrate all sheets.
      if (event?.target === session.sheets) {
        // Sheet added/removed/reordered.
        shouldHydrateAll = true;
        continue;
      }

      const path = event?.path;
      if (!Array.isArray(path) || path.length === 0) continue;
      const idx = path[0];
      if (typeof idx !== "number") continue;

      // Filter non-view mutations (e.g. renaming a sheet) to avoid needless work.
      let relevant = false;
      if (path.length === 1) {
        const keys = event?.changes?.keys;
        if (keys && typeof keys.has === "function") {
          for (const key of VIEW_KEYS) {
            if (keys.has(key)) {
              relevant = true;
              break;
            }
          }
        }
      } else {
        for (let i = 1; i < path.length; i += 1) {
          if (typeof path[i] === "string" && VIEW_KEYS.has(path[i])) {
            relevant = true;
            break;
          }
        }
      }
      if (!relevant) continue;

      const sheet = session.sheets.get(idx) as any;
      const sheetId = getSheetIdFromSheetMap(sheet);
      if (sheetId) changedSheetIds.add(sheetId);
    }

    if (shouldHydrateAll) {
      const arr = session.sheets?.toArray?.() ?? [];
      for (const entry of arr) {
        const id = getSheetIdFromSheetMap(entry as any);
        if (id) changedSheetIds.add(id);
      }
    }

    if (changedSheetIds.size === 0) return;
    applyYjsToDocumentController(changedSheetIds);
  };

  session.sheets.observeDeep(handleSheetsDeepChange);

  // Initial hydration (and for cases where the provider has already applied state).
  hydrateFromYjs();

  return {
    destroy() {
      if (destroyed) return;
      destroyed = true;
      unsubscribe?.();
      session.sheets.unobserveDeep(handleSheetsDeepChange);
      session.localOrigins?.delete?.(binderOrigin);
    },
  };
}
