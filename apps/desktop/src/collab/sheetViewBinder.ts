import * as Y from "yjs";
import { getYMap, getYText, yjsValueToJson } from "@formula/collab-yjs-utils";

import type { CollabSession } from "@formula/collab-session";
import type { DocumentController } from "../document/documentController.js";

export type SheetViewState = {
  frozenRows: number;
  frozenCols: number;
  backgroundImageId?: string;
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
  mergedRanges?: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>;
  drawings?: any[];
};

export type SheetViewDelta = {
  sheetId: string;
  before: SheetViewState;
  after: SheetViewState;
};

export type SheetViewBinder = { destroy: () => void };

const VIEW_KEYS = new Set([
  "view",
  "frozenRows",
  "frozenCols",
  "backgroundImageId",
  // Backwards compatibility with older clients / alternate naming.
  "background_image_id",
  "backgroundImage",
  "background_image",
  "colWidths",
  "rowHeights",
  "mergedRanges",
  "merged_ranges",
  "mergedRegions",
  "merged_regions",
  // Backwards compatibility with older clients.
  "mergedCells",
  "merged_cells",
  "drawings",
]);

function isRecord(value: unknown): value is Record<string, any> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isPlainObject(value: unknown): value is Record<string, any> {
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
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
  if (text) return String(yjsValueToJson(text));
  if (value == null) return null;
  return String(value);
}

function normalizeFrozenCount(value: unknown): number {
  const num = Number(yjsValueToJson(value));
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

function normalizeAxisSize(value: unknown): number | null {
  const num = Number(yjsValueToJson(value));
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

  // Support Y.Array (or duck-typed equivalents) by converting to a JS array.
  const maybeArray = (raw as any)?.toArray;
  if (!Array.isArray(raw) && typeof maybeArray === "function") {
    try {
      raw = maybeArray.call(raw);
    } catch {
      // ignore
    }
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

function cloneJsonValue<T>(value: T): T {
  if (value == null) return value;

  const structuredCloneImpl = (globalThis as any)?.structuredClone;
  if (typeof structuredCloneImpl === "function") {
    try {
      return structuredCloneImpl(value);
    } catch {
      // Fall through to JSON-based cloning.
    }
  }

  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    // Best-effort fallback: shallow clone common container types.
    if (Array.isArray(value)) return value.slice() as any;
    if (isRecord(value)) return { ...(value as any) };
    return value;
  }
}

function normalizeDrawings(raw: unknown): any[] | undefined {
  if (raw == null) return undefined;

  // Drawings are user-authored shared state (collab sheet view), so keep normalization defensive:
  // avoid deep-cloning entries with pathological ids (e.g. multi-megabyte strings).
  const MAX_DRAWING_ID_STRING_CHARS = 4096;

  const pushIfValid = (entry: unknown, out: any[]) => {
    if (!entry || typeof entry !== "object") return;
    // Accept only string or safe-integer ids (mirrors DocumentController validation).
    const rawId = (entry as any)?.get?.("id") ?? (entry as any).id;
    if (typeof rawId === "string") {
      if (rawId.length > MAX_DRAWING_ID_STRING_CHARS) return;
      if (!rawId.trim()) return;
    } else if (typeof rawId === "number") {
      if (!Number.isSafeInteger(rawId)) return;
    } else {
      return;
    }

    // Convert Yjs types to JSON before cloning (best-effort).
    let json: any = entry;
    if (json && typeof json === "object" && typeof json.toJSON === "function") {
      try {
        json = json.toJSON();
      } catch {
        // Ignore JSON conversion errors and fall back to the raw entry.
      }
    }
    if (!isRecord(json)) return;

    try {
      out.push(cloneJsonValue(json));
    } catch {
      // ignore malformed/non-serializable drawing entries
    }
  };

  // Support:
  // - plain JS arrays (common)
  // - Y.Array-like values (legacy/experimental schemas)
  /** @type {any[]} */
  const normalized: any[] = [];
  if (Array.isArray(raw)) {
    for (const entry of raw) pushIfValid(entry, normalized);
  } else {
    const maybeYArray = raw as any;
    const length = typeof maybeYArray?.length === "number" ? maybeYArray.length : 0;
    const get = typeof maybeYArray?.get === "function" ? maybeYArray.get.bind(maybeYArray) : null;
    if (length > 0 && get) {
      for (let i = 0; i < length; i += 1) {
        pushIfValid(get(i), normalized);
      }
    } else if ((raw as any) && typeof (raw as any).toJSON === "function") {
      // Last resort: materialize via toJSON and filter.
      try {
        const json = (raw as any).toJSON();
        if (Array.isArray(json)) {
          for (const entry of json) pushIfValid(entry, normalized);
        }
      } catch {
        // ignore
      }
    }
  }
  if (normalized.length === 0) return undefined;

  return normalized;
}

function deepEquals(a: any, b: any): boolean {
  if (a === b) return true;
  if (a == null || b == null) return a === b;
  if (typeof a !== typeof b) return false;

  if (typeof a !== "object") return false;

  if (Array.isArray(a)) {
    if (!Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (!deepEquals(a[i], b[i])) return false;
    }
    return true;
  }

  if (isRecord(a)) {
    if (!isRecord(b)) return false;
    const aKeys = Object.keys(a);
    const bKeys = Object.keys(b);
    if (aKeys.length !== bKeys.length) return false;
    aKeys.sort();
    bKeys.sort();
    for (let i = 0; i < aKeys.length; i += 1) {
      const key = aKeys[i]!;
      if (key !== bKeys[i]) return false;
      if (!deepEquals(a[key], b[key])) return false;
    }
    return true;
  }

  try {
    return JSON.stringify(a) === JSON.stringify(b);
  } catch {
    return false;
  }
}

function mergedRangesEquals(
  left?: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>,
  right?: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>,
): boolean {
  if (left === right) return true;
  const aRanges = Array.isArray(left) ? left : [];
  const bRanges = Array.isArray(right) ? right : [];
  if (aRanges.length !== bRanges.length) return false;
  if (aRanges.length === 0) return true;
  const keys = new Set(aRanges.map((r) => `${r.startRow},${r.endRow},${r.startCol},${r.endCol}`));
  if (keys.size !== aRanges.length) return false;
  for (const r of bRanges) {
    const key = `${r.startRow},${r.endRow},${r.startCol},${r.endCol}`;
    if (!keys.has(key)) return false;
  }
  return true;
}

function normalizeOptionalId(value: unknown): string | null {
  const raw = coerceString(value);
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim();
  return trimmed ? trimmed : null;
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
    normalizeOptionalId(a.backgroundImageId) === normalizeOptionalId(b.backgroundImageId) &&
    axisEquals(a.colWidths, b.colWidths) &&
    axisEquals(a.rowHeights, b.rowHeights) &&
    mergedRangesEquals(a.mergedRanges, b.mergedRanges) &&
    deepEquals(
      Array.isArray(a.drawings) && a.drawings.length > 0 ? a.drawings : null,
      Array.isArray(b.drawings) && b.drawings.length > 0 ? b.drawings : null,
    )
  );
}

function readMergedRanges(raw: unknown): Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> | undefined {
  if (!raw) return undefined;
  // Support Y.Array (or duck-typed equivalents) by converting to a JS array.
  const maybeArray = (raw as any)?.toArray;
  if (!Array.isArray(raw) && typeof maybeArray === "function") {
    try {
      raw = maybeArray.call(raw);
    } catch {
      // ignore
    }
  }
  if (!Array.isArray(raw) || raw.length === 0) return undefined;

  const overlaps = (a: any, b: any) =>
    a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;

  const readEntryField = (entry: any, keys: string[]): any => {
    if (!entry || typeof entry !== "object") return undefined;
    const get = typeof entry.get === "function" ? entry.get.bind(entry) : null;
    if (get) {
      for (const key of keys) {
        try {
          const value = get(key);
          if (value !== undefined) return value;
        } catch {
          // ignore
        }
      }
    }
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(entry, key)) return entry[key];
    }
    return undefined;
  };

  const readNonNegInt = (entry: any, keys: string[]): number | null => {
    const rawValue = readEntryField(entry, keys);
    const num = Number(yjsValueToJson(rawValue));
    if (!Number.isInteger(num) || num < 0) return null;
    return num;
  };

  const out: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];
  for (const entry of raw) {
    const range = readEntryField(entry, ["range"]) ?? entry;

    let sr = readNonNegInt(range, ["startRow", "start_row", "sr"]);
    let er = readNonNegInt(range, ["endRow", "end_row", "er"]);
    let sc = readNonNegInt(range, ["startCol", "start_col", "sc"]);
    let ec = readNonNegInt(range, ["endCol", "end_col", "ec"]);

    if (sr == null || er == null || sc == null || ec == null) {
      const start = readEntryField(range, ["start"]);
      const end = readEntryField(range, ["end"]);
      if (sr == null) sr = readNonNegInt(start, ["row", "r"]);
      if (sc == null) sc = readNonNegInt(start, ["col", "c"]);
      if (er == null) er = readNonNegInt(end, ["row", "r"]);
      if (ec == null) ec = readNonNegInt(end, ["col", "c"]);
    }

    if (sr == null || er == null || sc == null || ec == null) continue;

    const startRow = Math.min(sr, er);
    const endRow = Math.max(sr, er);
    const startCol = Math.min(sc, ec);
    const endCol = Math.max(sc, ec);
    if (startRow === endRow && startCol === endCol) continue;

    const candidate = { startRow, endRow, startCol, endCol };
    for (let i = out.length - 1; i >= 0; i--) {
      if (overlaps(out[i], candidate)) out.splice(i, 1);
    }
    out.push(candidate);
  }

  if (out.length === 0) return undefined;
  out.sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol || a.endRow - b.endRow || a.endCol - b.endCol);

  // Deduplicate after sorting.
  const deduped: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];
  let lastKey: string | null = null;
  for (const r of out) {
    const key = `${r.startRow},${r.endRow},${r.startCol},${r.endCol}`;
    if (key === lastKey) continue;
    lastKey = key;
    deduped.push(r);
  }

  return deduped.length > 0 ? deduped : undefined;
}

function getSheetIdFromSheetMap(sheet: any): string | null {
  return coerceString(sheet?.get?.("id") ?? sheet?.id);
}

function readSheetViewFromSheetMap(sheet: any): SheetViewState {
  const viewRaw = sheet?.get?.("view") ?? sheet?.view;

  let frozenRowsRaw: unknown = undefined;
  if (viewRaw !== undefined) frozenRowsRaw = readYMapOrObject(viewRaw, "frozenRows");
  if (frozenRowsRaw === undefined) frozenRowsRaw = sheet?.get?.("frozenRows") ?? sheet?.frozenRows;
  const frozenRows = normalizeFrozenCount(frozenRowsRaw);

  let frozenColsRaw: unknown = undefined;
  if (viewRaw !== undefined) frozenColsRaw = readYMapOrObject(viewRaw, "frozenCols");
  if (frozenColsRaw === undefined) frozenColsRaw = sheet?.get?.("frozenCols") ?? sheet?.frozenCols;
  const frozenCols = normalizeFrozenCount(frozenColsRaw);

  let colWidthsRaw: unknown = undefined;
  if (viewRaw !== undefined) colWidthsRaw = readYMapOrObject(viewRaw, "colWidths");
  if (colWidthsRaw === undefined) colWidthsRaw = sheet?.get?.("colWidths") ?? sheet?.colWidths;
  const colWidths = readAxisOverrides(colWidthsRaw);

  let rowHeightsRaw: unknown = undefined;
  if (viewRaw !== undefined) rowHeightsRaw = readYMapOrObject(viewRaw, "rowHeights");
  if (rowHeightsRaw === undefined) rowHeightsRaw = sheet?.get?.("rowHeights") ?? sheet?.rowHeights;
  const rowHeights = readAxisOverrides(rowHeightsRaw);
  let mergedRangesRaw: unknown = undefined;
  if (viewRaw !== undefined) {
    mergedRangesRaw =
      readYMapOrObject(viewRaw, "mergedRanges") ??
      readYMapOrObject(viewRaw, "merged_ranges") ??
      readYMapOrObject(viewRaw, "mergedRegions") ??
      readYMapOrObject(viewRaw, "merged_regions") ??
      readYMapOrObject(viewRaw, "mergedCells") ??
      readYMapOrObject(viewRaw, "merged_cells");
  }
  if (mergedRangesRaw === undefined) {
    mergedRangesRaw =
      sheet?.get?.("mergedRanges") ??
      sheet?.get?.("merged_ranges") ??
      sheet?.get?.("mergedRegions") ??
      sheet?.get?.("merged_regions") ??
      sheet?.get?.("mergedCells") ??
      sheet?.get?.("merged_cells") ??
      sheet?.mergedRanges ??
      sheet?.merged_ranges ??
      sheet?.mergedRegions ??
      sheet?.merged_regions ??
      sheet?.mergedCells ??
      sheet?.merged_cells;
  }
  const mergedRanges = readMergedRanges(mergedRangesRaw);

  let backgroundRaw: unknown = undefined;
  if (viewRaw !== undefined) {
    backgroundRaw =
      readYMapOrObject(viewRaw, "backgroundImageId") ??
      readYMapOrObject(viewRaw, "background_image_id") ??
      readYMapOrObject(viewRaw, "backgroundImage") ??
      readYMapOrObject(viewRaw, "background_image");
  }
  if (backgroundRaw === undefined) {
    backgroundRaw =
      sheet?.get?.("backgroundImageId") ??
      sheet?.get?.("background_image_id") ??
      sheet?.get?.("backgroundImage") ??
      sheet?.get?.("background_image") ??
      sheet?.backgroundImageId ??
      (sheet as any)?.background_image_id ??
      sheet?.backgroundImage ??
      (sheet as any)?.background_image;
  }
  const backgroundImageId = normalizeOptionalId(backgroundRaw) ?? undefined;

  let drawingsRaw: unknown = undefined;
  if (viewRaw !== undefined) drawingsRaw = readYMapOrObject(viewRaw, "drawings");
  if (drawingsRaw === undefined) drawingsRaw = sheet?.get?.("drawings") ?? sheet?.drawings;
  const drawings = normalizeDrawings(drawingsRaw);

  const out: SheetViewState = { frozenRows, frozenCols };
  if (backgroundImageId) out.backgroundImageId = backgroundImageId;
  if (colWidths) out.colWidths = colWidths;
  if (rowHeights) out.rowHeights = rowHeights;
  if (mergedRanges) out.mergedRanges = mergedRanges;
  if (drawings) out.drawings = drawings;
  return out;
}

function ensureNestedYMap(parent: any, key: string): Y.Map<any> {
  const existing = parent?.get?.(key);
  const existingMap = getYMap(existing);
  if (existingMap) return existingMap;

  // Prefer constructing nested Yjs types using the parent's constructor to tolerate
  // mixed-module environments (ESM + CJS) where multiple `yjs` instances exist.
  const ParentCtor = (parent as any)?.constructor as { new (): any } | undefined;
  const next = (() => {
    if (typeof ParentCtor === "function" && ParentCtor !== Object) {
      try {
        return new ParentCtor();
      } catch {
        // Fall through.
      }
    }
    return new Y.Map();
  })();

  // Best-effort: if the existing value was a plain object, preserve entries.
  if (isPlainObject(existing)) {
    for (const [k, v] of Object.entries(existing)) {
      next.set(k, v);
    }
  }

  parent?.set?.(key, next);
  return next;
}

function ensureAxisOverridesYMap(parent: any, key: string): { map: Y.Map<any>; created: boolean } {
  const existing = parent?.get?.(key);
  const existingMap = getYMap(existing);
  if (existingMap) return { map: existingMap, created: false };

  // Prefer constructing nested Yjs types using the parent's constructor to tolerate
  // mixed-module environments (ESM + CJS) where multiple `yjs` instances exist.
  const ParentCtor = (parent as any)?.constructor as { new (): any } | undefined;
  const next = (() => {
    if (typeof ParentCtor === "function" && ParentCtor !== Object) {
      try {
        return new ParentCtor();
      } catch {
        // Fall through.
      }
    }
    return new Y.Map();
  })();

  parent?.set?.(key, next);
  return { map: next, created: true };
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

  const ownsOrigin = options?.origin == null;
  const binderOrigin = options?.origin ?? { type: "document-controller:sheet-view-binder" };
  session.localOrigins?.add?.(binderOrigin);

  let destroyed = false;
  let applyingRemote = false;

  const findYjsSheetEntriesById = (sheetId: string): any[] => {
    if (!sheetId) return [];
    /** @type {any[]} */
    const out: any[] = [];
    const len = typeof (session.sheets as any)?.length === "number" ? (session.sheets as any).length : 0;
    for (let i = 0; i < len; i += 1) {
      const entry = (session.sheets as any).get(i);
      const id = getSheetIdFromSheetMap(entry);
      if (id === sheetId) out.push(entry);
    }
    return out;
  };

  const findYjsSheetEntryById = (sheetId: string): any | null => {
    if (!sheetId) return null;
    // Deterministic choice: pick the last matching entry by array index. This mirrors
    // `ensureWorkbookSchema` duplicate-sheet pruning behavior.
    let found: any | null = null;
    const len = typeof (session.sheets as any)?.length === "number" ? (session.sheets as any).length : 0;
    for (let i = 0; i < len; i += 1) {
      const entry = (session.sheets as any).get(i);
      const id = getSheetIdFromSheetMap(entry);
      if (id === sheetId) found = entry;
    }
    return found;
  };

  const applyYjsToDocumentController = (sheetIds: Iterable<string>): void => {
    const viewDeltas: SheetViewDelta[] = [];
    /** @type {Array<{ sheetId: string, before: any[], after: any[] }>} */
    const drawingDeltas: Array<{ sheetId: string; before: any[]; after: any[] }> = [];

    const canApplyExternalDrawings =
      typeof (documentController as any).getSheetDrawings === "function" &&
      typeof (documentController as any).applyExternalDrawingDeltas === "function";

    for (const sheetId of sheetIds) {
      const sheet = findYjsSheetEntryById(sheetId);
      if (!sheet) continue;

      const after = readSheetViewFromSheetMap(sheet);

      // Sheet view (frozen panes + axis overrides) lives in `DocumentController.getSheetView`.
      const beforeRaw = documentController.getSheetView(sheetId) as any;

      // Compare only the keys this binder owns so we don't churn on unrelated view metadata.
      const beforeComparable: SheetViewState = {
        frozenRows: beforeRaw.frozenRows,
        frozenCols: beforeRaw.frozenCols,
        ...(normalizeOptionalId(beforeRaw.backgroundImageId) ? { backgroundImageId: beforeRaw.backgroundImageId } : {}),
        ...(beforeRaw.colWidths ? { colWidths: beforeRaw.colWidths } : {}),
        ...(beforeRaw.rowHeights ? { rowHeights: beforeRaw.rowHeights } : {}),
        ...(Array.isArray(beforeRaw.mergedRanges) && beforeRaw.mergedRanges.length > 0
          ? { mergedRanges: beforeRaw.mergedRanges }
          : {}),
        // If the controller stores drawings inside sheet view state (instead of `drawingDeltas`),
        // include them in the equality check so remote drawing updates hydrate correctly.
        ...(canApplyExternalDrawings
          ? {}
          : Array.isArray(beforeRaw.drawings) && beforeRaw.drawings.length > 0
            ? { drawings: beforeRaw.drawings }
            : {}),
      };

      const afterComparable: SheetViewState = {
        frozenRows: after.frozenRows,
        frozenCols: after.frozenCols,
        ...(after.backgroundImageId ? { backgroundImageId: after.backgroundImageId } : {}),
        ...(after.colWidths ? { colWidths: after.colWidths } : {}),
        ...(after.rowHeights ? { rowHeights: after.rowHeights } : {}),
        ...(Array.isArray(after.mergedRanges) && after.mergedRanges.length > 0 ? { mergedRanges: after.mergedRanges } : {}),
        ...(canApplyExternalDrawings ? {} : after.drawings ? { drawings: after.drawings } : {}),
      };

      if (!sheetViewEquals(beforeComparable, afterComparable)) {
        // Preserve unrelated view keys from the DocumentController (e.g. mergedRanges) so remote
        // frozen/axis updates don't accidentally wipe local metadata that this binder does not sync.
        const afterFull: any = { ...beforeRaw };
        afterFull.frozenRows = after.frozenRows;
        afterFull.frozenCols = after.frozenCols;
        if (after.backgroundImageId) afterFull.backgroundImageId = after.backgroundImageId;
        else delete afterFull.backgroundImageId;
        if (after.colWidths) afterFull.colWidths = after.colWidths;
        else delete afterFull.colWidths;
        if (after.rowHeights) afterFull.rowHeights = after.rowHeights;
        else delete afterFull.rowHeights;
        if (Array.isArray(after.mergedRanges) && after.mergedRanges.length > 0) afterFull.mergedRanges = after.mergedRanges;
        else delete afterFull.mergedRanges;

        if (!canApplyExternalDrawings) {
          if (after.drawings) afterFull.drawings = after.drawings;
          else delete afterFull.drawings;
        }

        viewDeltas.push({ sheetId, before: beforeRaw, after: afterFull });
      }

      // Drawings are stored separately in the DocumentController (Task 218) as `drawingDeltas`.
      if (canApplyExternalDrawings) {
        const beforeDrawingsRaw = (documentController as any).getSheetDrawings(sheetId);
        const afterDrawingsRaw = after.drawings ?? [];

        const beforeDrawings = Array.isArray(beforeDrawingsRaw) ? beforeDrawingsRaw : [];
        const afterDrawings = Array.isArray(afterDrawingsRaw) ? afterDrawingsRaw : [];

        const beforeComparable = beforeDrawings.length > 0 ? beforeDrawings : null;
        const afterComparable = afterDrawings.length > 0 ? afterDrawings : null;
        if (!deepEquals(beforeComparable, afterComparable)) {
          drawingDeltas.push({ sheetId, before: beforeDrawings, after: afterDrawings });
        }
      }
    }

    if (viewDeltas.length === 0 && drawingDeltas.length === 0) return;

    applyingRemote = true;
    try {
      if (viewDeltas.length > 0 && typeof (documentController as any).applyExternalSheetViewDeltas === "function") {
        (documentController as any).applyExternalSheetViewDeltas(viewDeltas, { source: "collab" });
      }

      if (drawingDeltas.length > 0 && canApplyExternalDrawings) {
        (documentController as any).applyExternalDrawingDeltas(drawingDeltas, { source: "collab" });
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

    // The desktop app wires both the full collab binder (cell/format/etc) and this sheet-view binder.
    // When the full binder applies remote deltas into the DocumentController, it emits `change`
    // events with `payload.source === "collab"`. Treat those as remote and do not write them back
    // into Yjs (otherwise we'd create redundant Yjs updates and potentially pollute collaborative undo).
    //
    // Similarly, `applyState` is used for version restore / snapshot hydration and should not
    // automatically overwrite the shared collaborative state.
    const source = typeof payload?.source === "string" ? payload.source : null;
    if (source === "collab" || source === "applyState") return;

    // In read-only collab sessions (viewer/commenter), avoid persisting local view metadata
    // back into the shared Yjs document.
    if (session.isReadOnly()) return;

    const sheetViewDeltas: SheetViewDelta[] = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
    const drawingDeltas: Array<{ sheetId: string; before: any[]; after: any[] }> = Array.isArray(payload?.drawingDeltas)
      ? payload.drawingDeltas
      : [];

    if (sheetViewDeltas.length === 0 && drawingDeltas.length === 0) return;

    session.doc?.transact?.(
      () => {
        for (const delta of sheetViewDeltas) {
          const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
          if (!sheetId) continue;

          const sheetsToUpdate = findYjsSheetEntriesById(sheetId);
          if (sheetsToUpdate.length === 0) continue;

          const after = delta?.after as SheetViewState | undefined;
          const before = delta?.before as SheetViewState | undefined;
          if (!after) continue;

          const nextBackgroundImageId = normalizeOptionalId((after as any).backgroundImageId);
          const prevBackgroundImageId = normalizeOptionalId((before as any)?.backgroundImageId);

          const nextMergedRanges =
            Array.isArray(after.mergedRanges) && after.mergedRanges.length > 0 ? after.mergedRanges : undefined;
          const prevMergedRanges =
            Array.isArray(before?.mergedRanges) && before.mergedRanges.length > 0 ? before.mergedRanges : undefined;

           const drawingsBefore = Array.isArray(before?.drawings) && before.drawings.length > 0 ? before.drawings : null;
           const drawingsAfter = Array.isArray(after.drawings) && after.drawings.length > 0 ? after.drawings : null;

           for (const sheet of sheetsToUpdate) {
             // Some historical docs/tests may store sheet entries as plain objects in the Y.Array
             // rather than Y.Maps. Hydration can still read from those, but we cannot write back.
             if (!sheet || typeof sheet.get !== "function" || typeof sheet.set !== "function" || typeof sheet.delete !== "function") {
               continue;
             }
             // Mirror top-level frozen rows/cols for backwards compatibility.
             if (sheet.get("frozenRows") !== after.frozenRows) sheet.set("frozenRows", after.frozenRows);
             if (sheet.get("frozenCols") !== after.frozenCols) sheet.set("frozenCols", after.frozenCols);

            const viewMap = ensureNestedYMap(sheet, "view");

            if (viewMap.get("frozenRows") !== after.frozenRows) viewMap.set("frozenRows", after.frozenRows);
            if (viewMap.get("frozenCols") !== after.frozenCols) viewMap.set("frozenCols", after.frozenCols);

            if (nextBackgroundImageId !== prevBackgroundImageId) {
              if (nextBackgroundImageId) {
                viewMap.set("backgroundImageId", nextBackgroundImageId);
                // Back-compat mirror (similar to other sheet view keys).
                sheet.set("backgroundImageId", nextBackgroundImageId);
                // Converge legacy key names to the canonical one.
                viewMap.delete("background_image_id");
                sheet.delete("background_image_id");
                viewMap.delete("backgroundImage");
                sheet.delete("backgroundImage");
                viewMap.delete("background_image");
                sheet.delete("background_image");
              } else {
                viewMap.delete("backgroundImageId");
                sheet.delete("backgroundImageId");
                viewMap.delete("background_image_id");
                sheet.delete("background_image_id");
                viewMap.delete("backgroundImage");
                sheet.delete("backgroundImage");
                viewMap.delete("background_image");
                sheet.delete("background_image");
              }
            }

            const beforeColWidths = before?.colWidths;
            const afterColWidths = after.colWidths;
            if (beforeColWidths || afterColWidths) {
              if (afterColWidths) {
                const { map: colWidthsMap, created } = ensureAxisOverridesYMap(viewMap, "colWidths");
                applyAxisDelta(colWidthsMap, created ? undefined : beforeColWidths, afterColWidths);
              } else {
                viewMap.delete("colWidths");
              }
              // Converge legacy top-level encoding to the nested `view.colWidths` map.
              sheet.delete("colWidths");
            }

            const beforeRowHeights = before?.rowHeights;
            const afterRowHeights = after.rowHeights;
            if (beforeRowHeights || afterRowHeights) {
              if (afterRowHeights) {
                const { map: rowHeightsMap, created } = ensureAxisOverridesYMap(viewMap, "rowHeights");
                applyAxisDelta(rowHeightsMap, created ? undefined : beforeRowHeights, afterRowHeights);
              } else {
                viewMap.delete("rowHeights");
              }
              // Converge legacy top-level encoding to the nested `view.rowHeights` map.
              sheet.delete("rowHeights");
            }

            if (!nextMergedRanges) {
              if (prevMergedRanges) {
                viewMap.delete("mergedRanges");
                sheet.delete("mergedRanges");
                viewMap.delete("merged_ranges");
                sheet.delete("merged_ranges");
                // Backwards compatibility cleanup.
                viewMap.delete("mergedCells");
                sheet.delete("mergedCells");
                viewMap.delete("merged_cells");
                sheet.delete("merged_cells");
                viewMap.delete("mergedRegions");
                sheet.delete("mergedRegions");
                viewMap.delete("merged_regions");
                sheet.delete("merged_regions");
              }
            } else if (!mergedRangesEquals(prevMergedRanges, nextMergedRanges)) {
              const cloned = nextMergedRanges.map((r) => ({
                startRow: r.startRow,
                endRow: r.endRow,
                startCol: r.startCol,
                endCol: r.endCol,
              }));
              // Store on both the nested view map (preferred) and the top-level for backwards compatibility.
              // Keep this minimal (avoid writing the same payload under multiple alias keys) to prevent
              // workbook bloat when many merged ranges exist.
              viewMap.set("mergedRanges", cloned);
              sheet.set("mergedRanges", cloned);
              // Back-compat mirror: older clients used `mergedCells`.
              viewMap.set("mergedCells", cloned);
              sheet.set("mergedCells", cloned);

              // Converge any legacy/alternate key names to the canonical + legacy pair above.
              viewMap.delete("merged_ranges");
              sheet.delete("merged_ranges");
              viewMap.delete("mergedRegions");
              sheet.delete("mergedRegions");
              viewMap.delete("merged_regions");
              sheet.delete("merged_regions");
              viewMap.delete("merged_cells");
              sheet.delete("merged_cells");
            }

            if (!deepEquals(drawingsBefore, drawingsAfter)) {
              if (drawingsAfter) {
                const cloned = cloneJsonValue(drawingsAfter);
                viewMap.set("drawings", cloned);
                // Back-compat: mirror to the sheet root so older clients can still observe updates.
                sheet.set("drawings", cloned);
              } else {
                viewMap.delete("drawings");
                sheet.delete("drawings");
              }
            }
          }
        }

        for (const delta of drawingDeltas) {
          const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
          if (!sheetId) continue;

          const sheetsToUpdate = findYjsSheetEntriesById(sheetId);
          if (sheetsToUpdate.length === 0) continue;

          const beforeDrawings = Array.isArray(delta?.before) && delta.before.length > 0 ? delta.before : null;
          const afterDrawings = Array.isArray(delta?.after) && delta.after.length > 0 ? delta.after : null;

          if (deepEquals(beforeDrawings, afterDrawings)) continue;

          for (const sheet of sheetsToUpdate) {
            // Some historical docs/tests may store sheet entries as plain objects in the Y.Array
            // rather than Y.Maps. Hydration can still read from those, but we cannot write back.
            if (!sheet || typeof sheet.get !== "function" || typeof sheet.set !== "function" || typeof sheet.delete !== "function") {
              continue;
            }
            const viewMap = ensureNestedYMap(sheet, "view");
            if (afterDrawings) {
              const cloned = cloneJsonValue(afterDrawings);
              viewMap.set("drawings", cloned);
              // Back-compat: mirror to the sheet root so older clients can still observe updates.
              sheet.set("drawings", cloned);
            } else {
              viewMap.delete("drawings");
              sheet.delete("drawings");
            }
          }
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
      if (ownsOrigin) {
        session.localOrigins?.delete?.(binderOrigin);
      }
    },
  };
}
