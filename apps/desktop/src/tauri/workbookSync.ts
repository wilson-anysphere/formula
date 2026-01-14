import { normalizeFormulaTextOpt } from "@formula/engine";
import { showToast } from "../extensions/ui.js";
import { parseImageCellValue } from "../shared/imageCellValue.js";

import { getTauriInvokeOrNull, type TauriInvoke } from "./api";

type SheetVisibility = "visible" | "hidden" | "veryHidden";

type TabColor = {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
};

type SheetMetaState = {
  name: string;
  visibility: SheetVisibility;
  tabColor?: TabColor;
};

type SheetMetaDelta = {
  sheetId: string;
  before: SheetMetaState | null;
  after: SheetMetaState | null;
};

type SheetOrderDelta = {
  before: string[];
  after: string[];
};

type CellState = {
  value: unknown;
  formula: string | null;
  // DocumentController tracks styleId too; workbook sync mirrors formatting changes
  // through a dedicated formatting IPC channel (not via set_cell/set_range).
  styleId?: number;
};

type CellDelta = {
  sheetId: string;
  row: number;
  col: number;
  before: CellState;
  after: CellState;
};

type SheetViewState = {
  frozenRows?: number;
  frozenCols?: number;
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
};

type SheetViewDelta = {
  sheetId: string;
  before: SheetViewState;
  after: SheetViewState;
};

type DocumentControllerLike = {
  on(
    event: "change",
    listener: (payload: {
      deltas: CellDelta[];
      sheetViewDeltas?: SheetViewDelta[];
      rowStyleDeltas?: Array<{ sheetId: string; row: number; beforeStyleId: number; afterStyleId: number }>;
      colStyleDeltas?: Array<{ sheetId: string; col: number; beforeStyleId: number; afterStyleId: number }>;
      sheetStyleDeltas?: Array<{ sheetId: string; beforeStyleId: number; afterStyleId: number }>;
      rangeRunDeltas?: Array<{
        sheetId: string;
        col: number;
        startRow: number;
        endRowExclusive: number;
        beforeRuns: Array<{ startRow: number; endRowExclusive: number; styleId: number }>;
        afterRuns: Array<{ startRow: number; endRowExclusive: number; styleId: number }>;
      }>;
      sheetMetaDeltas?: SheetMetaDelta[];
      sheetOrderDelta?: SheetOrderDelta | null;
      source?: string;
      recalc?: boolean;
    }) => void,
  ): () => void;
  markSaved(): void;
  readonly isDirty: boolean;
  getSheetIds?(): string[];
  getSheetMeta?(sheetId: string): SheetMetaState | null;
  readonly styleTable?: { get(styleId: number): any };
  // Optional APIs on the real DocumentController used to apply authoritative backend updates
  // (e.g. pivot auto-refresh output).
  getCell?(sheetId: string, coord: { row: number; col: number }): any;
  applyExternalDeltas?(deltas: any[], options?: { source?: string; markDirty?: boolean }): void;
};

type RangeCellEdit = { value: unknown | null; formula: string | null };

function resolveInvoke(engineBridge: unknown): TauriInvoke | null {
  if (engineBridge && typeof engineBridge === "object") {
    const maybe = (engineBridge as any).invoke;
    if (typeof maybe === "function") {
      return maybe as TauriInvoke;
    }
  }
  return getTauriInvokeOrNull();
}

function valuesEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a !== "object" || typeof b !== "object") return false;
  try {
    return JSON.stringify(a) === JSON.stringify(b);
  } catch {
    return false;
  }
}

function inputEquals(before: { value: unknown; formula: string | null }, after: { value: unknown; formula: string | null }): boolean {
  return valuesEqual(before.value ?? null, after.value ?? null) && (before.formula ?? null) === (after.formula ?? null);
}

function tabColorToKey(input: unknown): string {
  if (!input) return "";
  if (typeof input === "string") return input.trim();
  if (typeof input !== "object") return "";
  const color = input as any;
  const parts: string[] = [];
  if (typeof color.rgb === "string") parts.push(`rgb:${color.rgb.trim()}`);
  if (typeof color.theme === "number") parts.push(`theme:${color.theme}`);
  if (typeof color.indexed === "number") parts.push(`indexed:${color.indexed}`);
  if (typeof color.tint === "number") parts.push(`tint:${color.tint}`);
  if (typeof color.auto === "boolean") parts.push(`auto:${color.auto ? 1 : 0}`);
  return parts.join("|");
}

function tabColorEquals(a: unknown, b: unknown): boolean {
  return tabColorToKey(a) === tabColorToKey(b);
}

function tabColorToBackendArg(input: unknown): TabColor | null {
  if (!input) return null;

  // Some legacy/internal flows represent tab color as a raw rgb string.
  if (typeof input === "string") {
    const rgb = input.trim();
    return rgb ? { rgb: rgb.toUpperCase() } : null;
  }

  if (typeof input !== "object") return null;
  const color = input as any;
  const out: TabColor = {};

  if (typeof color.rgb === "string") {
    const rgb = color.rgb.trim();
    if (rgb) out.rgb = rgb.toUpperCase();
  }
  if (typeof color.theme === "number" && Number.isFinite(color.theme)) out.theme = color.theme;
  if (typeof color.indexed === "number" && Number.isFinite(color.indexed)) out.indexed = color.indexed;
  if (typeof color.tint === "number" && Number.isFinite(color.tint)) out.tint = color.tint;
  if (typeof color.auto === "boolean") out.auto = color.auto;

  if (
    out.rgb == null &&
    out.theme == null &&
    out.indexed == null &&
    out.tint == null &&
    out.auto == null
  ) {
    return null;
  }
  return out;
}

function normalizeSheetVisibility(raw: unknown): SheetVisibility {
  return raw === "hidden" || raw === "veryHidden" || raw === "visible" ? raw : "visible";
}

function normalizeSheetMeta(raw: unknown, fallbackSheetId: string): SheetMetaState {
  const obj = raw && typeof raw === "object" ? (raw as any) : null;
  const name = String(obj?.name ?? fallbackSheetId ?? "").trim() || fallbackSheetId;
  const visibility = normalizeSheetVisibility(obj?.visibility);
  const tabColorRaw = obj?.tabColor;
  const tabColor = tabColorRaw && typeof tabColorRaw === "object" ? { ...(tabColorRaw as any) } : undefined;
  return tabColor ? { name, visibility, tabColor } : { name, visibility };
}

// NOTE: In desktop mode, the Tauri backend workbook is the persistence layer for both cell edits
// and sheet structure/metadata (rename/reorder/add/delete/hide/tabColor). Instead of relying on
// ad-hoc UI hooks to keep the backend in sync (which can drift from undo/redo/applyState), we
// mirror the authoritative DocumentController change stream to the backend here.

type PendingEdit = { sheetId: string; row: number; col: number; edit: RangeCellEdit };

type CellFormatDelta = { sheetId: string; row: number; col: number; beforeFormat: any | null; afterFormat: any | null };
type RowStyleDelta = { sheetId: string; row: number; beforeFormat: any | null; afterFormat: any | null };
type ColStyleDelta = { sheetId: string; col: number; beforeFormat: any | null; afterFormat: any | null };
type SheetStyleDelta = { sheetId: string; beforeFormat: any | null; afterFormat: any | null };
type RangeRun = { startRow: number; endRowExclusive: number; format: any | null };
type RangeRunDelta = {
  sheetId: string;
  col: number;
  startRow: number;
  endRowExclusive: number;
  beforeRuns: RangeRun[];
  afterRuns: RangeRun[];
};

function toRangeCellEdit(state: CellState): RangeCellEdit {
  if (state.formula != null) {
    const normalized = normalizeFormulaTextOpt(state.formula);
    if (normalized != null) {
      return { value: null, formula: normalized };
    }
  }

  const value = (state.value ?? null) as unknown | null;
  // Desktop-only: embedded in-cell images are represented as JSON-ish objects in the
  // DocumentController so the grid can resolve bytes by id. The Rust workbook backend does
  // not (yet) accept these objects via `set_cell`/`set_range`, so strip them out when
  // mirroring deltas to avoid IPC serialization failures.
  if (parseImageCellValue(value)) return { value: null, formula: null };
  return { value, formula: null };
}

function normalizeFormulaText(formula: unknown): string | null {
  if (typeof formula !== "string") return null;
  return normalizeFormulaTextOpt(formula);
}

function cellKey(sheetId: string, row: number, col: number): string {
  return `${sheetId}:${row},${col}`;
}

function rowStyleKey(sheetId: string, row: number): string {
  return `${sheetId}:row:${row}`;
}

function colStyleKey(sheetId: string, col: number): string {
  return `${sheetId}:col:${col}`;
}

function rangeRunKey(sheetId: string, col: number): string {
  return `${sheetId}:rangeRun:${col}`;
}

function colWidthKey(sheetId: string, col: number): string {
  return `${sheetId}:colWidth:${col}`;
}

function rowHeightKey(sheetId: string, row: number): string {
  return `${sheetId}:rowHeight:${row}`;
}

function styleIdToFormat(document: DocumentControllerLike, styleId: unknown): any | null {
  const id = typeof styleId === "number" ? styleId : 0;
  if (!Number.isInteger(id) || id === 0) return null;
  const table = document.styleTable;
  if (!table || typeof table.get !== "function") return null;
  return table.get(id) ?? null;
}

function sortPendingEdits(a: PendingEdit, b: PendingEdit): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  if (a.row !== b.row) return a.row - b.row;
  return a.col - b.col;
}

function sortCellFormatDeltas(a: CellFormatDelta, b: CellFormatDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  if (a.row !== b.row) return a.row - b.row;
  return a.col - b.col;
}

function sortRowStyleDeltas(a: RowStyleDelta, b: RowStyleDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  return a.row - b.row;
}

function sortColStyleDeltas(a: ColStyleDelta, b: ColStyleDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  return a.col - b.col;
}

function sortSheetStyleDeltas(a: SheetStyleDelta, b: SheetStyleDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  return 0;
}

function sortRangeRunDeltas(a: RangeRunDelta, b: RangeRunDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  if (a.col !== b.col) return a.col - b.col;
  return a.startRow - b.startRow;
}

type ColWidthDelta = { sheetId: string; col: number; width: number | null };
type RowHeightDelta = { sheetId: string; row: number; height: number | null };

function sortColWidthDeltas(a: ColWidthDelta, b: ColWidthDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  return a.col - b.col;
}

function sortRowHeightDeltas(a: RowHeightDelta, b: RowHeightDelta): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  return a.row - b.row;
}

function isFullRectangle(edits: PendingEdit[]): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  if (edits.length === 0) return null;
  const sheetId = edits[0]?.sheetId;
  if (!sheetId) return null;
  if (edits.some((e) => e.sheetId !== sheetId)) return null;

  let startRow = Number.POSITIVE_INFINITY;
  let startCol = Number.POSITIVE_INFINITY;
  let endRow = Number.NEGATIVE_INFINITY;
  let endCol = Number.NEGATIVE_INFINITY;

  const coords = new Set<string>();
  for (const e of edits) {
    startRow = Math.min(startRow, e.row);
    startCol = Math.min(startCol, e.col);
    endRow = Math.max(endRow, e.row);
    endCol = Math.max(endCol, e.col);
    coords.add(`${e.row},${e.col}`);
  }

  const expected = (endRow - startRow + 1) * (endCol - startCol + 1);
  if (coords.size !== expected) return null;
  return { startRow, startCol, endRow, endCol };
}

function applyBackendUpdates(document: DocumentControllerLike, raw: unknown, opts?: { skipSheetIds?: Set<string> }): void {
  if (typeof document.getCell !== "function" || typeof document.applyExternalDeltas !== "function") return;
  if (!Array.isArray(raw) || raw.length === 0) return;

  const skip = opts?.skipSheetIds ?? null;
  const deltas: any[] = [];
  for (const u of raw as any[]) {
    if (!u || typeof u !== "object") continue;
    const sheetId = String((u as any).sheet_id ?? "").trim();
    const row = Number((u as any).row);
    const col = Number((u as any).col);
    if (!sheetId) continue;
    if (skip && skip.has(sheetId)) continue;
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    // Backend returns computed value updates for formula cells; the frontend has its own calc engine.
    // We only apply input changes for non-formula cells (e.g. pivot output values).
    if (normalizeFormulaText((u as any).formula) != null) continue;

    const docAny: any = document as any;
    const before = typeof docAny.peekCell === "function" ? docAny.peekCell(sheetId, { row, col }) : document.getCell(sheetId, { row, col });
    const after = { value: (u as any).value ?? null, formula: null, styleId: before?.styleId ?? 0 };
    if (inputEquals(before, after)) continue;
    deltas.push({ sheetId, row, col, before, after });
  }

  if (deltas.length === 0) return;
  // These updates already happened in the backend. Apply them without creating a new undo step,
  // and tag them so the workbook sync bridge doesn't echo them back.
  document.applyExternalDeltas(deltas, { source: "backend", markDirty: false });
}

async function sendEditsViaTauri(invoke: TauriInvoke, edits: PendingEdit[]): Promise<any[]> {
  if (edits.length === 0) return [];

  /** @type {any[]} */
  const collected = [];

  const bySheet = new Map<string, PendingEdit[]>();
  for (const edit of edits) {
    const list = bySheet.get(edit.sheetId) ?? [];
    list.push(edit);
    bySheet.set(edit.sheetId, list);
  }

  for (const [sheetId, sheetEdits] of bySheet.entries()) {
    sheetEdits.sort(sortPendingEdits);

    const rect = isFullRectangle(sheetEdits);
    if (rect) {
      const byCoord = new Map<string, RangeCellEdit>();
      for (const e of sheetEdits) {
        byCoord.set(`${e.row},${e.col}`, e.edit);
      }

      const values: RangeCellEdit[][] = [];
      for (let r = rect.startRow; r <= rect.endRow; r++) {
        const row: RangeCellEdit[] = [];
        for (let c = rect.startCol; c <= rect.endCol; c++) {
          const edit = byCoord.get(`${r},${c}`);
          if (!edit) {
            throw new Error("Invariant violated: rectangle expected to include all edits");
          }
          row.push(edit);
        }
        values.push(row);
      }

      const updates = await invoke("set_range", {
        sheet_id: sheetId,
        start_row: rect.startRow,
        start_col: rect.startCol,
        end_row: rect.endRow,
        end_col: rect.endCol,
        values
      });
      if (Array.isArray(updates) && updates.length > 0) {
        collected.push(...updates);
      }
      continue;
    }

    // Non-rectangular: fall back to per-cell updates.
    const results = await Promise.all(
      sheetEdits.map((edit) =>
        invoke("set_cell", {
          sheet_id: sheetId,
          row: edit.row,
          col: edit.col,
          value: edit.edit.value,
          formula: edit.edit.formula
        }).catch(() => null),
      ),
    );
    for (const result of results) {
      if (Array.isArray(result) && result.length > 0) {
        collected.push(...result);
      }
    }
  }

  return collected;
}

type SheetSnapshot = { order: string[]; metaById: Map<string, SheetMetaState> };

type SheetSyncAction =
  | { kind: "delta"; sheetMetaDeltas: SheetMetaDelta[]; sheetOrderDelta: SheetOrderDelta | null; source?: string }
  | { kind: "applyState"; snapshot: SheetSnapshot };

function captureSheetSnapshot(document: DocumentControllerLike): SheetSnapshot | null {
  if (typeof document.getSheetIds !== "function" || typeof document.getSheetMeta !== "function") return null;
  const order = document.getSheetIds();
  const metaById = new Map<string, SheetMetaState>();
  for (const rawId of order) {
    const id = String(rawId ?? "").trim();
    if (!id) continue;
    const meta = document.getSheetMeta(id);
    metaById.set(id, normalizeSheetMeta(meta, id));
  }
  const normalizedOrder = order.map((id) => String(id ?? "").trim()).filter(Boolean);
  // DocumentController sheets are lazily materialized: a freshly constructed controller may
  // report zero sheets until the first access/edit. Treat an empty snapshot as "unknown"
  // so we don't accidentally filter out the first batch of cell edits.
  if (normalizedOrder.length === 0) return null;
  return { order: normalizedOrder, metaById };
}

function applySheetDeltaToSnapshot(snapshot: SheetSnapshot, deltas: SheetMetaDelta[], sheetOrderDelta: SheetOrderDelta | null): SheetSnapshot {
  const metaById = new Map(snapshot.metaById);
  for (const delta of deltas) {
    const sheetId = String((delta as any)?.sheetId ?? "").trim();
    if (!sheetId) continue;
    const after = (delta as any)?.after;
    if (!after) {
      metaById.delete(sheetId);
    } else {
      metaById.set(sheetId, normalizeSheetMeta(after, sheetId));
    }
  }

  let order = snapshot.order.slice();
  if (sheetOrderDelta && Array.isArray((sheetOrderDelta as any).after)) {
    const desiredRaw = (sheetOrderDelta as any).after as any[];
    const seen = new Set<string>();
    const desired: string[] = [];
    for (const raw of desiredRaw) {
      const id = String(raw ?? "").trim();
      if (!id) continue;
      if (seen.has(id)) continue;
      if (!metaById.has(id)) continue;
      seen.add(id);
      desired.push(id);
    }
    // Preserve any remaining sheets (shouldn't happen, but keep deterministic behavior).
    for (const id of metaById.keys()) {
      if (seen.has(id)) continue;
      desired.push(id);
    }
    order = desired;
  } else {
    // No explicit ordering delta: keep prior order but drop deleted sheets and append new ones.
    const seen = new Set<string>();
    const nextOrder: string[] = [];
    for (const id of order) {
      if (!metaById.has(id)) continue;
      if (seen.has(id)) continue;
      seen.add(id);
      nextOrder.push(id);
    }
    for (const id of metaById.keys()) {
      if (seen.has(id)) continue;
      nextOrder.push(id);
    }
    order = nextOrder;
  }

  return { order, metaById };
}

function safeShowToast(message: string): void {
  try {
    if (typeof document === "undefined") return;
    showToast(message, "error", { timeoutMs: 8_000 });
  } catch {
    // Best-effort: tests / environments may not have a toast root.
  }
}

async function sendFormattingViaTauri(
  invoke: TauriInvoke,
  deltas: {
    cellDeltas: CellFormatDelta[];
    rowStyleDeltas: RowStyleDelta[];
    colStyleDeltas: ColStyleDelta[];
    sheetStyleDeltas: SheetStyleDelta[];
    rangeRunDeltas: RangeRunDelta[];
  },
): Promise<void> {
  const hasAny =
    deltas.cellDeltas.length > 0 ||
    deltas.rowStyleDeltas.length > 0 ||
    deltas.colStyleDeltas.length > 0 ||
    deltas.sheetStyleDeltas.length > 0 ||
    deltas.rangeRunDeltas.length > 0;
  if (!hasAny) return;

  type SheetPayload = {
    sheetId: string;
    defaultFormat?: any | null;
    rowFormats?: Array<{ row: number; format: any | null }>;
    colFormats?: Array<{ col: number; format: any | null }>;
    formatRunsByCol?: Array<{ col: number; runs: Array<{ startRow: number; endRowExclusive: number; format: any | null }> }>;
    cellFormats?: Array<{ row: number; col: number; format: any | null }>;
  };

  const bySheet = new Map<string, SheetPayload>();
  const ensure = (sheetId: string): SheetPayload => {
    let payload = bySheet.get(sheetId);
    if (!payload) {
      payload = { sheetId };
      bySheet.set(sheetId, payload);
    }
    return payload;
  };

  for (const d of deltas.cellDeltas) {
    const p = ensure(d.sheetId);
    if (!p.cellFormats) p.cellFormats = [];
    p.cellFormats.push({ row: d.row, col: d.col, format: d.afterFormat });
  }
  for (const d of deltas.rowStyleDeltas) {
    const p = ensure(d.sheetId);
    if (!p.rowFormats) p.rowFormats = [];
    p.rowFormats.push({ row: d.row, format: d.afterFormat });
  }
  for (const d of deltas.colStyleDeltas) {
    const p = ensure(d.sheetId);
    if (!p.colFormats) p.colFormats = [];
    p.colFormats.push({ col: d.col, format: d.afterFormat });
  }
  for (const d of deltas.sheetStyleDeltas) {
    const p = ensure(d.sheetId);
    p.defaultFormat = d.afterFormat;
  }
  for (const d of deltas.rangeRunDeltas) {
    const p = ensure(d.sheetId);
    if (!p.formatRunsByCol) p.formatRunsByCol = [];
    p.formatRunsByCol.push({
      col: d.col,
      runs: d.afterRuns.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive, format: r.format })),
    });
  }

  const sheetIds = Array.from(bySheet.keys()).sort();
  for (const sheetId of sheetIds) {
    const payload = bySheet.get(sheetId);
    if (!payload) continue;
    payload.cellFormats?.sort((a, b) => (a.row - b.row === 0 ? a.col - b.col : a.row - b.row));
    payload.rowFormats?.sort((a, b) => a.row - b.row);
    payload.colFormats?.sort((a, b) => a.col - b.col);
    payload.formatRunsByCol?.sort((a, b) => a.col - b.col);
    for (const entry of payload.formatRunsByCol ?? []) {
      entry.runs.sort((a, b) => (a.startRow - b.startRow === 0 ? a.endRowExclusive - b.endRowExclusive : a.startRow - b.startRow));
    }

    // Backend expects a single `payload` argument with camelCase keys.
    await invoke("apply_sheet_formatting_deltas", { payload });
  }
}

async function sendSheetViewViaTauri(
  invoke: TauriInvoke,
  deltas: {
    colWidths: ColWidthDelta[];
    rowHeights: RowHeightDelta[];
  },
): Promise<void> {
  if (deltas.colWidths.length === 0 && deltas.rowHeights.length === 0) return;

  type SheetPayload = {
    sheetId: string;
    colWidths?: Array<{ col: number; width: number | null }>;
    rowHeights?: Array<{ row: number; height: number | null }>;
  };

  const bySheet = new Map<string, SheetPayload>();
  const ensure = (sheetId: string): SheetPayload => {
    let payload = bySheet.get(sheetId);
    if (!payload) {
      payload = { sheetId };
      bySheet.set(sheetId, payload);
    }
    return payload;
  };

  for (const d of deltas.colWidths) {
    const p = ensure(d.sheetId);
    if (!p.colWidths) p.colWidths = [];
    p.colWidths.push({ col: d.col, width: d.width });
  }
  for (const d of deltas.rowHeights) {
    const p = ensure(d.sheetId);
    if (!p.rowHeights) p.rowHeights = [];
    p.rowHeights.push({ row: d.row, height: d.height });
  }

  const sheetIds = Array.from(bySheet.keys()).sort();
  for (const sheetId of sheetIds) {
    const payload = bySheet.get(sheetId);
    if (!payload) continue;
    payload.colWidths?.sort((a, b) => a.col - b.col);
    payload.rowHeights?.sort((a, b) => a.row - b.row);

    await invoke("apply_sheet_view_deltas", { payload });
  }
}

export function startWorkbookSync(args: {
  document: DocumentControllerLike;
  // Reserved for future engine-in-worker integration (e.g. skipping backend recalc).
  engineBridge?: unknown;
}): { stop(): void; markSaved(): Promise<void> } {
  const invoke = resolveInvoke(args.engineBridge);
  if (!invoke) {
    return {
      stop() {},
      async markSaved() {
        args.document.markSaved();
      }
    };
  }
  const invokeFn = invoke;

  const pendingCellEdits = new Map<string, PendingEdit>();
  const pendingSheetActions: SheetSyncAction[] = [];

  const pendingCellFormats = new Map<string, CellFormatDelta>();
  const pendingRowStyles = new Map<string, RowStyleDelta>();
  const pendingColStyles = new Map<string, ColStyleDelta>();
  const pendingSheetStyles = new Map<string, SheetStyleDelta>();
  const pendingRangeRuns = new Map<string, RangeRunDelta>();
  const pendingColWidths = new Map<string, ColWidthDelta>();
  const pendingRowHeights = new Map<string, RowHeightDelta>();

  let sheetMirror: SheetSnapshot | null = captureSheetSnapshot(args.document);
  // Track sheet ids that have been explicitly deleted so we can ignore any stale backend updates
  // that reference them (avoids recreating "phantom" sheets via `applyExternalDeltas`).
  const deletedSheetIdsForBackendUpdates = new Set<string>();

  let stopped = false;
  let flushScheduled = false;
  let flushQueued = false;
  let flushPromise: Promise<void> | null = null;

  const stopListening = args.document.on(
    "change",
    ({ deltas, source, sheetMetaDeltas, sheetOrderDelta, sheetViewDeltas, rowStyleDeltas, colStyleDeltas, sheetStyleDeltas, rangeRunDeltas }) => {
    if (stopped) return;
    const hasSheetMetaDeltas = Array.isArray(sheetMetaDeltas) && sheetMetaDeltas.length > 0;
    const hasSheetOrderDelta = Boolean(sheetOrderDelta);

    const backendOriginated = source === "macro" || source === "python" || source === "pivot" || source === "backend";

    // Keep a best-effort set of deleted sheet ids so we can ignore late/stale backend updates
    // that reference a sheet id the user has removed.
    if (hasSheetMetaDeltas) {
      for (const delta of sheetMetaDeltas ?? []) {
        const sheetId = typeof (delta as any)?.sheetId === "string" ? String((delta as any).sheetId).trim() : "";
        if (!sheetId) continue;
        if ((delta as any).after == null) {
          deletedSheetIdsForBackendUpdates.add(sheetId);
        } else {
          deletedSheetIdsForBackendUpdates.delete(sheetId);
        }
      }
    }

    if (source === "applyState") {
      // `applyState` replaces the entire DocumentController snapshot (restore/sync from an external source).
      // Treat this as a reset boundary:
      // - drop any queued edits/sheet actions from the prior state
      // - do not echo applyState cell clears back via set_cell/set_range (can be huge sparse clears)
      // - do not mirror applyState sheet structure changes (create/delete/reorder) back to the backend
      // - refresh our local sheet snapshot after the controller finishes applying the snapshot so
      //   subsequent *user* edits are scoped to the correct sheet ids.
      pendingCellEdits.clear();
      pendingSheetActions.length = 0;
      pendingCellFormats.clear();
      pendingRowStyles.clear();
      pendingColStyles.clear();
      pendingSheetStyles.clear();
      pendingRangeRuns.clear();
      pendingColWidths.clear();
      pendingRowHeights.clear();
      queueMicrotask(() => {
        if (stopped) return;
        const snap = captureSheetSnapshot(args.document);
        if (!snap) return;
        // `applyState` is generally used to apply an authoritative snapshot (e.g. open/reload/collab).
        // Do not echo sheet structure changes back to the backend; instead, treat the post-applyState
        // snapshot as the new baseline for filtering future cell/format deltas.
        sheetMirror = snap;
      });
      return;
    }

    // Queue sheet structure/metadata updates (DocumentController-driven). Historically the desktop UI
    // persisted sheet tab operations directly via `invoke(...)` calls in main.ts. That approach can
    // easily drift out of sync with the DocumentController undo/redo stack, so we now treat the
    // DocumentController sheet deltas as the single source of truth and mirror them to the backend
    // here (while still ignoring changes that originated in the backend itself).
    if (hasSheetMetaDeltas || hasSheetOrderDelta) {
      if (backendOriginated) {
        // Backend already performed these operations. Track them in the mirror so future
        // applyState reconciliations start from the correct baseline.
        if (sheetMirror) {
          sheetMirror = applySheetDeltaToSnapshot(
            sheetMirror,
            (sheetMetaDeltas ?? []) as SheetMetaDelta[],
            (sheetOrderDelta as SheetOrderDelta | null) ?? null,
          );
        }
      } else {
        pendingSheetActions.push({
          kind: "delta",
          sheetMetaDeltas: (sheetMetaDeltas ?? []) as SheetMetaDelta[],
          sheetOrderDelta: (sheetOrderDelta as SheetOrderDelta | null) ?? null,
          source,
        });
        scheduleFlush();
      }
    }

    // Some subsystems (VBA runtime, native Python) execute in the backend and then return
    // cell updates to apply to the frontend DocumentController. Those should not be echoed
    // back to the backend via set_cell/set_range.
    if (backendOriginated) return;

    // Sheet deletion emits per-cell deltas that clear the deleted sheet's sparse cell map.
    // Those should NOT be mirrored to the backend via `set_cell`/`set_range` because:
    // - the desktop shell persists deletions via the dedicated `delete_sheet` command
    // - mirroring sparse clears can be extremely expensive (N-per-cell IPC)
    // - it can race with `delete_sheet` and fail with UnknownSheet errors
    const deletedSheets = new Set<string>();
    const metaDeltas = Array.isArray(sheetMetaDeltas) ? sheetMetaDeltas : [];
    for (const delta of metaDeltas) {
      const sheetId = typeof delta?.sheetId === "string" ? String(delta.sheetId) : "";
      if (!sheetId) continue;
      if ((delta as any)?.after == null) {
        deletedSheets.add(sheetId);
      }
    }

    let didEnqueue = false;
    if (Array.isArray(deltas) && deltas.length > 0) {
      for (const delta of deltas) {
        if (deletedSheets.has(delta.sheetId)) continue;

        if (!inputEquals(delta.before, delta.after)) {
          const edit: PendingEdit = {
            sheetId: delta.sheetId,
            row: delta.row,
            col: delta.col,
            edit: toRangeCellEdit(delta.after),
          };
          pendingCellEdits.set(cellKey(delta.sheetId, delta.row, delta.col), edit);
          didEnqueue = true;
        }

        if ((delta.before?.styleId ?? 0) !== (delta.after?.styleId ?? 0)) {
          const key = cellKey(delta.sheetId, delta.row, delta.col);
          const existing = pendingCellFormats.get(key);
          const beforeFormat = styleIdToFormat(args.document, delta.before?.styleId ?? 0);
          const afterFormat = styleIdToFormat(args.document, delta.after?.styleId ?? 0);
          if (existing) {
            existing.afterFormat = afterFormat;
          } else {
            pendingCellFormats.set(key, {
              sheetId: delta.sheetId,
              row: delta.row,
              col: delta.col,
              beforeFormat,
              afterFormat,
            });
          }
          didEnqueue = true;
        }
      }
    }

    if (Array.isArray(sheetViewDeltas) && sheetViewDeltas.length > 0) {
      const getAxisValue = (raw: unknown): number | null => {
        const n = Number(raw);
        if (!Number.isFinite(n) || n <= 0) return null;
        return n;
      };

      const diffAxis = (
        sheetId: string,
        before: Record<string, number> | undefined,
        after: Record<string, number> | undefined,
        enqueue: (index: number, value: number | null) => void,
      ): void => {
        const beforeKeys = before ? Object.keys(before) : [];
        const afterKeys = after ? Object.keys(after) : [];
        const keys = new Set([...beforeKeys, ...afterKeys]);
        for (const key of keys) {
          const idx = Number(key);
          if (!Number.isInteger(idx) || idx < 0) continue;
          const beforeValue = getAxisValue(before?.[key]);
          const afterValue = getAxisValue(after?.[key]);
          if (beforeValue == null && afterValue == null) continue;
          if (beforeValue != null && afterValue != null && Math.abs(beforeValue - afterValue) <= 1e-6) continue;
          enqueue(idx, afterValue);
        }
      };

      for (const d of sheetViewDeltas) {
        if (!d || typeof d !== "object") continue;
        const sheetId = String((d as any).sheetId ?? "").trim();
        if (!sheetId) continue;
        if (deletedSheets.has(sheetId)) continue;

        const before = ((d as any).before ?? {}) as SheetViewState;
        const after = ((d as any).after ?? {}) as SheetViewState;

        diffAxis(sheetId, before.colWidths, after.colWidths, (col, width) => {
          const key = colWidthKey(sheetId, col);
          pendingColWidths.set(key, { sheetId, col, width });
          didEnqueue = true;
        });
        diffAxis(sheetId, before.rowHeights, after.rowHeights, (row, height) => {
          const key = rowHeightKey(sheetId, row);
          pendingRowHeights.set(key, { sheetId, row, height });
          didEnqueue = true;
        });
      }
    }

    if (Array.isArray(rowStyleDeltas) && rowStyleDeltas.length > 0) {
      for (const d of rowStyleDeltas) {
        if (deletedSheets.has(d.sheetId)) continue;
        const key = rowStyleKey(d.sheetId, d.row);
        const existing = pendingRowStyles.get(key);
        const beforeFormat = styleIdToFormat(args.document, d.beforeStyleId);
        const afterFormat = styleIdToFormat(args.document, d.afterStyleId);
        if (existing) {
          existing.afterFormat = afterFormat;
        } else {
          pendingRowStyles.set(key, { sheetId: d.sheetId, row: d.row, beforeFormat, afterFormat });
        }
        didEnqueue = true;
      }
    }

    if (Array.isArray(colStyleDeltas) && colStyleDeltas.length > 0) {
      for (const d of colStyleDeltas) {
        if (deletedSheets.has(d.sheetId)) continue;
        const key = colStyleKey(d.sheetId, d.col);
        const existing = pendingColStyles.get(key);
        const beforeFormat = styleIdToFormat(args.document, d.beforeStyleId);
        const afterFormat = styleIdToFormat(args.document, d.afterStyleId);
        if (existing) {
          existing.afterFormat = afterFormat;
        } else {
          pendingColStyles.set(key, { sheetId: d.sheetId, col: d.col, beforeFormat, afterFormat });
        }
        didEnqueue = true;
      }
    }

    if (Array.isArray(sheetStyleDeltas) && sheetStyleDeltas.length > 0) {
      for (const d of sheetStyleDeltas) {
        if (deletedSheets.has(d.sheetId)) continue;
        const key = d.sheetId;
        const existing = pendingSheetStyles.get(key);
        const beforeFormat = styleIdToFormat(args.document, d.beforeStyleId);
        const afterFormat = styleIdToFormat(args.document, d.afterStyleId);
        if (existing) {
          existing.afterFormat = afterFormat;
        } else {
          pendingSheetStyles.set(key, { sheetId: d.sheetId, beforeFormat, afterFormat });
        }
        didEnqueue = true;
      }
    }

    if (Array.isArray(rangeRunDeltas) && rangeRunDeltas.length > 0) {
      for (const d of rangeRunDeltas) {
        if (deletedSheets.has(d.sheetId)) continue;
        const key = rangeRunKey(d.sheetId, d.col);
        const convertRuns = (runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>): RangeRun[] =>
          Array.isArray(runs)
            ? runs.map((r) => ({
                startRow: r.startRow,
                endRowExclusive: r.endRowExclusive,
                format: styleIdToFormat(args.document, r.styleId),
              }))
            : [];
        const beforeRuns = convertRuns(d.beforeRuns);
        const afterRuns = convertRuns(d.afterRuns);

        const existing = pendingRangeRuns.get(key);
        if (existing) {
          existing.startRow = Math.min(existing.startRow, d.startRow);
          existing.endRowExclusive = Math.max(existing.endRowExclusive, d.endRowExclusive);
          existing.afterRuns = afterRuns;
        } else {
          pendingRangeRuns.set(key, {
            sheetId: d.sheetId,
            col: d.col,
            startRow: d.startRow,
            endRowExclusive: d.endRowExclusive,
            beforeRuns,
            afterRuns,
          });
        }
        didEnqueue = true;
      }
    }

    if (didEnqueue) scheduleFlush();
  });

  function scheduleFlush(): void {
    if (stopped) return;
    if (flushScheduled) return;
    flushScheduled = true;

    queueMicrotask(() => {
      flushScheduled = false;
      startFlush();
    });
  }

  async function applyBackendSheetOrder(order: ReadonlyArray<string>): Promise<void> {
    if (order.length <= 1) return;

    const desired: string[] = [];
    const seen = new Set<string>();
    for (const raw of order) {
      const id = String(raw ?? "").trim();
      if (!id) continue;
      if (seen.has(id)) continue;
      seen.add(id);
      desired.push(id);
    }
    if (desired.length <= 1) return;

    try {
      await invokeFn("reorder_sheets", { sheet_ids: desired });
      return;
    } catch {
      // Graceful degradation: older backends may not implement reorder_sheets yet.
      for (let i = 0; i < desired.length; i++) {
        const sheetId = desired[i]!;
        await invokeFn("move_sheet", { sheet_id: sheetId, to_index: i });
      }
    }
  }

  function startFlush(): void {
    if (stopped) return;

    if (flushPromise) {
      flushQueued = true;
      return;
    }

    const hasFormatting =
      pendingCellFormats.size > 0 ||
      pendingRowStyles.size > 0 ||
      pendingColStyles.size > 0 ||
      pendingSheetStyles.size > 0 ||
      pendingRangeRuns.size > 0;

    const hasSheetView = pendingColWidths.size > 0 || pendingRowHeights.size > 0;

    if (pendingCellEdits.size === 0 && pendingSheetActions.length === 0 && !hasFormatting && !hasSheetView) {
      return;
    }

    flushPromise = (async () => {
      while (
        pendingCellEdits.size > 0 ||
        pendingSheetActions.length > 0 ||
        pendingCellFormats.size > 0 ||
        pendingRowStyles.size > 0 ||
        pendingColStyles.size > 0 ||
        pendingSheetStyles.size > 0 ||
        pendingRangeRuns.size > 0 ||
        pendingColWidths.size > 0 ||
        pendingRowHeights.size > 0
      ) {
        const sheetActions = pendingSheetActions.splice(0, pendingSheetActions.length);
        const cellBatch = Array.from(pendingCellEdits.values());
        pendingCellEdits.clear();

        const formatBatch = {
          cellDeltas: Array.from(pendingCellFormats.values()).sort(sortCellFormatDeltas),
          rowStyleDeltas: Array.from(pendingRowStyles.values()).sort(sortRowStyleDeltas),
          colStyleDeltas: Array.from(pendingColStyles.values()).sort(sortColStyleDeltas),
          sheetStyleDeltas: Array.from(pendingSheetStyles.values()).sort(sortSheetStyleDeltas),
          rangeRunDeltas: Array.from(pendingRangeRuns.values()).sort(sortRangeRunDeltas),
        };
        pendingCellFormats.clear();
        pendingRowStyles.clear();
        pendingColStyles.clear();
        pendingSheetStyles.clear();
        pendingRangeRuns.clear();

        const viewBatch = {
          colWidths: Array.from(pendingColWidths.values()).sort(sortColWidthDeltas),
          rowHeights: Array.from(pendingRowHeights.values()).sort(sortRowHeightDeltas),
        };
        pendingColWidths.clear();
        pendingRowHeights.clear();

        // Apply sheet add/delete/rename/reorder deltas *before* sending cell edits so the backend
        // has the correct sheet structure for set_cell/set_range (especially for undo/redo of sheet deletes).
        if (sheetActions.length > 0) {
          try {
            for (const action of sheetActions) {
              if (action.kind === "applyState") {
                if (!sheetMirror) {
                  sheetMirror = action.snapshot;
                  continue;
                }
                const next = action.snapshot;
                const existing = new Set(sheetMirror.order);
                const desired = new Set(next.order);

                // Delete removed sheets first so we never apply cell edits to them.
                for (const sheetId of sheetMirror.order) {
                  if (!desired.has(sheetId)) {
                    await invokeFn("delete_sheet", { sheet_id: sheetId });
                  }
                }

                // Insert added sheets (stable ids).
                for (let i = 0; i < next.order.length; i++) {
                  const sheetId = next.order[i]!;
                  if (existing.has(sheetId)) continue;
                  const meta = next.metaById.get(sheetId) ?? { name: sheetId, visibility: "visible" };
                  await invokeFn("add_sheet", { sheet_id: sheetId, name: meta.name, index: i });
                  if (meta.visibility && meta.visibility !== "visible") {
                    await invokeFn("set_sheet_visibility", { sheet_id: sheetId, visibility: meta.visibility });
                  }
                  const tabColor = tabColorToBackendArg(meta.tabColor);
                  await invokeFn("set_sheet_tab_color", { sheet_id: sheetId, tab_color: tabColor });
                }

                // Update metadata on remaining sheets.
                for (const sheetId of next.order) {
                  if (!existing.has(sheetId)) continue;
                  const before = sheetMirror.metaById.get(sheetId);
                  const after = next.metaById.get(sheetId);
                  if (!before || !after) continue;
                  if (before.name !== after.name) {
                    await invokeFn("rename_sheet", { sheet_id: sheetId, name: after.name });
                  }
                  if (before.visibility !== after.visibility) {
                    await invokeFn("set_sheet_visibility", { sheet_id: sheetId, visibility: after.visibility });
                  }
                  if (!tabColorEquals(before.tabColor, after.tabColor)) {
                    await invokeFn("set_sheet_tab_color", { sheet_id: sheetId, tab_color: tabColorToBackendArg(after.tabColor) });
                  }
                }

                // Apply canonical ordering.
                await applyBackendSheetOrder(next.order);

                sheetMirror = next;
                continue;
              }

              // action.kind === "delta"
              const sheetMetaDeltas = action.sheetMetaDeltas;
              const sheetOrderDelta = action.sheetOrderDelta;

              // Delete sheets first.
              for (const delta of sheetMetaDeltas) {
                const sheetId = String((delta as any)?.sheetId ?? "").trim();
                if (!sheetId) continue;
                const before = (delta as any)?.before;
                const after = (delta as any)?.after;
                if (before && !after) {
                  await invokeFn("delete_sheet", { sheet_id: sheetId });
                }
              }

              // Insert sheets with stable ids.
              for (const delta of sheetMetaDeltas) {
                const sheetId = String((delta as any)?.sheetId ?? "").trim();
                if (!sheetId) continue;
                const before = (delta as any)?.before;
                const after = (delta as any)?.after;
                if (!before && after) {
                  const meta = normalizeSheetMeta(after, sheetId);
                  const desiredIndex = sheetOrderDelta?.after?.indexOf(sheetId) ?? -1;
                  const toIndex = desiredIndex >= 0 ? desiredIndex : (sheetMirror?.order.length ?? 0);
                  await invokeFn("add_sheet", { sheet_id: sheetId, name: meta.name, index: toIndex });
                  if (meta.visibility && meta.visibility !== "visible") {
                    await invokeFn("set_sheet_visibility", { sheet_id: sheetId, visibility: meta.visibility });
                  }
                  const tabColor = tabColorToBackendArg(meta.tabColor);
                  await invokeFn("set_sheet_tab_color", { sheet_id: sheetId, tab_color: tabColor });
                }
              }

              // Apply metadata updates (rename, visibility, tabColor).
              for (const delta of sheetMetaDeltas) {
                const sheetId = String((delta as any)?.sheetId ?? "").trim();
                if (!sheetId) continue;
                const before = (delta as any)?.before;
                const after = (delta as any)?.after;
                if (!before || !after) continue;
                const beforeMeta = normalizeSheetMeta(before, sheetId);
                const afterMeta = normalizeSheetMeta(after, sheetId);

                if (beforeMeta.name !== afterMeta.name) {
                  await invokeFn("rename_sheet", { sheet_id: sheetId, name: afterMeta.name });
                }
                if (beforeMeta.visibility !== afterMeta.visibility) {
                  await invokeFn("set_sheet_visibility", { sheet_id: sheetId, visibility: afterMeta.visibility });
                }
                if (!tabColorEquals(beforeMeta.tabColor, afterMeta.tabColor)) {
                  await invokeFn("set_sheet_tab_color", { sheet_id: sheetId, tab_color: tabColorToBackendArg(afterMeta.tabColor) });
                }
              }

              // Apply canonical ordering.
              if (sheetOrderDelta && Array.isArray(sheetOrderDelta.after)) await applyBackendSheetOrder(sheetOrderDelta.after);

              if (sheetMirror) {
                sheetMirror = applySheetDeltaToSnapshot(sheetMirror, sheetMetaDeltas, sheetOrderDelta);
              }
            }
          } catch (err) {
            console.error("[formula][desktop] Failed to sync sheet changes to backend:", err);
            safeShowToast("Failed to sync workbook sheet changes to the desktop backend. Saving may be inconsistent.");
            throw err;
          }
        }

        // DocumentController materializes sheets lazily (the sheet map can be empty until the first cell is accessed).
        // Refresh our local sheet snapshot here so we don't accidentally drop the first edits to a newly-created sheet.
        // (Without this, `sheetMirror.order` can be `[]` at startup and we'd filter out all cell edits.)
        const refreshedSnapshot = captureSheetSnapshot(args.document);
        if (refreshedSnapshot) sheetMirror = refreshedSnapshot;

        // `DocumentController` materializes sheets lazily: a brand new controller reports zero
        // sheets until the first cell is accessed/edited. In that state, `sheetMirror.order`
        // can be empty even though the backend workbook has a default sheet. Treat an empty
        // mirror as "unknown" rather than "no sheets" so we don't silently drop edits.
        const existingSheetIds =
          sheetMirror && sheetMirror.order.length > 0 ? new Set(sheetMirror.order) : null;
        const filteredCellBatch = existingSheetIds ? cellBatch.filter((e) => existingSheetIds.has(e.sheetId)) : cellBatch;

        const filterSheet = <T extends { sheetId: string }>(items: T[]): T[] =>
          existingSheetIds ? items.filter((d) => existingSheetIds.has(d.sheetId)) : items;
        const filteredFormatBatch = {
          cellDeltas: filterSheet(formatBatch.cellDeltas),
          rowStyleDeltas: filterSheet(formatBatch.rowStyleDeltas),
          colStyleDeltas: filterSheet(formatBatch.colStyleDeltas),
          sheetStyleDeltas: filterSheet(formatBatch.sheetStyleDeltas),
          rangeRunDeltas: filterSheet(formatBatch.rangeRunDeltas),
        };
        const filteredViewBatch = {
          colWidths: filterSheet(viewBatch.colWidths),
          rowHeights: filterSheet(viewBatch.rowHeights),
        };

        const updates = await sendEditsViaTauri(invokeFn, filteredCellBatch);
        applyBackendUpdates(args.document, updates, { skipSheetIds: deletedSheetIdsForBackendUpdates });
        await sendFormattingViaTauri(invokeFn, filteredFormatBatch);
        await sendSheetViewViaTauri(invokeFn, filteredViewBatch);
      }

      // If the user undoes back to the last-saved state, the DocumentController becomes clean
      // again. The backend AppState dirty flag is a simple boolean that only resets when
      // explicitly marked/saved, so we clear it here to keep close prompts aligned.
      if (!args.document.isDirty) {
        try {
          // Yield to the event loop once before calling `mark_saved`. Some backend sync work can be
          // queued in microtasks (e.g. coalesced reorders, other listeners), and we do not want to
          // clear the backend dirty flag *before* those commands execute; otherwise the backend can
          // remain dirty while the document is clean.
          await new Promise<void>((resolve) => {
            if (typeof setTimeout === "function") {
              setTimeout(resolve, 0);
            } else {
              queueMicrotask(resolve);
            }
          });
          if (stopped) return;
          if (args.document.isDirty) return;
          const hasPendingFormatting =
            pendingCellFormats.size > 0 ||
            pendingRowStyles.size > 0 ||
            pendingColStyles.size > 0 ||
            pendingSheetStyles.size > 0 ||
            pendingRangeRuns.size > 0;
          const hasPendingSheetView = pendingColWidths.size > 0 || pendingRowHeights.size > 0;
          if (pendingCellEdits.size > 0 || pendingSheetActions.length > 0 || hasPendingFormatting || hasPendingSheetView || flushQueued) return;
          await invokeFn("mark_saved", {});
        } catch {
          // Graceful degradation: older backends may not implement this command.
        }
      }
    })();

    flushPromise
      .catch((err) => {
        // Avoid unhandled rejections, but don't silently swallow errors (the backend can now be out of sync).
        console.error("[formula][desktop] Workbook sync failed:", err);
        safeShowToast("Failed to sync workbook changes to the desktop backend. Saving may be inconsistent.");
      })
      .finally(() => {
        flushPromise = null;
        if (stopped) return;
        const hasPendingFormatting =
          pendingCellFormats.size > 0 ||
          pendingRowStyles.size > 0 ||
          pendingColStyles.size > 0 ||
          pendingSheetStyles.size > 0 ||
          pendingRangeRuns.size > 0;
        const hasPendingSheetView = pendingColWidths.size > 0 || pendingRowHeights.size > 0;
        if (pendingCellEdits.size > 0 || pendingSheetActions.length > 0 || hasPendingFormatting || hasPendingSheetView || flushQueued) {
          flushQueued = false;
          scheduleFlush();
        }
      });
  }

  async function flushAllPending(): Promise<void> {
    // Ensure we start any scheduled flush promptly.
    startFlush();
    while (flushPromise) {
      await flushPromise;
      startFlush();
    }
  }

  return {
    stop() {
      stopped = true;
      pendingCellEdits.clear();
      pendingSheetActions.length = 0;
      pendingCellFormats.clear();
      pendingRowStyles.clear();
      pendingColStyles.clear();
      pendingSheetStyles.clear();
      pendingRangeRuns.clear();
      pendingColWidths.clear();
      pendingRowHeights.clear();
      stopListening();
    },
    async markSaved() {
      await flushAllPending();
      await invokeFn("save_workbook", {});
      args.document.markSaved();
    }
  };
}
