import { normalizeFormulaTextOpt } from "@formula/engine";
import { showToast } from "../extensions/ui.js";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

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
  // DocumentController tracks styleId too, but the desktop workbook IPC currently
  // only supports value/formula edits.
  styleId?: number;
};

type CellDelta = {
  sheetId: string;
  row: number;
  col: number;
  before: CellState;
  after: CellState;
};

type DocumentControllerLike = {
  on(
    event: "change",
    listener: (payload: {
      deltas: CellDelta[];
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
  // Optional APIs on the real DocumentController used to apply authoritative backend updates
  // (e.g. pivot auto-refresh output).
  getCell?(sheetId: string, coord: { row: number; col: number }): any;
  applyExternalDeltas?(deltas: any[], options?: { source?: string; markDirty?: boolean }): void;
};

type RangeCellEdit = { value: unknown | null; formula: string | null };

function getTauriInvoke(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return invoke ?? null;
}

function resolveInvoke(engineBridge: unknown): TauriInvoke | null {
  if (engineBridge && typeof engineBridge === "object") {
    const maybe = (engineBridge as any).invoke;
    if (typeof maybe === "function") {
      return maybe as TauriInvoke;
    }
  }
  return getTauriInvoke();
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

// NOTE: In desktop mode, sheet metadata operations (rename/reorder/add/delete/hide/tabColor)
// are persisted to the backend by `main.ts` (both for direct UI actions and doc-driven
// undo/redo/applyState reconciliations). Workbook sync only needs to:
// - track sheet ids to avoid sending cell edits for deleted sheets
// - ensure `mark_saved` runs when undo/redo returns to the last-saved state, even when the
//   last change was sheet-metadata-only (no cell deltas).

type PendingEdit = { sheetId: string; row: number; col: number; edit: RangeCellEdit };

function toRangeCellEdit(state: CellState): RangeCellEdit {
  if (state.formula != null) {
    const normalized = normalizeFormulaTextOpt(state.formula);
    if (normalized != null) {
      return { value: null, formula: normalized };
    }
  }
  return { value: (state.value ?? null) as unknown | null, formula: null };
}

function normalizeFormulaText(formula: unknown): string | null {
  if (typeof formula !== "string") return null;
  return normalizeFormulaTextOpt(formula);
}

function cellKey(sheetId: string, row: number, col: number): string {
  return `${sheetId}:${row},${col}`;
}

function sortPendingEdits(a: PendingEdit, b: PendingEdit): number {
  if (a.sheetId < b.sheetId) return -1;
  if (a.sheetId > b.sheetId) return 1;
  if (a.row !== b.row) return a.row - b.row;
  return a.col - b.col;
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

function applyBackendUpdates(document: DocumentControllerLike, raw: unknown): void {
  if (typeof document.getCell !== "function" || typeof document.applyExternalDeltas !== "function") return;
  if (!Array.isArray(raw) || raw.length === 0) return;

  const deltas: any[] = [];
  for (const u of raw as any[]) {
    if (!u || typeof u !== "object") continue;
    const sheetId = String((u as any).sheet_id ?? "").trim();
    const row = Number((u as any).row);
    const col = Number((u as any).col);
    if (!sheetId) continue;
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    // Backend returns computed value updates for formula cells; the frontend has its own calc engine.
    // We only apply input changes for non-formula cells (e.g. pivot output values).
    if (normalizeFormulaText((u as any).formula) != null) continue;

    const before = document.getCell(sheetId, { row, col });
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
  | { kind: "delta"; sheetMetaDeltas: SheetMetaDelta[]; sheetOrderDelta: SheetOrderDelta | null; source: "undo" | "redo" }
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
  return { order: order.map((id) => String(id ?? "").trim()).filter(Boolean), metaById };
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

  let sheetMirror: SheetSnapshot | null = captureSheetSnapshot(args.document);

  let stopped = false;
  let flushScheduled = false;
  let flushQueued = false;
  let flushPromise: Promise<void> | null = null;

  const stopListening = args.document.on("change", ({ deltas, source, sheetMetaDeltas, sheetOrderDelta }) => {
    if (stopped) return;
    const hasSheetMetaDeltas = Array.isArray(sheetMetaDeltas) && sheetMetaDeltas.length > 0;
    const hasSheetOrderDelta = Boolean(sheetOrderDelta);

    // Queue sheet structure/metadata updates for undo/redo events (DocumentController-driven).
    if ((source === "undo" || source === "redo") && (hasSheetMetaDeltas || hasSheetOrderDelta)) {
      pendingSheetActions.push({
        kind: "delta",
        sheetMetaDeltas: (sheetMetaDeltas ?? []) as SheetMetaDelta[],
        sheetOrderDelta: (sheetOrderDelta as SheetOrderDelta | null) ?? null,
        source,
      });
      scheduleFlush();
    } else if (source === "applyState") {
      // `applyState` replaces the entire DocumentController snapshot. It can create/delete sheets
      // without emitting sheetMetaDeltas/sheetOrderDelta, so reconcile against a post-applyState
      // snapshot in a microtask (after the controller finishes removing deleted sheets).
      //
      // Treat this as a reset boundary: pending backend sync for the old document state is no longer
      // meaningful, so drop any queued edits and reconcile to the new snapshot.
      pendingCellEdits.clear();
      pendingSheetActions.length = 0;

      queueMicrotask(() => {
        if (stopped) return;
        const snap = captureSheetSnapshot(args.document);
        if (!snap) return;
        pendingSheetActions.push({ kind: "applyState", snapshot: snap });
        scheduleFlush();
      });
    } else if (sheetMirror && (hasSheetMetaDeltas || hasSheetOrderDelta)) {
      // For non-undo/redo sheet changes, assume the UI already persisted them to the backend
      // via direct Tauri invocations (main.ts). Keep our local mirror in sync so future applyState
      // reconciliations have a reasonable baseline.
      sheetMirror = applySheetDeltaToSnapshot(
        sheetMirror,
        (sheetMetaDeltas ?? []) as SheetMetaDelta[],
        (sheetOrderDelta as SheetOrderDelta | null) ?? null,
      );
    }

    // Some subsystems (VBA runtime, native Python) execute in the backend and then return
    // cell updates to apply to the frontend DocumentController. Those should not be echoed
    // back to the backend via set_cell/set_range.
    if (source === "macro" || source === "python" || source === "pivot" || source === "backend") return;
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      // Ignore format-only deltas (we can't mirror those over set_cell/set_range yet).
      if (inputEquals(delta.before, delta.after)) continue;

      const edit: PendingEdit = {
        sheetId: delta.sheetId,
        row: delta.row,
        col: delta.col,
        edit: toRangeCellEdit(delta.after)
      };
      pendingCellEdits.set(cellKey(delta.sheetId, delta.row, delta.col), edit);
    }
    scheduleFlush();
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

  function startFlush(): void {
    if (stopped) return;

    if (flushPromise) {
      flushQueued = true;
      return;
    }

    if (pendingCellEdits.size === 0 && pendingSheetActions.length === 0) {
      return;
    }

    flushPromise = (async () => {
      while (pendingCellEdits.size > 0 || pendingSheetActions.length > 0) {
        const sheetActions = pendingSheetActions.splice(0, pendingSheetActions.length);
        const cellBatch = Array.from(pendingCellEdits.values());
        pendingCellEdits.clear();

         // Update our local sheet snapshot for undo/redo/applyState events so we can avoid
         // sending cell edits for deleted sheets (and include cell restores for re-inserted sheets).
         if (sheetActions.length > 0) {
           for (const action of sheetActions) {
             if (action.kind === "applyState") {
               sheetMirror = action.snapshot;
               continue;
             }
             if (sheetMirror) {
               sheetMirror = applySheetDeltaToSnapshot(sheetMirror, action.sheetMetaDeltas, action.sheetOrderDelta);
             }
           }
         }

        const existingSheetIds = sheetMirror ? new Set(sheetMirror.order) : null;
        const filteredCellBatch = existingSheetIds ? cellBatch.filter((e) => existingSheetIds.has(e.sheetId)) : cellBatch;
        const updates = await sendEditsViaTauri(invokeFn, filteredCellBatch);
        applyBackendUpdates(args.document, updates);
      }

      // If the user undoes back to the last-saved state, the DocumentController becomes clean
      // again. The backend AppState dirty flag is a simple boolean that only resets when
      // explicitly marked/saved, so we clear it here to keep close prompts aligned.
      if (!args.document.isDirty) {
        try {
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
        if (pendingCellEdits.size > 0 || pendingSheetActions.length > 0 || flushQueued) {
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
      stopListening();
    },
    async markSaved() {
      await flushAllPending();
      await invokeFn("save_workbook", {});
      args.document.markSaved();
    }
  };
}
