import { normalizeFormulaTextOpt } from "@formula/engine";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

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
  on(event: "change", listener: (payload: { deltas: CellDelta[]; source?: string; recalc?: boolean }) => void): () => void;
  markSaved(): void;
  readonly isDirty: boolean;
  // Optional APIs on the real DocumentController used to apply authoritative backend updates
  // (e.g. pivot auto-refresh output).
  getCell?(sheetId: string, coord: { row: number; col: number }): any;
  applyExternalDeltas?(deltas: any[], options?: { source?: string }): void;
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
  document.applyExternalDeltas(deltas, { source: "backend" });
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

  const pending = new Map<string, PendingEdit>();

  let stopped = false;
  let flushScheduled = false;
  let flushQueued = false;
  let flushPromise: Promise<void> | null = null;

  const stopListening = args.document.on("change", ({ deltas, source }) => {
    if (stopped) return;
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
      pending.set(cellKey(delta.sheetId, delta.row, delta.col), edit);
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

    if (pending.size === 0) {
      return;
    }

    flushPromise = (async () => {
      while (pending.size > 0) {
        const batch = Array.from(pending.values());
        pending.clear();
        const updates = await sendEditsViaTauri(invokeFn, batch);
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
      .catch(() => {
        // Swallow errors to avoid unhandled rejections; logging handled by the caller/UI.
      })
      .finally(() => {
        flushPromise = null;
        if (stopped) return;
        if (pending.size > 0 || flushQueued) {
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
      pending.clear();
      stopListening();
    },
    async markSaved() {
      await flushAllPending();
      await invokeFn("save_workbook", {});
      args.document.markSaved();
    }
  };
}
