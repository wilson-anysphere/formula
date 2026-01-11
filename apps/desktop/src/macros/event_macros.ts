import { DefaultMacroSecurityController } from "./security";
import type { MacroCellUpdate, MacroSecurityStatus, MacroTrustDecision } from "./types";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

type Range = { startRow: number; startCol: number; endRow: number; endCol: number };

type SelectionState = { ranges: Range[] };

type CellState = { value: unknown; formula: string | null };

type CellDelta = {
  sheetId: string;
  row: number;
  col: number;
  before: CellState;
  after: CellState;
};

type DocumentControllerLike = {
  on(event: "change", listener: (payload: { deltas: CellDelta[]; source?: string; recalc?: boolean }) => void): () => void;
  // These APIs are supported by the real DocumentController, but tests/embedders may
  // provide lighter stubs.
  getCell?(sheetId: string, coord: { row: number; col: number }): any;
  applyExternalDeltas?(deltas: any[], options?: { source?: string }): void;
  beginBatch?(options?: { label?: string }): void;
  endBatch?(): void;
  cancelBatch?(): void;
};

type SpreadsheetAppLike = {
  getCurrentSheetId(): string;
  getActiveCell(): { row: number; col: number };
  getSelectionRanges(): Range[];
  subscribeSelection(listener: (selection: SelectionState) => void): () => void;
  getDocument(): DocumentControllerLike;
  refresh(): void;
  whenIdle(): Promise<void>;
};

export interface InstallVbaEventMacrosArgs {
  app: SpreadsheetAppLike;
  workbookId: string;
  invoke: TauriInvoke;
  drainBackendSync: () => Promise<void>;
}

export interface VbaEventMacrosHandle {
  dispose(): void;
  applyMacroUpdates(updates: readonly MacroCellUpdate[], options?: { label?: string }): Promise<void>;
}

type Rect = { startRow: number; startCol: number; endRow: number; endCol: number };

const EVENT_MACRO_BATCH_LABEL = "VBA event macro";
const WORKBOOK_OPEN_EVENT_ID = "Workbook_Open";
const WORKBOOK_BEFORE_CLOSE_EVENT_ID = "Workbook_BeforeClose";
const SELECTION_CHANGE_DEBOUNCE_MS = 100;

function errorMessage(err: unknown): string {
  if (typeof err === "string") return err;
  if (err instanceof Error) return err.message;
  if (err && typeof err === "object" && "message" in err) {
    try {
      return String((err as any).message);
    } catch {
      return "Unknown error";
    }
  }
  try {
    return String(err);
  } catch {
    return "Unknown error";
  }
}

function isNoWorkbookLoadedError(err: unknown): boolean {
  return errorMessage(err).toLowerCase().includes("no workbook loaded");
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

function inputEquals(before: CellState | null | undefined, after: CellState | null | undefined): boolean {
  return valuesEqual(before?.value ?? null, after?.value ?? null) && (before?.formula ?? null) === (after?.formula ?? null);
}

function normalizeRect(rect: Rect): Rect {
  const startRow = Math.min(rect.startRow, rect.endRow);
  const endRow = Math.max(rect.startRow, rect.endRow);
  const startCol = Math.min(rect.startCol, rect.endCol);
  const endCol = Math.max(rect.startCol, rect.endCol);
  return { startRow, startCol, endRow, endCol };
}

function normalizeFormulaText(formula: unknown): string | null {
  if (formula == null) return null;
  if (typeof formula !== "string") return null;
  const trimmed = formula.trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

function unionCell(rect: Rect | undefined, row: number, col: number): Rect {
  if (!rect) {
    return { startRow: row, startCol: col, endRow: row, endCol: col };
  }
  return {
    startRow: Math.min(rect.startRow, row),
    startCol: Math.min(rect.startCol, col),
    endRow: Math.max(rect.endRow, row),
    endCol: Math.max(rect.endCol, col),
  };
}

function getUiSelectionRect(app: SpreadsheetAppLike): Rect {
  const active = app.getActiveCell();
  const ranges = app.getSelectionRanges();
  const first = ranges[0] ?? { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
  return normalizeRect({
    startRow: Number(first.startRow) || 0,
    startCol: Number(first.startCol) || 0,
    endRow: Number(first.endRow) || 0,
    endCol: Number(first.endCol) || 0,
  });
}

async function setMacroUiContext(args: InstallVbaEventMacrosArgs): Promise<void> {
  const sheetId = args.app.getCurrentSheetId();
  const active = args.app.getActiveCell();
  const selection = getUiSelectionRect(args.app);

  try {
    await args.invoke("set_macro_ui_context", {
      workbook_id: args.workbookId,
      sheet_id: sheetId,
      active_row: active.row,
      active_col: active.col,
      selection: {
        start_row: selection.startRow,
        start_col: selection.startCol,
        end_row: selection.endRow,
        end_col: selection.endCol,
      },
    });
  } catch (err) {
    if (isNoWorkbookLoadedError(err)) return;
    // Older backends may not expose UI context sync (event macros still run, but
    // `ActiveCell`/`Selection` may be stale).
    console.warn("Failed to sync macro UI context:", err);
  }
}

function normalizeUpdates(raw: any[] | undefined): MacroCellUpdate[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0) return undefined;
  const out: MacroCellUpdate[] = [];
  for (const u of raw) {
    if (!u || typeof u !== "object") continue;
    const sheetId = String((u as any).sheet_id ?? "").trim();
    const row = Number((u as any).row);
    const col = Number((u as any).col);
    if (!sheetId) continue;
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    out.push({
      sheetId,
      row,
      col,
      value: (u as any).value ?? null,
      formula: typeof (u as any).formula === "string" ? (u as any).formula : null,
      displayValue: String((u as any).display_value ?? ""),
    });
  }
  return out.length > 0 ? out : undefined;
}

function normalizeMacroSecurityStatus(raw: any): MacroSecurityStatus {
  return {
    hasMacros: Boolean(raw?.has_macros),
    originPath: typeof raw?.origin_path === "string" ? String(raw.origin_path) : undefined,
    workbookFingerprint: typeof raw?.workbook_fingerprint === "string" ? String(raw.workbook_fingerprint) : undefined,
    signature: raw?.signature
      ? {
          status: typeof raw.signature.status === "string" ? raw.signature.status : "unsigned",
          signerSubject: typeof raw.signature.signer_subject === "string" ? String(raw.signature.signer_subject) : undefined,
          signatureBase64:
            typeof raw.signature.signature_base64 === "string" ? String(raw.signature.signature_base64) : undefined,
        }
      : undefined,
    trust: typeof raw?.trust === "string" ? (raw.trust as MacroTrustDecision) : "blocked",
  };
}

function macroTrustAllowsRun(status: MacroSecurityStatus): boolean {
  switch (status.trust) {
    case "trusted_always":
    case "trusted_once":
      return true;
    case "trusted_signed_only": {
      const sig = status.signature?.status;
      return sig === "signed_verified" || sig === "signed_untrusted";
    }
    case "blocked":
    default:
      return false;
  }
}

type MacroEventResult = {
  ok: boolean;
  updates?: MacroCellUpdate[];
  blocked: boolean;
  permissionRequest: boolean;
};

function normalizeMacroEventResult(raw: any): MacroEventResult {
  const ok = Boolean(raw?.ok);
  const updates = normalizeUpdates(raw?.updates);
  const blocked = Boolean(raw?.error?.blocked);
  const permissionRequest = Boolean(raw?.permission_request);
  return { ok, updates, blocked, permissionRequest };
}

async function applyMacroUpdatesToDocument(
  app: SpreadsheetAppLike,
  updates: readonly MacroCellUpdate[],
  options: { label: string }
): Promise<void> {
  if (!updates.length) return;
  const doc = app.getDocument();

  if (typeof doc.getCell === "function" && typeof doc.applyExternalDeltas === "function") {
    const deltas: any[] = [];
    for (const update of updates) {
      const sheetId = String(update?.sheetId ?? "").trim();
      if (!sheetId) continue;
      const row = Number(update?.row);
      const col = Number(update?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      const before = doc.getCell(sheetId, { row, col });
      const formula = normalizeFormulaText(update.formula);
      const value = formula ? null : (update.value ?? null);
      const after = { value, formula, styleId: before?.styleId ?? 0 };
      if (inputEquals(before, after)) continue;
      deltas.push({ sheetId, row, col, before, after });
    }
    if (deltas.length) {
      // These updates already happened in the backend (VBA runtime). Apply them without
      // creating a new undo step, and tag them so the workbook sync bridge doesn't echo
      // them back to the backend via set_cell/set_range.
      doc.applyExternalDeltas(deltas, { source: "macro" });
    }
  } else if (typeof doc.beginBatch === "function" && typeof doc.endBatch === "function") {
    // Fallback for lightweight embedders/tests: apply updates as a normal batch edit.
    doc.beginBatch({ label: options.label });
    let committed = false;
    try {
      for (const update of updates) {
        const sheetId = String(update?.sheetId ?? "").trim();
        if (!sheetId) continue;
        const row = Number(update?.row);
        const col = Number(update?.col);
        if (!Number.isInteger(row) || row < 0) continue;
        if (!Number.isInteger(col) || col < 0) continue;

        const formula = normalizeFormulaText(update.formula);
        const value = formula ? null : (update.value ?? null);
        (doc as any).setCellInput?.(sheetId, { row, col }, { value, formula });
      }
      committed = true;
    } finally {
      if (committed) doc.endBatch();
      else doc.cancelBatch?.() ?? doc.endBatch();
    }
  }

  app.refresh();
  await app.whenIdle();
  app.refresh();
}

/**
 * Install listeners to fire Excel-compatible VBA event macros for the current workbook.
 */
export function installVbaEventMacros(args: InstallVbaEventMacrosArgs): VbaEventMacrosHandle {
  const security = new DefaultMacroSecurityController();

  let disposed = false;
  let runningEventMacro = false;
  let applyingMacroUpdates = false;
  let eventsDisabled = false;

  let macroStatus: MacroSecurityStatus | null = null;
  let eventsAllowed = true;
  let promptedForTrust = false;
  let workbookOpenFired = false;

  let pendingChangesBySheet = new Map<string, Rect>();
  let changeFlushScheduled = false;
  let changeQueuedAfterMacro = false;

  let selectionTimer: ReturnType<typeof setTimeout> | null = null;
  let pendingSelectionRect: Rect | null = null;
  let pendingSelectionKey = "";
  let lastSelectionKeyFired = "";
  let selectionQueuedAfterMacro = false;
  let sawInitialSelection = false;

  const statusPromise = (async () => {
    try {
      const statusRaw = await args.invoke("get_macro_security_status", { workbook_id: args.workbookId });
      macroStatus = normalizeMacroSecurityStatus(statusRaw);
      eventsAllowed = !macroStatus.hasMacros || macroTrustAllowsRun(macroStatus);
    } catch (err) {
      console.warn("Failed to read macro security status for event macros:", err);
      macroStatus = null;
      eventsAllowed = true;
    }
  })();

  async function runEventMacro(
    kind: "workbook_open" | "worksheet_change" | "selection_change" | "workbook_before_close",
    cmd: string,
    cmdArgs: Record<string, unknown>,
  ): Promise<void> {
    await statusPromise;
    if (disposed) return;
    if (eventsDisabled) return;

    if ((kind === "worksheet_change" || kind === "selection_change" || kind === "workbook_before_close") && !eventsAllowed) {
      return;
    }

    if (runningEventMacro) {
      // Avoid re-entrancy; callers will re-schedule after the current macro finishes.
      if (kind === "worksheet_change") changeQueuedAfterMacro = true;
      if (kind === "selection_change") selectionQueuedAfterMacro = true;
      return;
    }

    runningEventMacro = true;
    try {
      // Allow microtask-batched edits (e.g. `startWorkbookSync`) to enqueue into the backend
      // sync chain before we drain it, so event macros see the latest persisted workbook state.
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      await args.drainBackendSync();
      await setMacroUiContext(args);

      let raw: any;
      try {
        raw = await args.invoke(cmd, {
          workbook_id: args.workbookId,
          permissions: [],
          ...cmdArgs,
        });
      } catch (err) {
        if (isNoWorkbookLoadedError(err)) return;
        console.warn(`VBA event macro (${cmd}) invoke failed:`, err);
        return;
      }

      const result = normalizeMacroEventResult(raw);
      if (result.blocked) {
        eventsDisabled = true;
        return;
      }

      if (result.permissionRequest) {
        console.warn(`VBA event macro (${cmd}) requested additional permissions; refusing to escalate.`);
        eventsDisabled = true;
        return;
      }

      if (result.updates && result.updates.length) {
        await applyMacroUpdates(result.updates, { label: EVENT_MACRO_BATCH_LABEL });
      }

      // If the macro errored without an explicit Trust Center block, keep processing future events.
      // The backend already returns `ok=true` when no event handler exists.
    } finally {
      runningEventMacro = false;
      if (disposed) return;

      if (changeQueuedAfterMacro) {
        changeQueuedAfterMacro = false;
        scheduleFlushWorksheetChanges();
      }
      if (selectionQueuedAfterMacro && !selectionTimer) {
        selectionQueuedAfterMacro = false;
        void flushSelectionChange();
      }
    }
  }

  async function applyMacroUpdates(
    updates: readonly MacroCellUpdate[],
    options: { label: string }
  ): Promise<void> {
    if (!Array.isArray(updates) || updates.length === 0) return;
    if (disposed) return;
    const previous = applyingMacroUpdates;
    applyingMacroUpdates = true;
    try {
      await applyMacroUpdatesToDocument(args.app, updates, options);
    } finally {
      applyingMacroUpdates = previous;
    }
  }

  async function fireWorkbookOpen(): Promise<void> {
    await statusPromise;
    if (disposed) return;
    if (eventsDisabled) return;
    if (workbookOpenFired) return;
    workbookOpenFired = true;

    if (!macroStatus) {
      // Proceed anyway; backend will no-op if no macros exist.
      await runEventMacro("workbook_open", "fire_workbook_open", {});
      return;
    }

    if (macroStatus.hasMacros && !macroTrustAllowsRun(macroStatus)) {
      if (promptedForTrust) {
        eventsAllowed = false;
        eventsDisabled = true;
        return;
      }
      promptedForTrust = true;

      const decision = await security.requestTrustDecision({
        workbookId: args.workbookId,
        macroId: WORKBOOK_OPEN_EVENT_ID,
        status: macroStatus,
      });

      if (!decision || decision === "blocked") {
        eventsAllowed = false;
        eventsDisabled = true;
        return;
      }

      try {
        const statusRaw = await args.invoke("set_macro_trust", { workbook_id: args.workbookId, decision });
        macroStatus = normalizeMacroSecurityStatus(statusRaw);
        eventsAllowed = !macroStatus.hasMacros || macroTrustAllowsRun(macroStatus);
      } catch (err) {
        console.warn("Failed to update macro trust decision:", err);
        eventsAllowed = false;
        eventsDisabled = true;
        return;
      }
    }

    await runEventMacro("workbook_open", "fire_workbook_open", {});
  }

  async function flushWorksheetChanges(): Promise<void> {
    await statusPromise;
    if (disposed) return;
    if (eventsDisabled) return;
    if (!eventsAllowed) {
      pendingChangesBySheet.clear();
      return;
    }
    if (applyingMacroUpdates) return;

    if (runningEventMacro) {
      changeQueuedAfterMacro = true;
      return;
    }

    if (pendingChangesBySheet.size === 0) return;
    const entries = Array.from(pendingChangesBySheet.entries());
    pendingChangesBySheet = new Map();

    for (const [sheetId, rect] of entries) {
      if (disposed || eventsDisabled) return;
      if (pendingChangesBySheet.size > 0) {
        // Avoid starvation: new changes arrived while we were awaiting; let the scheduler pick them up.
        scheduleFlushWorksheetChanges();
      }
      await runEventMacro("worksheet_change", "fire_worksheet_change", {
        sheet_id: sheetId,
        start_row: rect.startRow,
        start_col: rect.startCol,
        end_row: rect.endRow,
        end_col: rect.endCol,
      });
    }
  }

  function scheduleFlushWorksheetChanges(): void {
    if (disposed) return;
    if (changeFlushScheduled) return;
    changeFlushScheduled = true;
    queueMicrotask(() => {
      changeFlushScheduled = false;
      void flushWorksheetChanges();
    });
  }

  async function flushSelectionChange(): Promise<void> {
    await statusPromise;
    if (disposed) return;
    if (eventsDisabled) return;
    if (!eventsAllowed) {
      pendingSelectionRect = null;
      pendingSelectionKey = "";
      return;
    }
    if (!pendingSelectionRect) return;
    if (applyingMacroUpdates) return;

    if (runningEventMacro) {
      selectionQueuedAfterMacro = true;
      return;
    }

    const rect = pendingSelectionRect;
    const key = pendingSelectionKey;
    pendingSelectionRect = null;
    pendingSelectionKey = "";

    if (!rect) return;
    if (key && key === lastSelectionKeyFired) return;

    const sheetId = args.app.getCurrentSheetId();
    await runEventMacro("selection_change", "fire_selection_change", {
      sheet_id: sheetId,
      start_row: rect.startRow,
      start_col: rect.startCol,
      end_row: rect.endRow,
      end_col: rect.endCol,
    });

    lastSelectionKeyFired = key;
  }

  const stopDocListening = args.app.getDocument().on("change", ({ deltas, source }) => {
    if (disposed) return;
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    if (eventsDisabled) return;
    if (applyingMacroUpdates) return;
    if (source === "applyState" || source === "macro" || source === "python") return;

    // Only run Worksheet_Change when macros are already trusted.
    if (!eventsAllowed) return;

    for (const delta of deltas) {
      if (!delta?.before || !delta?.after) continue;
      if (inputEquals(delta.before, delta.after)) continue;
      const sheetId = String(delta?.sheetId ?? "").trim();
      if (!sheetId) continue;
      const row = Number(delta?.row);
      const col = Number(delta?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      const prev = pendingChangesBySheet.get(sheetId);
      pendingChangesBySheet.set(sheetId, unionCell(prev, row, col));
    }

    scheduleFlushWorksheetChanges();
  });

  const stopSelectionListening = args.app.subscribeSelection((selection) => {
    if (disposed) return;
    if (!selection || !Array.isArray(selection.ranges)) return;
    if (eventsDisabled) return;
    if (!eventsAllowed) return;
    if (applyingMacroUpdates) return;

    if (!sawInitialSelection) {
      sawInitialSelection = true;
      return;
    }

    const active = args.app.getActiveCell();
    const first = selection.ranges[0] ?? { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
    const rect = normalizeRect({
      startRow: Number(first.startRow) || 0,
      startCol: Number(first.startCol) || 0,
      endRow: Number(first.endRow) || 0,
      endCol: Number(first.endCol) || 0,
    });

    const sheetId = args.app.getCurrentSheetId();
    const key = `${sheetId}:${rect.startRow},${rect.startCol}:${rect.endRow},${rect.endCol}@${active.row},${active.col}`;
    if (key === pendingSelectionKey) return;

    pendingSelectionKey = key;
    pendingSelectionRect = rect;

    if (selectionTimer) clearTimeout(selectionTimer);
    selectionTimer = setTimeout(() => {
      selectionTimer = null;
      void flushSelectionChange();
    }, SELECTION_CHANGE_DEBOUNCE_MS);
  });

  // Trigger Workbook_Open once per workbook load. This is intentionally "fire and forget"
  // so installation remains synchronous.
  void fireWorkbookOpen().catch((err) => {
    console.warn("Workbook_Open event macro failed:", err);
  });

  return {
    async applyMacroUpdates(updates, options) {
      await applyMacroUpdates(updates, { label: options?.label ?? EVENT_MACRO_BATCH_LABEL });
    },
    dispose() {
      disposed = true;
      stopDocListening();
      stopSelectionListening();
      pendingChangesBySheet.clear();
      pendingSelectionRect = null;
      pendingSelectionKey = "";
      if (selectionTimer) {
        clearTimeout(selectionTimer);
        selectionTimer = null;
      }
    },
  };
}

/**
 * Best-effort Workbook_BeforeClose used by the desktop shell when swapping workbooks.
 *
 * Does not prompt; only runs if the workbook is already trusted.
 */
export async function fireWorkbookBeforeCloseBestEffort(args: InstallVbaEventMacrosArgs): Promise<void> {
  let status: MacroSecurityStatus;
  try {
    const raw = await args.invoke("get_macro_security_status", { workbook_id: args.workbookId });
    status = normalizeMacroSecurityStatus(raw);
  } catch (err) {
    console.warn("Failed to read macro security status before close:", err);
    return;
  }

  if (status.hasMacros && !macroTrustAllowsRun(status)) {
    return;
  }

  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await args.drainBackendSync();
  await setMacroUiContext(args);

  let raw: any;
  try {
    raw = await args.invoke("fire_workbook_before_close", {
      workbook_id: args.workbookId,
      permissions: [],
    });
  } catch (err) {
    if (isNoWorkbookLoadedError(err)) return;
    console.warn("Failed to invoke Workbook_BeforeClose:", err);
    return;
  }

  const result = normalizeMacroEventResult(raw);
  if (result.blocked) return;
  if (result.permissionRequest) {
    console.warn(`VBA event macro (${WORKBOOK_BEFORE_CLOSE_EVENT_ID}) requested additional permissions; refusing to escalate.`);
    return;
  }

  if (result.updates && result.updates.length) {
    await applyMacroUpdatesToDocument(args.app, result.updates, { label: EVENT_MACRO_BATCH_LABEL });
  }
}
