import { DefaultMacroSecurityController } from "./security";
import type { MacroCellUpdate, MacroSecurityStatus, MacroTrustDecision } from "./types";
import { normalizeFormulaTextOpt } from "@formula/engine";

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

type UiContext = {
  sheetId: string;
  activeRow: number;
  activeCol: number;
  selection: Rect;
};

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

function nonNegativeInt(value: unknown): number {
  const num = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(num)) return 0;
  const floored = Math.floor(num);
  if (!Number.isSafeInteger(floored) || floored < 0) return 0;
  return floored;
}

function normalizeFormulaText(formula: unknown): string | null {
  if (typeof formula !== "string") return null;
  return normalizeFormulaTextOpt(formula);
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

function getUiSelectionRectFromState(active: { row: number; col: number }, ranges: Range[]): Rect {
  const containing =
    ranges.find(
      (r) =>
        active.row >= Math.min(r.startRow, r.endRow) &&
        active.row <= Math.max(r.startRow, r.endRow) &&
        active.col >= Math.min(r.startCol, r.endCol) &&
        active.col <= Math.max(r.startCol, r.endCol),
    ) ?? ranges[0];
  const first = containing ?? { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
  return normalizeRect({
    startRow: nonNegativeInt(first.startRow),
    startCol: nonNegativeInt(first.startCol),
    endRow: nonNegativeInt(first.endRow),
    endCol: nonNegativeInt(first.endCol),
  });
}

function getUiSelectionRect(app: SpreadsheetAppLike): Rect {
  const active = app.getActiveCell();
  return getUiSelectionRectFromState(active, app.getSelectionRanges());
}

async function setMacroUiContext(args: InstallVbaEventMacrosArgs, context?: UiContext): Promise<void> {
  const sheetId = context?.sheetId ?? args.app.getCurrentSheetId();
  const active = context ? { row: context.activeRow, col: context.activeCol } : args.app.getActiveCell();
  const selectionRaw =
    context?.selection ?? getUiSelectionRectFromState(active, context ? [] : args.app.getSelectionRanges());
  const selection = normalizeRect({
    startRow: nonNegativeInt(selectionRaw.startRow),
    startCol: nonNegativeInt(selectionRaw.startCol),
    endRow: nonNegativeInt(selectionRaw.endRow),
    endCol: nonNegativeInt(selectionRaw.endCol),
  });

  try {
    await args.invoke("set_macro_ui_context", {
      workbook_id: args.workbookId,
      sheet_id: sheetId,
      active_row: nonNegativeInt(active.row),
      active_col: nonNegativeInt(active.col),
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

      // Avoid resurrecting deleted sheets: only apply backend macro updates when the sheet exists.
      const docAny: any = doc as any;
      const sheetMeta = typeof docAny.getSheetMeta === "function" ? docAny.getSheetMeta(sheetId) : null;
      if (!sheetMeta) {
        const ids = typeof docAny.getSheetIds === "function" ? docAny.getSheetIds() : [];
        if (Array.isArray(ids) && ids.length > 0) continue;
      }

      const before = typeof docAny.peekCell === "function" ? docAny.peekCell(sheetId, { row, col }) : doc.getCell(sheetId, { row, col });
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
  let currentEventMacroKind:
    | "workbook_open"
    | "worksheet_change"
    | "selection_change"
    | "workbook_before_close"
    | null = null;
  let applyingMacroUpdates = false;
  let eventsDisabled = false;

  let macroStatus: MacroSecurityStatus | null = null;
  let eventsAllowed = true;
  let promptedForTrust = false;
  let workbookOpenFired = false;
  // `await` always yields, even when the awaited promise is already resolved. Track when macro
  // security status has been loaded so we can avoid extra microtask turns that can re-order event
  // scheduling during startup (e.g. Workbook_Open racing Worksheet_Change / SelectionChange flush).
  let macroStatusResolved = false;

  let pendingChangesBySheet = new Map<string, Rect>();
  let pendingWorksheetContexts = new Map<string, UiContext>();
  let changeFlushScheduled = false;
  let changeQueuedAfterMacro = false;
  let changeFlushPromise: Promise<void> | null = null;
  let changeFlushQueued = false;
  let changeFlushDeferredUntilIdle = false;

  let selectionTimer: ReturnType<typeof setTimeout> | null = null;
  let pendingSelectionContext: UiContext | null = null;
  let pendingSelectionKey = "";
  let lastSelectionKeyFired = "";
  let selectionQueuedAfterMacro = false;
  let sawInitialSelection = false;
  let selectionFlushPromise: Promise<void> | null = null;
  let selectionFlushQueued = false;
  let selectionFlushDeferredUntilIdle = false;
  const statusPromise = (async () => {
    try {
      const statusRaw = await args.invoke("get_macro_security_status", { workbook_id: args.workbookId });
      macroStatus = normalizeMacroSecurityStatus(statusRaw);
      eventsAllowed = !macroStatus.hasMacros || macroTrustAllowsRun(macroStatus);
    } catch (err) {
      console.warn("Failed to read macro security status for event macros:", err);
      macroStatus = null;
      eventsAllowed = true;
    } finally {
      macroStatusResolved = true;
    }
  })();

  async function runEventMacro(
    kind: "workbook_open" | "worksheet_change" | "selection_change" | "workbook_before_close",
    cmd: string,
    cmdArgs: Record<string, unknown>,
    uiContext?: UiContext,
  ): Promise<void> {
    // `await statusPromise` always yields a microtask (even if already resolved),
    // which can introduce races between Workbook_Open and events queued while the
    // status request was pending. Once the status is resolved, avoid the extra
    // yield so `runningEventMacro` is set synchronously and queued events can
    // reliably defer without dropping their pending state.
    if (!macroStatusResolved) await statusPromise;
    if (disposed) return;
    if (eventsDisabled) return;

    if ((kind === "worksheet_change" || kind === "selection_change" || kind === "workbook_before_close") && !eventsAllowed) {
      return;
    }

    if (runningEventMacro) {
      // Avoid re-entrancy; callers will re-schedule after the current macro finishes.
      if (kind === "worksheet_change") {
        // Worksheet change flushes batch and clear their pending map before awaiting
        // `runEventMacro()`. If we bail out here (because another event macro is already
        // running), we must restore the rect into the pending map or it can be dropped
        // permanently.
        const sheetId = String(cmdArgs["sheet_id"] ?? "").trim();
        const startRow = Number(cmdArgs["start_row"]);
        const startCol = Number(cmdArgs["start_col"]);
        const endRow = Number(cmdArgs["end_row"]);
        const endCol = Number(cmdArgs["end_col"]);
        if (
          sheetId &&
          Number.isInteger(startRow) &&
          startRow >= 0 &&
          Number.isInteger(startCol) &&
          startCol >= 0 &&
          Number.isInteger(endRow) &&
          endRow >= 0 &&
          Number.isInteger(endCol) &&
          endCol >= 0
        ) {
          const prev = pendingChangesBySheet.get(sheetId);
          let next = unionCell(prev, startRow, startCol);
          next = unionCell(next, endRow, endCol);
          pendingChangesBySheet.set(sheetId, next);

          // Preserve the captured UI context for this pending sheet change, so a delayed rerun
          // (after another event macro finishes) doesn't incorrectly sync the macro runtime to the
          // current UI state instead of the state observed when the edit happened.
          if (uiContext) {
            pendingWorksheetContexts.set(
              sheetId,
              uiContext.sheetId === sheetId ? uiContext : { ...uiContext, sheetId }
            );
          }
        }

        changeQueuedAfterMacro = true;
      }
      if (kind === "selection_change") selectionQueuedAfterMacro = true;
      return;
    }

    runningEventMacro = true;
    currentEventMacroKind = kind;
    try {
      // Allow microtask-batched workbook operations (e.g. `startWorkbookSync` queueing `set_cell`)
      // to attach to the backend sync promise chain before we drain it, so the VBA runtime sees
      // the latest persisted workbook state.
      //
      // Use a Promise-based yield (rather than `queueMicrotask`) to stay compatible with Vitest
      // fake timers and other environments that stub `queueMicrotask`.
      if (kind !== "selection_change") {
        await Promise.resolve();
      }

      // Sync any pending workbook changes (backend sync chain) and the current UI context before
      // firing the macro. Selection changes are sourced from the UI (not the backend sync chain),
      // so draining backend sync here only adds latency.
      if (kind === "selection_change") {
        await setMacroUiContext(args, uiContext);
      } else {
        await Promise.all([args.drainBackendSync(), setMacroUiContext(args, uiContext)]);
      }

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
      currentEventMacroKind = null;
      if (disposed) return;

      if (changeQueuedAfterMacro) {
        changeQueuedAfterMacro = false;
        scheduleFlushWorksheetChanges();
      }
      if (selectionQueuedAfterMacro && !selectionTimer) {
        selectionQueuedAfterMacro = false;
        startFlushSelectionChange();
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
      if (!disposed && !eventsDisabled && !previous) {
        if (changeFlushDeferredUntilIdle && pendingChangesBySheet.size > 0) {
          changeFlushDeferredUntilIdle = false;
          scheduleFlushWorksheetChanges();
        } else {
          changeFlushDeferredUntilIdle = false;
        }

        if (selectionFlushDeferredUntilIdle && pendingSelectionContext && !selectionTimer) {
          selectionFlushDeferredUntilIdle = false;
          startFlushSelectionChange();
        } else {
          selectionFlushDeferredUntilIdle = false;
        }
      }
    }
  }

  async function fireWorkbookOpen(): Promise<void> {
    if (!macroStatusResolved) await statusPromise;
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

  async function doFlushWorksheetChanges(): Promise<void> {
    if (!macroStatusResolved) await statusPromise;
    if (disposed) return;
    if (eventsDisabled) {
      pendingChangesBySheet.clear();
      pendingWorksheetContexts.clear();
      return;
    }
    if (!eventsAllowed) {
      pendingChangesBySheet.clear();
      pendingWorksheetContexts.clear();
      return;
    }
    if (applyingMacroUpdates) {
      changeFlushDeferredUntilIdle = true;
      return;
    }

    if (runningEventMacro) {
      changeQueuedAfterMacro = true;
      return;
    }

    if (pendingChangesBySheet.size === 0) return;
    const entries = Array.from(pendingChangesBySheet.entries());
    const contexts = pendingWorksheetContexts;
    pendingChangesBySheet = new Map();
    pendingWorksheetContexts = new Map();

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
      }, contexts.get(sheetId));
    }
  }

  function scheduleFlushWorksheetChanges(): void {
    if (disposed) return;
    if (changeFlushScheduled) return;
    changeFlushScheduled = true;
    void Promise.resolve()
      .then(() => {
        changeFlushScheduled = false;
        startFlushWorksheetChanges();
      })
      .catch(() => {
        // Best-effort: avoid unhandled rejections from the microtask scheduler.
      });
  }

  function startFlushWorksheetChanges(): void {
    if (disposed) return;
    if (changeFlushPromise) {
      changeFlushQueued = true;
      return;
    }
    if (pendingChangesBySheet.size === 0) return;
    changeFlushPromise = doFlushWorksheetChanges()
      .catch((err) => {
        console.warn("Failed to flush Worksheet_Change event macros:", err);
      })
      .finally(() => {
        changeFlushPromise = null;
        if (disposed) return;
        if (eventsDisabled) {
          pendingChangesBySheet.clear();
          pendingWorksheetContexts.clear();
          changeFlushQueued = false;
          return;
        }

        const shouldRerun = pendingChangesBySheet.size > 0 || changeFlushQueued;
        changeFlushQueued = false;
        if (!shouldRerun) return;

        if (runningEventMacro || changeQueuedAfterMacro) {
          // Defer until the current macro finishes; `runEventMacro` will re-schedule.
          return;
        }

        if (applyingMacroUpdates) {
          changeFlushDeferredUntilIdle = true;
          return;
        }

        scheduleFlushWorksheetChanges();
      });
  }

  async function doFlushSelectionChange(): Promise<void> {
    if (!macroStatusResolved) await statusPromise;
    if (disposed) return;
    if (eventsDisabled) {
      pendingSelectionContext = null;
      pendingSelectionKey = "";
      return;
    }
    if (!eventsAllowed) {
      pendingSelectionContext = null;
      pendingSelectionKey = "";
      return;
    }
    if (!pendingSelectionContext) return;
    if (applyingMacroUpdates) {
      selectionFlushDeferredUntilIdle = true;
      return;
    }

    if (runningEventMacro) {
      selectionQueuedAfterMacro = true;
      return;
    }

    const context = pendingSelectionContext;
    const key = pendingSelectionKey;
    pendingSelectionContext = null;
    pendingSelectionKey = "";

    if (!context) return;
    if (key && key === lastSelectionKeyFired) return;

    await runEventMacro("selection_change", "fire_selection_change", {
      sheet_id: context.sheetId,
      start_row: context.selection.startRow,
      start_col: context.selection.startCol,
      end_row: context.selection.endRow,
      end_col: context.selection.endCol,
    }, context);

    lastSelectionKeyFired = key;
  }

  function startFlushSelectionChange(): void {
    if (disposed) return;
    if (selectionFlushPromise) {
      selectionFlushQueued = true;
      return;
    }
    if (!pendingSelectionContext) return;
    selectionFlushPromise = doFlushSelectionChange()
      .catch((err) => {
        console.warn("Failed to flush SelectionChange event macros:", err);
      })
      .finally(() => {
        selectionFlushPromise = null;
        if (disposed) return;

        if (eventsDisabled) {
          pendingSelectionContext = null;
          pendingSelectionKey = "";
          selectionFlushQueued = false;
          return;
        }

        const shouldRerun = Boolean(pendingSelectionContext) || selectionFlushQueued;
        selectionFlushQueued = false;
        if (!shouldRerun) return;

        if (runningEventMacro || selectionQueuedAfterMacro) {
          // Defer until the current macro finishes; `runEventMacro` will re-schedule.
          return;
        }

        if (applyingMacroUpdates) {
          selectionFlushDeferredUntilIdle = true;
          return;
        }

        void Promise.resolve()
          .then(() => startFlushSelectionChange())
          .catch(() => {
            // Best-effort: avoid unhandled rejections from the microtask scheduler.
          });
      });
  }

  const stopDocListening = args.app.getDocument().on("change", ({ deltas, source }) => {
    if (disposed) return;
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    if (eventsDisabled) return;
    if (applyingMacroUpdates) return;
    if (
      source === "applyState" ||
      source === "macro" ||
      source === "python" ||
      source === "backend" ||
      source === "collab" ||
      source === "undo" ||
      source === "redo" ||
      source === "cancelBatch"
    )
      return;

    // Only run Worksheet_Change when macros are already trusted.
    if (!eventsAllowed) return;

    const active = args.app.getActiveCell();
    const selectionRect = getUiSelectionRectFromState(active, args.app.getSelectionRanges());

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
      pendingWorksheetContexts.set(sheetId, {
        sheetId,
        activeRow: active.row,
        activeCol: active.col,
        selection: selectionRect,
      });
    }

    scheduleFlushWorksheetChanges();
  });

  const stopSelectionListening = args.app.subscribeSelection((selection) => {
    if (disposed) return;
    if (!selection || !Array.isArray(selection.ranges)) return;
    if (eventsDisabled) return;
    if (!eventsAllowed) return;
    if (applyingMacroUpdates) return;
    // Avoid infinite recursion where a `Worksheet_SelectionChange` event macro changes the
    // selection, which would ordinarily trigger another `Worksheet_SelectionChange`.
    if (runningEventMacro && currentEventMacroKind === "selection_change") return;

    if (!sawInitialSelection) {
      sawInitialSelection = true;
      return;
    }

    const active = args.app.getActiveCell();
    const rect = getUiSelectionRectFromState(active, selection.ranges);

    const sheetId = args.app.getCurrentSheetId();
    const key = `${sheetId}:${rect.startRow},${rect.startCol}:${rect.endRow},${rect.endCol}@${active.row},${active.col}`;
    if (key === pendingSelectionKey) return;

    pendingSelectionKey = key;
    pendingSelectionContext = { sheetId, activeRow: active.row, activeCol: active.col, selection: rect };

    if (selectionTimer) clearTimeout(selectionTimer);
    selectionTimer = setTimeout(() => {
      selectionTimer = null;
      startFlushSelectionChange();
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
      pendingWorksheetContexts.clear();
      pendingSelectionContext = null;
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

  // Use a Promise-based microtask yield instead of `queueMicrotask` so this stays compatible
  // with Vitest fake timers (which may stub `queueMicrotask`).
  await Promise.resolve();
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
