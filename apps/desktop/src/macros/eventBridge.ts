import type { DocumentController } from "../document/documentController.js";
import { applyMacroCellUpdates } from "./applyUpdates";
import type { MacroCellUpdate, MacroPermission } from "./types";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

type MacroPermissionRequest = {
  reason: string;
  macro_id: string;
  workbook_origin_path: string | null;
  requested: MacroPermission[];
};

type MacroError = {
  message: string;
  code?: string | null;
  blocked?: unknown;
};

type TauriMacroRunResult = {
  ok: boolean;
  output: string[];
  updates?: MacroCellUpdate[];
  error?: MacroError;
  permission_request?: MacroPermissionRequest | null;
};

export type SelectionRect = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

function selectionKey(selection: SelectionRect): string {
  return `${selection.sheetId}:${selection.startRow},${selection.startCol}-${selection.endRow},${selection.endCol}`;
}

type CellState = { value: unknown; formula: string | null };
type CellDelta = {
  sheetId: string;
  row: number;
  col: number;
  before: CellState;
  after: CellState;
};

type DocumentChangePayload = {
  deltas: CellDelta[];
  source?: string;
};

type BannerAction = { label: string; onClick: () => void };

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

function inputEquals(before: CellState, after: CellState): boolean {
  return valuesEqual(before.value ?? null, after.value ?? null) && (before.formula ?? null) === (after.formula ?? null);
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

function normalizeTauriMacroResult(raw: any): TauriMacroRunResult {
  const output = Array.isArray(raw?.output) ? raw.output.map(String) : [];
  const updates = normalizeUpdates(raw?.updates);
  const errorRaw = raw?.error;
  const error =
    errorRaw && typeof errorRaw === "object"
      ? {
          message: String(errorRaw.message ?? errorRaw),
          code: errorRaw.code != null ? String(errorRaw.code) : undefined,
          blocked: (errorRaw as any).blocked,
        }
      : errorRaw
        ? { message: String(errorRaw) }
        : undefined;

  const requestRaw = raw?.permission_request;
  const permission_request =
    requestRaw && typeof requestRaw === "object"
      ? {
          reason: String(requestRaw.reason ?? ""),
          macro_id: String(requestRaw.macro_id ?? ""),
          workbook_origin_path: requestRaw.workbook_origin_path != null ? String(requestRaw.workbook_origin_path) : null,
          requested: Array.isArray(requestRaw.requested)
            ? (requestRaw.requested.map(String) as MacroPermission[])
            : [],
        }
      : undefined;

  return { ok: Boolean(raw?.ok), output, updates, error, permission_request };
}

function ensureBannerContainer(): HTMLElement | null {
  if (typeof document === "undefined") return null;
  const existing = document.getElementById("macro-event-banner-container");
  if (existing) return existing;

  const container = document.createElement("div");
  container.id = "macro-event-banner-container";
  container.style.position = "fixed";
  container.style.bottom = "12px";
  container.style.left = "12px";
  container.style.zIndex = "9999";
  container.style.display = "flex";
  container.style.flexDirection = "column";
  container.style.gap = "8px";
  document.body.appendChild(container);
  return container;
}

function findBanner(container: HTMLElement, key: string): HTMLElement | null {
  for (const child of Array.from(container.children)) {
    if (!(child instanceof HTMLElement)) continue;
    if (child.dataset.macroBannerKey === key) return child;
  }
  return null;
}

function removeBanner(key: string): void {
  const container = ensureBannerContainer();
  if (!container) return;
  findBanner(container, key)?.remove();
}

function showBanner(key: string, message: string, actions: BannerAction[] = [], opts: { autoDismissMs?: number } = {}) {
  const container = ensureBannerContainer();
  if (!container) {
    console.info(message);
    return;
  }

  const existing = findBanner(container, key);
  const banner = existing ?? document.createElement("div");
  banner.dataset.macroBannerKey = key;
  banner.style.background = "rgba(20, 20, 20, 0.9)";
  banner.style.color = "white";
  banner.style.padding = "10px 12px";
  banner.style.borderRadius = "8px";
  banner.style.fontSize = "12px";
  banner.style.display = "flex";
  banner.style.alignItems = "center";
  banner.style.gap = "10px";
  banner.style.maxWidth = "420px";
  banner.style.boxShadow = "0 4px 14px rgba(0, 0, 0, 0.3)";

  const text = document.createElement("div");
  text.textContent = message;
  text.style.flex = "1";

  const buttons = document.createElement("div");
  buttons.style.display = "flex";
  buttons.style.gap = "6px";
  buttons.replaceChildren();

  for (const action of actions) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = action.label;
    btn.style.fontSize = "12px";
    btn.style.padding = "4px 8px";
    btn.style.borderRadius = "6px";
    btn.style.border = "1px solid rgba(255,255,255,0.2)";
    btn.style.background = "rgba(255,255,255,0.08)";
    btn.style.color = "white";
    btn.addEventListener("click", action.onClick);
    buttons.appendChild(btn);
  }

  banner.replaceChildren(text, buttons);
  if (!existing) container.appendChild(banner);

  const dismiss = () => {
    banner.remove();
  };

  if (opts.autoDismissMs != null && opts.autoDismissMs > 0) {
    globalThis.setTimeout(dismiss, opts.autoDismissMs);
  }
}

export class MacroEventBridge {
  private readonly workbookId: string;
  private readonly document: DocumentController;
  private readonly invoke: TauriInvoke;
  private readonly drainBackendSync: () => Promise<void>;
  private readonly getSelection: () => SelectionRect;

  private unsubscribeDocChange: (() => void) | null = null;
  private worksheetChangeTimer: ReturnType<typeof setTimeout> | null = null;
  private selectionChangeTimer: ReturnType<typeof setTimeout> | null = null;

  private pendingSheetChanges = new Map<string, SelectionRect>();
  private pendingSheetContexts = new Map<string, SelectionRect>();
  private pendingSelection: SelectionRect | null = null;
  private lastSelectionKey: string | null = null;

  private suppressWorksheetEventsDepth = 0;
  private lastBlockedBannerAt = 0;

  private eventQueue: Promise<void> = Promise.resolve();
  private inEventDepth = 0;
  private readonly recentEventTimestamps: number[] = [];

  constructor(args: {
    workbookId: string;
    document: DocumentController;
    invoke: TauriInvoke;
    drainBackendSync: () => Promise<void>;
    getSelection: () => SelectionRect;
    debounceWorksheetMs?: number;
    debounceSelectionMs?: number;
  }) {
    this.workbookId = args.workbookId;
    this.document = args.document;
    this.invoke = args.invoke;
    this.drainBackendSync = args.drainBackendSync;
    this.getSelection = args.getSelection;
    this.debounceWorksheetMs = args.debounceWorksheetMs ?? 200;
    this.debounceSelectionMs = args.debounceSelectionMs ?? 150;
  }

  private readonly debounceWorksheetMs: number;
  private readonly debounceSelectionMs: number;

  start(): void {
    if (this.unsubscribeDocChange) return;
    try {
      this.lastSelectionKey = selectionKey(this.getSelection());
    } catch {
      // Ignore selection lookup failures (tests/non-UI environments may stub this differently).
    }
    this.unsubscribeDocChange = this.document.on("change", (payload: DocumentChangePayload) => {
      this.onDocumentChange(payload);
    });
  }

  /**
   * Resolves once all currently enqueued macro events have finished running.
   *
   * Primarily intended for tests and one-off flows (e.g. close handling).
   */
  async whenIdle(): Promise<void> {
    await this.eventQueue;
  }

  stop(): void {
    this.unsubscribeDocChange?.();
    this.unsubscribeDocChange = null;
    if (this.worksheetChangeTimer) globalThis.clearTimeout(this.worksheetChangeTimer);
    if (this.selectionChangeTimer) globalThis.clearTimeout(this.selectionChangeTimer);
    this.worksheetChangeTimer = null;
    this.selectionChangeTimer = null;
    this.pendingSheetChanges.clear();
    this.pendingSheetContexts.clear();
    this.pendingSelection = null;
    this.lastSelectionKey = null;
  }

  notifySelectionChanged(selection: SelectionRect): void {
    const key = selectionKey(selection);
    if (key === this.lastSelectionKey) return;
    this.lastSelectionKey = key;

    this.pendingSelection = selection;
    if (this.selectionChangeTimer) globalThis.clearTimeout(this.selectionChangeTimer);
    this.selectionChangeTimer = globalThis.setTimeout(() => {
      this.selectionChangeTimer = null;
      const next = this.pendingSelection;
      this.pendingSelection = null;
      if (!next) return;
      this.enqueue(async () => {
        await this.fireSelectionChange(next);
      });
    }, this.debounceSelectionMs);
  }

  fireWorkbookOpen(): Promise<void> {
    return this.enqueue(async () => {
      const selection = this.getSelection();
      await this.fireMacroEvent({
        kind: "Workbook_Open",
        cmd: "fire_workbook_open",
        args: { workbook_id: this.workbookId },
        selection,
      });
    });
  }

  async fireWorkbookBeforeClose(): Promise<{ ran: boolean; permissionRequest?: MacroPermissionRequest | null }> {
    return this.enqueue(async () => {
      const selection = this.getSelection();
      const result = await this.fireMacroEvent({
        kind: "Workbook_BeforeClose",
        cmd: "fire_workbook_before_close",
        args: { workbook_id: this.workbookId },
        selection,
      });
      if (result.permission_request) {
        return { ran: false, permissionRequest: result.permission_request };
      }
      return { ran: true };
    });
  }

  applyMacroUpdates(updates: readonly MacroCellUpdate[], options: { label: string }): void {
    if (!Array.isArray(updates) || updates.length === 0) return;
    this.withWorksheetChangeSuppressed(() => {
      this.document.beginBatch({ label: options.label });
      let committed = false;
      try {
        applyMacroCellUpdates(this.document, updates);
        committed = true;
      } finally {
        if (committed) this.document.endBatch();
        else this.document.cancelBatch();
      }
    });
  }

  private withWorksheetChangeSuppressed(fn: () => void): void {
    this.suppressWorksheetEventsDepth += 1;
    try {
      fn();
    } finally {
      this.suppressWorksheetEventsDepth = Math.max(0, this.suppressWorksheetEventsDepth - 1);
    }
  }

  private onDocumentChange(payload: DocumentChangePayload): void {
    if (this.suppressWorksheetEventsDepth > 0) return;
    if (!payload || payload.source === "applyState") return;
    const deltas = Array.isArray(payload.deltas) ? payload.deltas : [];
    if (deltas.length === 0) return;

    let selectionContext: SelectionRect | null = null;
    try {
      selectionContext = this.getSelection();
    } catch {
      selectionContext = null;
    }

    for (const delta of deltas) {
      if (!delta || typeof delta !== "object") continue;
      if (inputEquals(delta.before, delta.after)) continue;

      const sheetId = String(delta.sheetId ?? "").trim();
      if (!sheetId) continue;

      const row = Number(delta.row);
      const col = Number(delta.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      const existing = this.pendingSheetChanges.get(sheetId);
      const next: SelectionRect = existing
        ? {
            sheetId,
            startRow: Math.min(existing.startRow, row),
            startCol: Math.min(existing.startCol, col),
            endRow: Math.max(existing.endRow, row),
            endCol: Math.max(existing.endCol, col),
          }
        : { sheetId, startRow: row, startCol: col, endRow: row, endCol: col };
      this.pendingSheetChanges.set(sheetId, next);
      if (selectionContext) {
        this.pendingSheetContexts.set(sheetId, selectionContext);
      }
    }

    if (this.pendingSheetChanges.size === 0) return;
    if (this.worksheetChangeTimer) return;

    this.worksheetChangeTimer = globalThis.setTimeout(() => {
      this.worksheetChangeTimer = null;
      const batch = new Map(this.pendingSheetChanges);
      const contexts = new Map(this.pendingSheetContexts);
      this.pendingSheetChanges.clear();
      this.pendingSheetContexts.clear();
      if (batch.size === 0) return;
      this.enqueue(async () => {
        for (const rect of batch.values()) {
          let ctx = contexts.get(rect.sheetId) ?? null;
          if (!ctx) {
            try {
              ctx = this.getSelection();
            } catch {
              ctx = rect;
            }
          }
          await this.fireWorksheetChange(rect, ctx);
        }
      });
    }, this.debounceWorksheetMs);
  }

  private enqueue<T>(task: () => Promise<T>): Promise<T> {
    const next = this.eventQueue.then(task, task);
    this.eventQueue = next.then(
      () => undefined,
      () => undefined,
    );
    return next;
  }

  private async setMacroUiContext(selection: SelectionRect): Promise<void> {
    try {
      await this.invoke("set_macro_ui_context", {
        workbook_id: this.workbookId,
        sheet_id: selection.sheetId,
        active_row: selection.startRow,
        active_col: selection.startCol,
        selection: {
          start_row: selection.startRow,
          start_col: selection.startCol,
          end_row: selection.endRow,
          end_col: selection.endCol,
        },
      });
    } catch (err) {
      // Older backends may not implement UI context sync; event macros should still run.
      console.warn("Failed to sync macro UI context:", err);
    }
  }

  private async fireWorksheetChange(rect: SelectionRect, selection: SelectionRect): Promise<void> {
    await this.fireMacroEvent({
      kind: "Worksheet_Change",
      cmd: "fire_worksheet_change",
      args: {
        workbook_id: this.workbookId,
        sheet_id: rect.sheetId,
        start_row: rect.startRow,
        start_col: rect.startCol,
        end_row: rect.endRow,
        end_col: rect.endCol,
      },
      selection,
    });
  }

  private async fireSelectionChange(selection: SelectionRect): Promise<void> {
    await this.fireMacroEvent({
      kind: "Worksheet_SelectionChange",
      cmd: "fire_selection_change",
      args: {
        workbook_id: this.workbookId,
        sheet_id: selection.sheetId,
        start_row: selection.startRow,
        start_col: selection.startCol,
        end_row: selection.endRow,
        end_col: selection.endCol,
      },
      selection,
    });
  }

  private async fireMacroEvent(args: {
    kind: string;
    cmd: string;
    args: Record<string, unknown>;
    selection: SelectionRect;
    permissions?: MacroPermission[];
  }): Promise<TauriMacroRunResult> {
    // Guard against runaway recursion / event storms (e.g. event macros that indirectly
    // trigger more events). The frontend already debounces worksheet + selection events,
    // but this provides an additional safety net.
    if (this.inEventDepth >= 5) {
      console.warn(`Skipping macro event ${args.kind}: recursion limit reached.`);
      return { ok: true, output: [] };
    }
    const now = Date.now();
    while (this.recentEventTimestamps.length > 0 && now - this.recentEventTimestamps[0]! > 2_000) {
      this.recentEventTimestamps.shift();
    }
    if (this.recentEventTimestamps.length >= 25) {
      console.warn(`Skipping macro event ${args.kind}: rate limit exceeded.`);
      return { ok: true, output: [] };
    }
    this.recentEventTimestamps.push(now);

    this.inEventDepth += 1;
    try {
    // Allow microtask-batched edits to enqueue into the backend sync queue first.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await this.drainBackendSync();

    await this.setMacroUiContext(args.selection);

    const raw = await this.invoke(args.cmd, {
      ...args.args,
      permissions: args.permissions,
    });

    const result = normalizeTauriMacroResult(raw);

    if (result.permission_request) {
      const req = result.permission_request;
      const requestedList = req.requested.length > 0 ? req.requested.join(", ") : "additional permissions";
      showBanner(
        `macro-permission:${req.macro_id}:${args.kind}`,
        `${args.kind} requested ${requestedList}.`,
        [
          {
            label: "Allow & rerun",
            onClick: () => {
              removeBanner(`macro-permission:${req.macro_id}:${args.kind}`);
              void this.enqueue(async () => {
                await this.fireMacroEvent({
                  ...args,
                  permissions: req.requested,
                });
              });
            },
          },
          { label: "Dismiss", onClick: () => removeBanner(`macro-permission:${req.macro_id}:${args.kind}`) },
        ],
      );
      return result;
    }

    if (result.error?.code === "macro_blocked") {
      const now = Date.now();
      if (now - this.lastBlockedBannerAt > 10_000) {
        this.lastBlockedBannerAt = now;
        showBanner("macro-blocked", "Macros are blocked by Trust Center policy.", [], { autoDismissMs: 6_000 });
      }
      return result;
    }

    if (result.updates && result.updates.length > 0) {
      this.applyMacroUpdates(result.updates, { label: `Macro event: ${args.kind}` });
    }

    if (!result.ok && result.error) {
      console.error(`Macro event failed (${args.kind}):`, result.error.message);
      showBanner(`macro-error:${args.kind}`, `Macro event ${args.kind} failed: ${result.error.message}`, [], { autoDismissMs: 8_000 });
    }

    return result;
    } finally {
      this.inEventDepth = Math.max(0, this.inEventDepth - 1);
    }
  }
}
