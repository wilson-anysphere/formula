import { CellEditorOverlay } from "../editor/cellEditorOverlay";
import { FormulaBarTabCompletionController } from "../ai/completion/formulaBarTabCompletion.js";
import { FormulaBarView } from "../formula-bar/FormulaBarView";
import { Outline, groupDetailRange, isHidden } from "../grid/outline/outline.js";
import { parseA1Range } from "../charts/a1.js";
import { anchorToRectPx } from "../charts/overlay.js";
import { renderChartSvg } from "../charts/renderSvg.js";
import { ChartStore, type ChartRecord } from "../charts/chartStore";
import { applyPlainTextEdit } from "../grid/text/rich-text/edit.js";
import { renderRichText } from "../grid/text/rich-text/render.js";
import { cellToA1, rangeToA1 } from "../selection/a1";
import { navigateSelectionByKey } from "../selection/navigation";
import { SelectionRenderer } from "../selection/renderer";
import type { CellCoord, GridLimits, Range, SelectionState } from "../selection/types";
import { resolveCssVar } from "../theme/cssVars.js";
import { t, tWithVars } from "../i18n/index.js";
import {
  DEFAULT_GRID_LIMITS,
  addCellToSelection,
  createSelection,
  buildSelection,
  extendSelectionToCell,
  selectAll,
  selectColumns,
  selectRows,
  setActiveCell
} from "../selection/selection";
import { DocumentController } from "../document/documentController.js";
import { MockEngine } from "../document/engine.js";
import { isRedoKeyboardEvent, isUndoKeyboardEvent } from "../document/shortcuts.js";
import {
  createEngineClient,
  engineApplyDeltas,
  engineHydrateFromDocument,
  type EngineClient
} from "@formula/engine";
import { drawCommentIndicator } from "../comments/CommentIndicator";
import { evaluateFormula, type SpreadsheetValue } from "../spreadsheet/evaluateFormula";
import { DocumentWorkbookAdapter } from "../search/documentWorkbookAdapter.js";
import { parseGoTo } from "../../../../packages/search/index.js";
import type { CreateChartResult, CreateChartSpec } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import { colToName as colToNameA1, fromA1 as fromA1A1 } from "@formula/spreadsheet-frontend/a1";

import * as Y from "yjs";
import { CommentManager, bindDocToStorage } from "@formula/collab-comments";
import type { Comment, CommentAuthor } from "@formula/collab-comments";

type EngineCellRef = { sheetId?: string; sheet?: string; row?: number; col?: number; address?: string; value?: unknown };

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return typeof (value as { then?: unknown } | null)?.then === "function";
}

/**
 * Tracks async engine work so tests can deterministically await "engine idle".
 *
 * This intentionally *does not* assume a specific engine implementation; it just
 * aggregates promises representing:
 *  - batched input syncing (e.g. worker `setCells`)
 *  - recalculation
 *  - applying computed CellChange[] into the app's computed-value cache
 */
class IdleTracker {
  private readonly pending = new Set<Promise<unknown>>();
  private waiters: Array<() => void> = [];
  private error: unknown = null;
  private resolveScheduled = false;

  track(value: PromiseLike<unknown>): void {
    const promise = Promise.resolve(value);
    // Prevent unhandled rejection noise if no one is currently awaiting `whenIdle`.
    promise.catch((err) => {
      if (this.error == null) this.error = err;
    });
    this.pending.add(promise);
    promise.finally(() => {
      this.pending.delete(promise);
      this.maybeResolve();
    });
  }

  whenIdle(): Promise<void> {
    if (this.pending.size === 0) {
      return this.error == null ? Promise.resolve() : Promise.reject(this.error);
    }

    return new Promise((resolve, reject) => {
      this.waiters.push(() => {
        if (this.error == null) resolve();
        else reject(this.error);
      });
    });
  }

  private maybeResolve(): void {
    if (this.pending.size !== 0) return;
    if (this.resolveScheduled) return;
    this.resolveScheduled = true;

    // Wait one microtask turn before resolving to allow follow-up work to be
    // tracked (e.g. `.then(() => track(...))` attached after the original
    // promise was tracked).
    queueMicrotask(() => {
      this.resolveScheduled = false;
      if (this.pending.size !== 0) return;
      const waiters = this.waiters;
      this.waiters = [];
      for (const resolve of waiters) resolve();
    });
  }
}

type EngineIntegration = {
  applyChanges?: (changes: unknown) => unknown;
  recalculate?: () => unknown;
  beginBatch?: () => unknown;
  endBatch?: () => unknown;
};

class IdleTrackingEngine {
  recalcCount = 0;

  constructor(
    private readonly inner: EngineIntegration,
    private readonly idle: IdleTracker,
    private readonly onInputsChanged: (changes: unknown) => void,
    private readonly onComputedChanges: (changes: unknown) => void | Promise<void>
  ) {}

  applyChanges(changes: unknown): void {
    this.onInputsChanged(changes);
    const result = this.inner.applyChanges?.(changes);
    if (isThenable(result)) {
      this.idle.track(result);
    }
  }

  recalculate(): void {
    this.recalcCount += 1;
    const result = this.inner.recalculate?.();

    if (isThenable(result)) {
      const tracked = Promise.resolve(result).then((changes) => this.onComputedChanges(changes));
      this.idle.track(tracked);
      return;
    }

    // Some engines may return the changes synchronously.
    if (result !== undefined) {
      const applied = this.onComputedChanges(result);
      if (isThenable(applied)) this.idle.track(applied);
    }
  }

  beginBatch(): void {
    const result = this.inner.beginBatch?.();
    if (isThenable(result)) this.idle.track(result);
  }

  endBatch(): void {
    const result = this.inner.endBatch?.();
    if (isThenable(result)) this.idle.track(result);
  }
}

type ChartSeriesDef = {
  name?: string | null;
  categories?: string | null;
  values?: string | null;
  xValues?: string | null;
  yValues?: string | null;
};

type ChartDef = {
  chartType: { kind: string; name?: string };
  title?: string;
  series: ChartSeriesDef[];
  anchor: {
    kind: string;
    fromCol?: number;
    fromRow?: number;
    fromColOffEmu?: number;
    fromRowOffEmu?: number;
    toCol?: number;
    toRow?: number;
    toColOffEmu?: number;
    toRowOffEmu?: number;
    xEmu?: number;
    yEmu?: number;
    cxEmu?: number;
    cyEmu?: number;
  };
};

export interface SpreadsheetAppStatusElements {
  activeCell: HTMLElement;
  selectionRange: HTMLElement;
  activeValue: HTMLElement;
}

export class SpreadsheetApp {
  private sheetId = "Sheet1";
  private readonly idle = new IdleTracker();
  private readonly computedValues = new Map<string, SpreadsheetValue>();
  private uiReady = false;
  private readonly engine = new IdleTrackingEngine(
    new MockEngine(),
    this.idle,
    (changes) => this.invalidateComputedValues(changes),
    (changes) => this.applyComputedChanges(changes)
  );
  private readonly document = new DocumentController({ engine: this.engine });
  private readonly searchWorkbook = new DocumentWorkbookAdapter({ document: this.document });
  private limits: GridLimits;

  private wasmEngine: EngineClient | null = null;
  private wasmSyncSuspended = false;
  private wasmUnsubscribe: (() => void) | null = null;

  private gridCanvas: HTMLCanvasElement;
  private chartLayer: HTMLDivElement;
  private referenceCanvas: HTMLCanvasElement;
  private selectionCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private referenceCtx: CanvasRenderingContext2D;
  private selectionCtx: CanvasRenderingContext2D;

  private outline = new Outline();
  private outlineLayer: HTMLDivElement;

  private dpr = 1;
  private width = 0;
  private height = 0;

  private readonly cellWidth = 100;
  private readonly cellHeight = 24;
  private readonly rowHeaderWidth = 48;
  private readonly colHeaderHeight = 24;

  private visibleRows: number[] = [];
  private visibleCols: number[] = [];
  private rowToVisual = new Map<number, number>();
  private colToVisual = new Map<number, number>();

  private selection: SelectionState;
  private selectionRenderer = new SelectionRenderer();
  private readonly selectionListeners = new Set<(selection: SelectionState) => void>();

  private editor: CellEditorOverlay;
  private formulaBar: FormulaBarView | null = null;
  private formulaBarCompletion: FormulaBarTabCompletionController | null = null;
  private formulaEditCell: CellCoord | null = null;
  private referencePreview: { start: CellCoord; end: CellCoord } | null = null;

  private dragState: { pointerId: number; mode: "normal" | "formula" } | null = null;

  private resizeObserver: ResizeObserver;

  private readonly currentUser: CommentAuthor;
  private readonly commentsDoc = new Y.Doc();
  private readonly commentManager = new CommentManager(this.commentsDoc);
  private commentCells = new Set<string>();
  private commentsPanelVisible = false;
  private stopCommentPersistence: (() => void) | null = null;

  private readonly chartStore: ChartStore;

  private commentsPanel!: HTMLDivElement;
  private commentsPanelThreads!: HTMLDivElement;
  private commentsPanelCell!: HTMLDivElement;
  private newCommentInput!: HTMLInputElement;
  private commentTooltip!: HTMLDivElement;

  private renderScheduled = false;

  constructor(
    private root: HTMLElement,
    private status: SpreadsheetAppStatusElements,
    opts: { limits?: GridLimits; formulaBar?: HTMLElement } = {}
  ) {
    this.limits = opts.limits ?? { ...DEFAULT_GRID_LIMITS, maxRows: 10_000, maxCols: 200 };
    this.selection = createSelection({ row: 0, col: 0 }, this.limits);
    this.currentUser = { id: "local", name: t("chat.role.user") };

    // Seed a simple outline group: rows 2-4 with a summary row at 5 (Excel 1-based indices).
    this.outline.groupRows(2, 4);
    this.outline.recomputeOutlineHiddenRows();
    // And columns 2-4 with a summary column at 5.
    this.outline.groupCols(2, 4);
    this.outline.recomputeOutlineHiddenCols();

    // Seed data for navigation tests (used range ends at D5).
    this.document.setCellValue(this.sheetId, { row: 0, col: 0 }, "Seed");
    this.document.setCellValue(this.sheetId, { row: 0, col: 1 }, {
      text: "Rich Bold",
      runs: [
        { start: 0, end: 5, style: {} },
        { start: 5, end: 9, style: { bold: true } }
      ]
    });
    this.document.setCellValue(this.sheetId, { row: 4, col: 3 }, "BottomRight");

    // Seed a small data range for the demo chart without expanding the used range past D5.
    this.document.setCellValue(this.sheetId, { row: 1, col: 0 }, "A");
    this.document.setCellValue(this.sheetId, { row: 1, col: 1 }, 2);
    this.document.setCellValue(this.sheetId, { row: 2, col: 0 }, "B");
    this.document.setCellValue(this.sheetId, { row: 2, col: 1 }, 4);
    this.document.setCellValue(this.sheetId, { row: 3, col: 0 }, "C");
    this.document.setCellValue(this.sheetId, { row: 3, col: 1 }, 3);
    this.document.setCellValue(this.sheetId, { row: 4, col: 0 }, "D");
    this.document.setCellValue(this.sheetId, { row: 4, col: 1 }, 5);

    // Best-effort: keep the WASM engine worker hydrated from the DocumentController.
    // When the WASM module isn't available (e.g. local dev without building it),
    // the app continues to operate using the in-process mock engine.
    void this.initWasmEngine();

    this.gridCanvas = document.createElement("canvas");
    this.gridCanvas.className = "grid-canvas";
    this.gridCanvas.setAttribute("aria-hidden", "true");

    this.chartLayer = document.createElement("div");
    this.chartLayer.className = "chart-layer";
    this.chartLayer.setAttribute("aria-hidden", "true");

    this.referenceCanvas = document.createElement("canvas");
    this.referenceCanvas.className = "grid-canvas";
    this.referenceCanvas.setAttribute("aria-hidden", "true");
    this.selectionCanvas = document.createElement("canvas");
    this.selectionCanvas.className = "grid-canvas";
    this.selectionCanvas.setAttribute("aria-hidden", "true");

    this.root.appendChild(this.gridCanvas);
    this.root.appendChild(this.chartLayer);
    this.root.appendChild(this.referenceCanvas);
    this.root.appendChild(this.selectionCanvas);

    this.chartStore = new ChartStore({
      defaultSheet: this.sheetId,
      getCellValue: (sheetId, row, col) => {
        const state = this.document.getCell(sheetId, { row, col }) as {
          value: unknown;
          formula: string | null;
        };
        const value = state?.value ?? null;
        return isRichTextValue(value) ? value.text : value;
      },
      onChange: () => this.renderCharts()
    });

    this.outlineLayer = document.createElement("div");
    this.outlineLayer.className = "outline-layer";
    this.root.appendChild(this.outlineLayer);

    this.commentsPanel = this.createCommentsPanel();
    this.root.appendChild(this.commentsPanel);

    this.commentTooltip = this.createCommentTooltip();
    this.root.appendChild(this.commentTooltip);

    const gridCtx = this.gridCanvas.getContext("2d");
    const referenceCtx = this.referenceCanvas.getContext("2d");
    const selectionCtx = this.selectionCanvas.getContext("2d");
    if (!gridCtx || !referenceCtx || !selectionCtx) {
      throw new Error("Canvas 2D context not available");
    }
    this.gridCtx = gridCtx;
    this.referenceCtx = referenceCtx;
    this.selectionCtx = selectionCtx;

    this.editor = new CellEditorOverlay(this.root, {
      onCommit: (commit) => {
        this.applyEdit(commit.cell, commit.value);

        const next = navigateSelectionByKey(
          this.selection,
          commit.reason === "enter" ? "Enter" : "Tab",
          { shift: commit.shift, primary: false },
          this.usedRangeProvider(),
          this.limits
        );

        if (next) this.selection = next;
        this.refresh();
        this.focus();
      },
      onCancel: () => {
        this.renderSelection();
        this.updateStatus();
        this.focus();
      }
    });

    this.root.addEventListener("pointerdown", (e) => this.onPointerDown(e));
    this.root.addEventListener("pointermove", (e) => this.onPointerMove(e));
    this.root.addEventListener("pointerup", (e) => this.onPointerUp(e));
    this.root.addEventListener("pointercancel", (e) => this.onPointerUp(e));
    this.root.addEventListener("pointerleave", () => this.hideCommentTooltip());
    this.root.addEventListener("keydown", (e) => this.onKeyDown(e));

    this.resizeObserver = new ResizeObserver(() => this.onResize());
    this.resizeObserver.observe(this.root);

    this.commentsDoc.on("update", () => {
      this.reindexCommentCells();
      this.refresh();
    });

    if (typeof window !== "undefined") {
      try {
        this.stopCommentPersistence = bindDocToStorage(this.commentsDoc, window.localStorage, "formula:comments");
      } catch {
        // Ignore persistence failures (e.g. storage disabled).
      }
    }

    if (opts.formulaBar) {
      this.formulaBar = new FormulaBarView(opts.formulaBar, {
        onBeginEdit: () => {
          this.formulaEditCell = { ...this.selection.active };
        },
        onGoTo: (reference) => this.goTo(reference),
        onCommit: (text) => this.commitFormulaBar(text),
        onCancel: () => this.cancelFormulaBar(),
        onHoverRange: (range) => {
          this.referencePreview = range
            ? {
                start: { row: range.start.row, col: range.start.col },
                end: { row: range.end.row, col: range.end.col }
              }
            : null;
          this.renderReferencePreview();
        }
      });

      this.formulaBarCompletion = new FormulaBarTabCompletionController({
        formulaBar: this.formulaBar,
        document: this.document,
        getSheetId: () => this.sheetId,
        limits: this.limits,
      });
    }

    // Seed a demo chart using the chart store helpers so it matches the logic
    // used by AI chart creation.
    this.chartStore.createChart({
      chart_type: "bar",
      data_range: "Sheet1!A2:B5",
      title: "Example Chart"
    });

    // Initial layout + render.
    this.onResize();
    this.uiReady = true;
  }

  destroy(): void {
    this.formulaBarCompletion?.destroy();
    this.wasmUnsubscribe?.();
    this.wasmUnsubscribe = null;
    this.wasmEngine?.terminate();
    this.wasmEngine = null;
    this.stopCommentPersistence?.();
    this.resizeObserver.disconnect();
    this.root.replaceChildren();
  }

  /**
   * Request a full redraw of the grid + overlays.
   *
   * This is intentionally public so external callers (e.g. AI tool execution)
   * can update the UI after mutating the DocumentController.
   */
  refresh(): void {
    if (this.renderScheduled) return;
    this.renderScheduled = true;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.renderScheduled = false;
      this.renderGrid();
      this.renderCharts();
      this.renderReferencePreview();
      this.renderSelection();
      this.updateStatus();
    });
  }

  focus(): void {
    this.root.focus();
  }

  whenIdle(): Promise<void> {
    return this.idle.whenIdle();
  }

  getRecalcCount(): number {
    return this.engine.recalcCount;
  }

  getDocument(): DocumentController {
    return this.document;
  }

  repaint(): void {
    this.renderGrid();
    this.renderCharts();
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
  }

  getCurrentSheetId(): string {
    return this.sheetId;
  }

  addChart(spec: CreateChartSpec): CreateChartResult {
    return this.chartStore.createChart(spec);
  }

  listCharts(): readonly ChartRecord[] {
    return this.chartStore.listCharts();
  }

  /**
   * Replace the DocumentController state from a snapshot, then hydrate the WASM engine in one step.
   *
   * This avoids N-per-cell RPC roundtrips during version restore by using the engine JSON load path.
   */
  async restoreDocumentState(snapshot: Uint8Array): Promise<void> {
    this.wasmSyncSuspended = true;
    try {
      this.document.applyState(snapshot);
      if (this.wasmEngine) {
        await engineHydrateFromDocument(this.wasmEngine, this.document);
      }
    } finally {
      this.wasmSyncSuspended = false;
    }
  }

  private async initWasmEngine(): Promise<void> {
    if (this.wasmEngine) return;
    if (typeof Worker === "undefined") return;

    const env = (import.meta as any)?.env as Record<string, unknown> | undefined;
    const wasmModuleUrl =
      typeof env?.VITE_FORMULA_WASM_MODULE_URL === "string" ? env.VITE_FORMULA_WASM_MODULE_URL : undefined;
    const wasmBinaryUrl =
      typeof env?.VITE_FORMULA_WASM_BINARY_URL === "string" ? env.VITE_FORMULA_WASM_BINARY_URL : undefined;

    let engine: EngineClient | null = null;
    try {
      engine = createEngineClient({ wasmModuleUrl, wasmBinaryUrl });
      await engine.init();
      await engineHydrateFromDocument(engine, this.document);

      this.wasmEngine = engine;
      this.wasmUnsubscribe = this.document.on("change", ({ deltas }: { deltas: any[] }) => {
        if (!this.wasmEngine || this.wasmSyncSuspended) return;
        void engineApplyDeltas(this.wasmEngine, deltas).catch(() => {
          // Ignore WASM sync failures; the DocumentController remains the source of truth.
        });
      });
    } catch {
      // Ignore initialization failures (e.g. missing WASM bundle).
      engine?.terminate();
      this.wasmEngine = null;
      this.wasmUnsubscribe?.();
      this.wasmUnsubscribe = null;
    }
  }

  /**
   * Switch the active sheet id and re-render.
   */
  activateSheet(sheetId: string): void {
    if (!sheetId) return;
    if (sheetId === this.sheetId) return;
    this.sheetId = sheetId;
    this.chartStore.setDefaultSheet(sheetId);
    this.renderGrid();
    this.renderCharts();
    this.renderSelection();
    this.updateStatus();
  }

  /**
   * Programmatically set the active cell (and optionally change sheets).
   */
  activateCell(target: { sheetId?: string; row: number; col: number }): void {
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.chartStore.setDefaultSheet(target.sheetId);
      this.renderGrid();
      this.renderCharts();
    }
    this.selection = setActiveCell(this.selection, { row: target.row, col: target.col }, this.limits);
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  /**
   * Programmatically set the selection range (and optionally change sheets).
   */
  selectRange(target: { sheetId?: string; range: Range }): void {
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.chartStore.setDefaultSheet(target.sheetId);
      this.renderGrid();
      this.renderCharts();
    }
    const active = { row: target.range.startRow, col: target.range.startCol };
    this.selection = buildSelection(
      { ranges: [target.range], active, anchor: active, activeRangeIndex: 0 },
      this.limits
    );
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  getSelectionRanges(): Range[] {
    return this.selection.ranges;
  }

  getActiveCell(): CellCoord {
    return { ...this.selection.active };
  }

  subscribeSelection(listener: (selection: SelectionState) => void): () => void {
    this.selectionListeners.add(listener);
    listener(this.selection);
    return () => this.selectionListeners.delete(listener);
  }

  private goTo(reference: string): void {
    try {
      const parsed = parseGoTo(reference, { workbook: this.searchWorkbook, currentSheetName: this.sheetId });
      if (parsed.type !== "range") return;

      const { range } = parsed;
      if (range.startRow === range.endRow && range.startCol === range.endCol) {
        this.activateCell({ sheetId: parsed.sheetName, row: range.startRow, col: range.startCol });
      } else {
        this.selectRange({ sheetId: parsed.sheetName, range });
      }
    } catch {
      // Ignore invalid Go To inputs for now.
    }
  }

  async getCellValueA1(a1: string): Promise<string> {
    await this.whenIdle();
    const cell = parseA1(a1);
    return this.getCellDisplayValue(cell);
  }

  toggleCommentsPanel(): void {
    this.commentsPanelVisible = !this.commentsPanelVisible;
    this.commentsPanel.style.display = this.commentsPanelVisible ? "flex" : "none";
    if (this.commentsPanelVisible) {
      this.renderCommentsPanel();
      this.newCommentInput.focus();
    } else {
      this.focus();
    }
  }

  private createCommentsPanel(): HTMLDivElement {
    const panel = document.createElement("div");
    panel.dataset.testid = "comments-panel";
    panel.style.position = "absolute";
    panel.style.top = "0";
    panel.style.right = "0";
    panel.style.width = "320px";
    panel.style.height = "100%";
    panel.style.display = "none";
    panel.style.flexDirection = "column";
    panel.style.background = "var(--dialog-bg)";
    panel.style.borderInlineStart = "1px solid var(--dialog-border)";
    panel.style.boxShadow = "-2px 0 10px var(--dialog-border)";
    panel.style.zIndex = "20";
    panel.style.padding = "10px";
    panel.style.boxSizing = "border-box";

    panel.addEventListener("pointerdown", (e) => e.stopPropagation());
    panel.addEventListener("dblclick", (e) => e.stopPropagation());
    panel.addEventListener("keydown", (e) => e.stopPropagation());

    const header = document.createElement("div");
    header.style.display = "flex";
    header.style.alignItems = "center";
    header.style.justifyContent = "space-between";
    header.style.marginBottom = "8px";

    const title = document.createElement("div");
    title.textContent = t("comments.title");
    title.style.fontWeight = "600";

    const closeButton = document.createElement("button");
    closeButton.textContent = "×";
    closeButton.setAttribute("aria-label", t("comments.closePanel"));
    closeButton.addEventListener("click", () => this.toggleCommentsPanel());

    header.appendChild(title);
    header.appendChild(closeButton);
    panel.appendChild(header);

    this.commentsPanelCell = document.createElement("div");
    this.commentsPanelCell.dataset.testid = "comments-active-cell";
    this.commentsPanelCell.style.fontSize = "12px";
    this.commentsPanelCell.style.color = "var(--text-secondary)";
    this.commentsPanelCell.style.marginBottom = "10px";
    panel.appendChild(this.commentsPanelCell);

    this.commentsPanelThreads = document.createElement("div");
    this.commentsPanelThreads.style.flex = "1";
    this.commentsPanelThreads.style.overflow = "auto";
    this.commentsPanelThreads.style.display = "flex";
    this.commentsPanelThreads.style.flexDirection = "column";
    this.commentsPanelThreads.style.gap = "10px";
    panel.appendChild(this.commentsPanelThreads);

    const footer = document.createElement("div");
    footer.style.display = "flex";
    footer.style.gap = "8px";
    footer.style.paddingTop = "10px";
    footer.style.borderTop = "1px solid var(--border)";

    this.newCommentInput = document.createElement("input");
    this.newCommentInput.dataset.testid = "new-comment-input";
    this.newCommentInput.type = "text";
    this.newCommentInput.placeholder = t("comments.new.placeholder");
    this.newCommentInput.style.flex = "1";

    const submit = document.createElement("button");
    submit.dataset.testid = "submit-comment";
    submit.textContent = t("comments.new.submit");
    submit.addEventListener("click", () => this.submitNewComment());

    footer.appendChild(this.newCommentInput);
    footer.appendChild(submit);
    panel.appendChild(footer);

    return panel;
  }

  private createCommentTooltip(): HTMLDivElement {
    const tooltip = document.createElement("div");
    tooltip.dataset.testid = "comment-tooltip";
    tooltip.style.position = "absolute";
    tooltip.style.display = "none";
    tooltip.style.maxWidth = "260px";
    tooltip.style.padding = "8px 10px";
    tooltip.style.background = "var(--bg-tertiary)";
    tooltip.style.color = "var(--text-primary)";
    tooltip.style.border = "1px solid var(--border)";
    tooltip.style.fontSize = "12px";
    tooltip.style.borderRadius = "8px";
    tooltip.style.pointerEvents = "none";
    tooltip.style.whiteSpace = "pre-wrap";
    tooltip.style.zIndex = "30";
    return tooltip;
  }

  private hideCommentTooltip(): void {
    this.commentTooltip.style.display = "none";
  }

  private reindexCommentCells(): void {
    this.commentCells.clear();
    for (const comment of this.commentManager.listAll()) {
      this.commentCells.add(comment.cellRef);
    }
  }

  private renderCommentsPanel(): void {
    if (!this.commentsPanelVisible) return;

    const cellRef = cellToA1(this.selection.active);
    this.commentsPanelCell.textContent = tWithVars("comments.cellLabel", { cellRef });

    const threads = this.commentManager.listForCell(cellRef);
    this.commentsPanelThreads.replaceChildren();

    if (threads.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = t("comments.none");
      empty.style.fontSize = "12px";
      empty.style.color = "var(--text-secondary)";
      this.commentsPanelThreads.appendChild(empty);
      return;
    }

    for (const comment of threads) {
      this.commentsPanelThreads.appendChild(this.renderCommentThread(comment));
    }
  }

  private renderCommentThread(comment: Comment): HTMLElement {
    const container = document.createElement("div");
    container.dataset.testid = "comment-thread";
    container.dataset.commentId = comment.id;
    container.dataset.resolved = comment.resolved ? "true" : "false";
    container.style.border = "1px solid var(--border)";
    container.style.borderRadius = "8px";
    container.style.padding = "10px";
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.style.gap = "8px";
    if (comment.resolved) {
      container.style.opacity = "0.7";
    }

    const header = document.createElement("div");
    header.style.display = "flex";
    header.style.alignItems = "center";
    header.style.justifyContent = "space-between";

    const author = document.createElement("div");
    author.textContent = comment.author.name || t("presence.anonymous");
    author.style.fontSize = "12px";
    author.style.fontWeight = "600";

    const resolve = document.createElement("button");
    resolve.dataset.testid = "resolve-comment";
    resolve.textContent = comment.resolved ? t("comments.unresolve") : t("comments.resolve");
    resolve.addEventListener("click", () => {
      this.commentManager.setResolved({
        commentId: comment.id,
        resolved: !comment.resolved,
      });
    });

    header.appendChild(author);
    header.appendChild(resolve);
    container.appendChild(header);

    const body = document.createElement("div");
    body.textContent = comment.content;
    body.style.fontSize = "13px";
    body.style.whiteSpace = "pre-wrap";
    container.appendChild(body);

    for (const reply of comment.replies) {
      const replyEl = document.createElement("div");
      replyEl.style.paddingInlineStart = "10px";
      replyEl.style.borderInlineStart = "2px solid var(--border)";

      const replyAuthor = document.createElement("div");
      replyAuthor.textContent = reply.author.name || t("presence.anonymous");
      replyAuthor.style.fontSize = "12px";
      replyAuthor.style.fontWeight = "600";

      const replyBody = document.createElement("div");
      replyBody.textContent = reply.content;
      replyBody.style.fontSize = "13px";
      replyBody.style.whiteSpace = "pre-wrap";
      replyBody.style.marginTop = "4px";

      replyEl.appendChild(replyAuthor);
      replyEl.appendChild(replyBody);
      container.appendChild(replyEl);
    }

    const replyRow = document.createElement("div");
    replyRow.style.display = "flex";
    replyRow.style.gap = "8px";

    const replyInput = document.createElement("input");
    replyInput.dataset.testid = "reply-input";
    replyInput.type = "text";
    replyInput.placeholder = t("comments.reply.placeholder");
    replyInput.style.flex = "1";

    const submitReply = document.createElement("button");
    submitReply.dataset.testid = "submit-reply";
    submitReply.textContent = t("comments.reply.send");
    submitReply.addEventListener("click", () => {
      const content = replyInput.value.trim();
      if (!content) return;
      this.commentManager.addReply({
        commentId: comment.id,
        content,
        author: this.currentUser,
      });
      replyInput.value = "";
    });

    replyRow.appendChild(replyInput);
    replyRow.appendChild(submitReply);
    container.appendChild(replyRow);

    return container;
  }

  private submitNewComment(): void {
    const content = this.newCommentInput.value.trim();
    if (!content) return;
    const cellRef = cellToA1(this.selection.active);

    this.commentManager.addComment({
      cellRef,
      kind: "threaded",
      content,
      author: this.currentUser,
    });

    this.newCommentInput.value = "";
  }

  private onResize(): void {
    const rect = this.root.getBoundingClientRect();
    this.width = rect.width;
    this.height = rect.height;
    this.dpr = window.devicePixelRatio || 1;

    for (const canvas of [this.gridCanvas, this.referenceCanvas, this.selectionCanvas]) {
      canvas.width = Math.floor(this.width * this.dpr);
      canvas.height = Math.floor(this.height * this.dpr);
      canvas.style.width = `${this.width}px`;
      canvas.style.height = `${this.height}px`;
    }

    // Reset transforms and apply DPR scaling so drawing code uses CSS pixels.
    for (const ctx of [this.gridCtx, this.referenceCtx, this.selectionCtx]) {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.scale(this.dpr, this.dpr);
    }

    this.renderGrid();
    this.renderCharts();
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
  }

  private renderGrid(): void {
    this.updateViewportMapping();

    const ctx = this.gridCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.gridCanvas.width, this.gridCanvas.height);
    ctx.restore();

    ctx.save();
    ctx.fillStyle = resolveCssVar("--bg-primary", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.width, this.height);

    const cols = this.visibleCols.length;
    const rows = this.visibleRows.length;

    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;

    ctx.strokeStyle = resolveCssVar("--grid-line", { fallback: "CanvasText" });
    ctx.lineWidth = 1;

    // Header backgrounds.
    ctx.fillStyle = resolveCssVar("--grid-header-bg", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.width, this.colHeaderHeight);
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.height);

    // Corner cell.
    ctx.fillStyle = resolveCssVar("--bg-tertiary", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.colHeaderHeight);

    // Grid lines for the data region.
    for (let r = 0; r <= rows; r++) {
      const y = originY + r * this.cellHeight + 0.5;
      ctx.beginPath();
      ctx.moveTo(originX, y);
      ctx.lineTo(originX + cols * this.cellWidth, y);
      ctx.stroke();
    }

    for (let c = 0; c <= cols; c++) {
      const x = originX + c * this.cellWidth + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, originY);
      ctx.lineTo(x, originY + rows * this.cellHeight);
      ctx.stroke();
    }

    // Header separator lines.
    ctx.beginPath();
    ctx.moveTo(originX + 0.5, 0);
    ctx.lineTo(originX + 0.5, this.height);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(0, originY + 0.5);
    ctx.lineTo(this.width, originY + 0.5);
    ctx.stroke();

    const fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    const fontSizePx = 14;
    const defaultTextColor = resolveCssVar("--text-primary", { fallback: "CanvasText" });

    for (let visualRow = 0; visualRow < rows; visualRow++) {
      const row = this.visibleRows[visualRow]!;
      for (let visualCol = 0; visualCol < cols; visualCol++) {
        const col = this.visibleCols[visualCol]!;
        const state = this.document.getCell(this.sheetId, { row, col }) as {
          value: unknown;
          formula: string | null;
        };
        if (!state) continue;

        const rich = isRichTextValue(state.value)
          ? state.value
          : state.formula != null
            ? { text: state.formula, runs: [] }
            : state.value != null
              ? { text: String(state.value), runs: [] }
              : null;

        if (!rich || rich.text === "") continue;

        renderRichText(
          ctx,
          rich,
          {
            x: originX + visualCol * this.cellWidth,
            y: originY + visualRow * this.cellHeight,
            width: this.cellWidth,
            height: this.cellHeight
          },
          {
            padding: 4,
            align: "start",
            verticalAlign: "middle",
            fontFamily,
            fontSizePx,
            color: defaultTextColor
          }
        );
      }
    }

    // Comment indicators.
    for (let visualRow = 0; visualRow < rows; visualRow++) {
      const row = this.visibleRows[visualRow]!;
      for (let visualCol = 0; visualCol < cols; visualCol++) {
        const col = this.visibleCols[visualCol]!;
        const cellRef = cellToA1({ row, col });
        if (!this.commentCells.has(cellRef)) continue;
        drawCommentIndicator(ctx, {
          x: originX + visualCol * this.cellWidth,
          y: originY + visualRow * this.cellHeight,
          width: this.cellWidth,
          height: this.cellHeight,
        });
      }
    }

    // Header labels.
    ctx.fillStyle = resolveCssVar("--text-primary", { fallback: "CanvasText" });
    ctx.font = "12px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    for (let visualCol = 0; visualCol < cols; visualCol++) {
      const colIndex = this.visibleCols[visualCol]!;
      ctx.fillText(
        colToName(colIndex),
        originX + visualCol * this.cellWidth + this.cellWidth / 2,
        this.colHeaderHeight / 2
      );
    }

    for (let visualRow = 0; visualRow < rows; visualRow++) {
      const rowIndex = this.visibleRows[visualRow]!;
      ctx.fillText(
        String(rowIndex + 1),
        this.rowHeaderWidth / 2,
        originY + visualRow * this.cellHeight + this.cellHeight / 2
      );
    }

    ctx.restore();

    this.renderOutlineControls();
  }

  private renderCharts(): void {
    this.chartLayer.replaceChildren();
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return;

    const provider = {
      getRange: (rangeRef: string) => {
        const parsed = parseA1Range(rangeRef);
        if (!parsed) return [];
        const sheetId = parsed.sheetName ?? this.sheetId;

        const out: unknown[][] = [];
        for (let r = parsed.startRow; r <= parsed.endRow; r += 1) {
          const row: unknown[] = [];
          for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
            const state = this.document.getCell(sheetId, { row: r, col: c }) as {
              value: unknown;
              formula: string | null;
            };
            const value = state?.value ?? null;
            row.push(isRichTextValue(value) ? value.text : value);
          }
          out.push(row);
        }
        return out;
      }
    };

    for (const chart of charts) {
      const rect = anchorToRectPx(chart.anchor, {
        defaultColWidthPx: this.cellWidth,
        defaultRowHeightPx: this.cellHeight
      });
      if (!rect) continue;

      const host = document.createElement("div");
      host.setAttribute("data-testid", "chart-object");
      host.style.position = "absolute";
      host.style.left = `${rect.left}px`;
      host.style.top = `${rect.top}px`;
      host.style.width = `${rect.width}px`;
      host.style.height = `${rect.height}px`;
      host.style.pointerEvents = "none";
      host.style.overflow = "hidden";

      host.innerHTML = renderChartSvg(chart, provider, { width: rect.width, height: rect.height });
      this.chartLayer.appendChild(host);
    }
  }

  private renderSelection(): void {
    this.selectionRenderer.render(this.selectionCtx, this.selection, {
      getCellRect: (cell) => this.getCellRect(cell)
    });

    // If scrolling/resizing happened during editing, keep the editor aligned.
    if (this.editor.isOpen()) {
      const rect = this.getCellRect(this.selection.active);
      if (rect) this.editor.reposition(rect);
    }
  }

  private updateStatus(): void {
    this.status.activeCell.textContent = cellToA1(this.selection.active);
    this.status.selectionRange.textContent =
      this.selection.ranges.length === 1 ? rangeToA1(this.selection.ranges[0]) : `${this.selection.ranges.length} ranges`;
    this.status.activeValue.textContent = this.getCellDisplayValue(this.selection.active);

    if (this.formulaBar && !this.formulaBar.isEditing()) {
      const address = cellToA1(this.selection.active);
      const input = this.getCellInputText(this.selection.active);
      const value = this.getCellComputedValue(this.selection.active);
      this.formulaBar.setActiveCell({ address, input, value });
      this.formulaBarCompletion?.update();
    }

    this.renderCommentsPanel();

    for (const listener of this.selectionListeners) {
      listener(this.selection);
    }
  }

  private syncEngineNow(): void {
    (this.engine as unknown as { syncNow?: () => void }).syncNow?.();
  }

  private isRowHidden(row: number): boolean {
    const entry = this.outline.rows.entry(row + 1);
    return isHidden(entry.hidden);
  }

  private isColHidden(col: number): boolean {
    const entry = this.outline.cols.entry(col + 1);
    return isHidden(entry.hidden);
  }

  private updateViewportMapping(): void {
    const availableWidth = Math.max(0, this.width - this.rowHeaderWidth);
    const availableHeight = Math.max(0, this.height - this.colHeaderHeight);

    const cols = Math.max(1, Math.floor(availableWidth / this.cellWidth));
    const rows = Math.max(1, Math.floor(availableHeight / this.cellHeight));

    this.visibleRows = [];
    this.visibleCols = [];
    this.rowToVisual.clear();
    this.colToVisual.clear();

    for (let r = 0; r < this.limits.maxRows && this.visibleRows.length < rows; r++) {
      if (this.isRowHidden(r)) continue;
      this.rowToVisual.set(r, this.visibleRows.length);
      this.visibleRows.push(r);
    }

    for (let c = 0; c < this.limits.maxCols && this.visibleCols.length < cols; c++) {
      if (this.isColHidden(c)) continue;
      this.colToVisual.set(c, this.visibleCols.length);
      this.visibleCols.push(c);
    }
  }

  private renderOutlineControls(): void {
    this.outlineLayer.replaceChildren();
    if (!this.outline.pr.showOutlineSymbols) return;

    const size = 14;
    const padding = 4;

    // Row group toggles live in the row header.
    for (let visualRow = 0; visualRow < this.visibleRows.length; visualRow++) {
      const rowIndex = this.visibleRows[visualRow]!;
      const summaryIndex = rowIndex + 1; // 1-based
      const entry = this.outline.rows.entry(summaryIndex);
      const details = groupDetailRange(this.outline.rows, summaryIndex, entry.level, this.outline.pr.summaryBelow);
      if (!details) continue;

      const button = document.createElement("button");
      button.className = "outline-toggle";
      button.type = "button";
      button.textContent = entry.collapsed ? "+" : "-";
      button.setAttribute("data-testid", `outline-toggle-row-${summaryIndex}`);
      button.style.left = `${padding}px`;
      button.style.top = `${this.colHeaderHeight + visualRow * this.cellHeight + (this.cellHeight - size) / 2}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;

      button.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.outline.toggleRowGroup(summaryIndex);
        this.onOutlineUpdated();
      });
      button.addEventListener("pointerdown", (e) => {
        e.stopPropagation();
      });

      this.outlineLayer.appendChild(button);
    }

    // Column group toggles live in the column header.
    for (let visualCol = 0; visualCol < this.visibleCols.length; visualCol++) {
      const colIndex = this.visibleCols[visualCol]!;
      const summaryIndex = colIndex + 1; // 1-based
      const entry = this.outline.cols.entry(summaryIndex);
      const details = groupDetailRange(this.outline.cols, summaryIndex, entry.level, this.outline.pr.summaryRight);
      if (!details) continue;

      const button = document.createElement("button");
      button.className = "outline-toggle";
      button.type = "button";
      button.textContent = entry.collapsed ? "+" : "-";
      button.setAttribute("data-testid", `outline-toggle-col-${summaryIndex}`);
      button.style.left = `${this.rowHeaderWidth + visualCol * this.cellWidth + (this.cellWidth - size) / 2}px`;
      button.style.top = `${padding}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;

      button.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.outline.toggleColGroup(summaryIndex);
        this.onOutlineUpdated();
      });
      button.addEventListener("pointerdown", (e) => {
        e.stopPropagation();
      });

      this.outlineLayer.appendChild(button);
    }
  }

  private onOutlineUpdated(): void {
    this.ensureActiveCellVisible();
    this.renderGrid();
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private ensureActiveCellVisible(): void {
    let { row, col } = this.selection.active;

    if (this.isRowHidden(row)) {
      row = this.findNextVisibleRow(row, 1) ?? this.findNextVisibleRow(row, -1) ?? row;
    }
    if (this.isColHidden(col)) {
      col = this.findNextVisibleCol(col, 1) ?? this.findNextVisibleCol(col, -1) ?? col;
    }

    if (row !== this.selection.active.row || col !== this.selection.active.col) {
      this.selection = setActiveCell(this.selection, { row, col }, this.limits);
    }
  }

  private findNextVisibleRow(start: number, dir: 1 | -1): number | null {
    let row = start + dir;
    while (row >= 0 && row < this.limits.maxRows) {
      if (!this.isRowHidden(row)) return row;
      row += dir;
    }
    return null;
  }

  private findNextVisibleCol(start: number, dir: 1 | -1): number | null {
    let col = start + dir;
    while (col >= 0 && col < this.limits.maxCols) {
      if (!this.isColHidden(col)) return col;
      col += dir;
    }
    return null;
  }

  private getCellRect(cell: CellCoord) {
    const visualRow = this.rowToVisual.get(cell.row);
    const visualCol = this.colToVisual.get(cell.col);
    if (visualRow === undefined || visualCol === undefined) return null;

    return {
      x: this.rowHeaderWidth + visualCol * this.cellWidth,
      y: this.colHeaderHeight + visualRow * this.cellHeight,
      width: this.cellWidth,
      height: this.cellHeight
    };
  }

  private cellFromPoint(pointX: number, pointY: number): CellCoord {
    const colVisual = Math.floor((pointX - this.rowHeaderWidth) / this.cellWidth);
    const rowVisual = Math.floor((pointY - this.colHeaderHeight) / this.cellHeight);
    const col = this.visibleCols[Math.max(0, Math.min(this.visibleCols.length - 1, colVisual))] ?? 0;
    const row = this.visibleRows[Math.max(0, Math.min(this.visibleRows.length - 1, rowVisual))] ?? 0;
    return {
      row: Math.max(0, Math.min(this.limits.maxRows - 1, row)),
      col: Math.max(0, Math.min(this.limits.maxCols - 1, col))
    };
  }

  private onPointerDown(e: PointerEvent): void {
    if (this.editor.isOpen()) return;

    const rect = this.root.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    const primary = e.ctrlKey || e.metaKey;

    // Top-left corner selects the entire sheet.
    if (x < this.rowHeaderWidth && y < this.colHeaderHeight) {
      this.selection = selectAll(this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    // Column header selects entire column.
    if (y < this.colHeaderHeight && x >= this.rowHeaderWidth) {
      const visualCol = Math.floor((x - this.rowHeaderWidth) / this.cellWidth);
      const col = this.visibleCols[Math.max(0, Math.min(this.visibleCols.length - 1, visualCol))] ?? 0;
      this.selection = selectColumns(this.selection, col, col, { additive: primary }, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    // Row header selects entire row.
    if (x < this.rowHeaderWidth && y >= this.colHeaderHeight) {
      const visualRow = Math.floor((y - this.colHeaderHeight) / this.cellHeight);
      const row = this.visibleRows[Math.max(0, Math.min(this.visibleRows.length - 1, visualRow))] ?? 0;
      this.selection = selectRows(this.selection, row, row, { additive: primary }, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    const cell = this.cellFromPoint(x, y);

    if (this.formulaBar?.isFormulaEditing()) {
      e.preventDefault();
      this.dragState = { pointerId: e.pointerId, mode: "formula" };
      this.root.setPointerCapture(e.pointerId);
      this.selection = setActiveCell(this.selection, cell, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.formulaBar.beginRangeSelection({
        start: { row: cell.row, col: cell.col },
        end: { row: cell.row, col: cell.col }
      });
      return;
    }

    this.dragState = { pointerId: e.pointerId, mode: "normal" };
    this.root.setPointerCapture(e.pointerId);
    if (e.shiftKey) {
      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
    } else if (primary) {
      this.selection = addCellToSelection(this.selection, cell, this.limits);
    } else {
      this.selection = setActiveCell(this.selection, cell, this.limits);
    }

    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private onPointerMove(e: PointerEvent): void {
    const rect = this.root.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    if (this.dragState) {
      if (e.pointerId !== this.dragState.pointerId) return;
      if (this.editor.isOpen()) return;
      this.hideCommentTooltip();
      const cell = this.cellFromPoint(x, y);
      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
      this.renderSelection();
      this.updateStatus();

      if (this.dragState.mode === "formula" && this.formulaBar) {
        const r = this.selection.ranges[0];
        this.formulaBar.updateRangeSelection({
          start: { row: r.startRow, col: r.startCol },
          end: { row: r.endRow, col: r.endCol }
        });
      }
      return;
    }

    if (this.commentsPanelVisible) {
      // Don't show tooltips while the panel is open; it obscures the grid anyway.
      this.hideCommentTooltip();
      return;
    }

    if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
      this.hideCommentTooltip();
      return;
    }

    const cell = this.cellFromPoint(x, y);
    const cellRef = cellToA1(cell);
    if (!this.commentCells.has(cellRef)) {
      this.hideCommentTooltip();
      return;
    }

    const comments = this.commentManager.listForCell(cellRef);
    const preview = comments[0]?.content ?? "";
    if (!preview) {
      this.hideCommentTooltip();
      return;
    }

    this.commentTooltip.textContent = preview;
    this.commentTooltip.style.left = `${x + 12}px`;
    this.commentTooltip.style.top = `${y + 12}px`;
    this.commentTooltip.style.display = "block";
  }

  private onPointerUp(e: PointerEvent): void {
    if (!this.dragState) return;
    if (e.pointerId !== this.dragState.pointerId) return;
    const mode = this.dragState.mode;
    this.dragState = null;

    if (mode === "formula" && this.formulaBar) {
      this.formulaBar.endRangeSelection();
      // Restore focus to the formula bar without clearing its insertion state mid-drag.
      this.formulaBar.focus();
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (this.editor.isOpen()) {
      // The editor handles Enter/Tab/Escape itself. We keep focus on the textarea.
      return;
    }

    if (isUndoKeyboardEvent(e)) {
      if (this.formulaBar?.isEditing()) return;
      e.preventDefault();
      if (this.document.undo()) {
        this.syncEngineNow();
        this.refresh();
      }
      return;
    }

    if (isRedoKeyboardEvent(e)) {
      if (this.formulaBar?.isEditing()) return;
      e.preventDefault();
      if (this.document.redo()) {
        this.syncEngineNow();
        this.refresh();
      }
      return;
    }

    // Editing
    if (e.key === "F2") {
      e.preventDefault();
      const cell = this.selection.active;
      const bounds = this.getCellRect(cell);
      if (!bounds) return;
      const initialValue = this.getCellInputText(cell);
      this.editor.open(cell, bounds, initialValue, { cursor: "end" });
      return;
    }

    const primary = e.ctrlKey || e.metaKey;
    if (e.key === "Delete") {
      e.preventDefault();
      this.clearSelectionContents();
      this.refresh();
      return;
    }

    // Ctrl/Cmd+Shift+M toggles the comments panel.
    if (primary && e.shiftKey && (e.key === "m" || e.key === "M")) {
      e.preventDefault();
      this.toggleCommentsPanel();
      return;
    }

    // Selection shortcuts
    if (primary && (e.key === "a" || e.key === "A")) {
      e.preventDefault();
      this.selection = selectAll(this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (primary && e.code === "Space") {
      // Ctrl+Space selects entire column.
      e.preventDefault();
      this.selection = selectColumns(this.selection, this.selection.active.col, this.selection.active.col, {}, this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (!primary && e.shiftKey && e.code === "Space") {
      // Shift+Space selects entire row.
      e.preventDefault();
      this.selection = selectRows(this.selection, this.selection.active.row, this.selection.active.row, {}, this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    // Outline grouping shortcuts (Excel-style): Alt+Shift+Right/Left.
    if (e.altKey && e.shiftKey && (e.key === "ArrowRight" || e.key === "ArrowLeft")) {
      e.preventDefault();
      const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      if (!range) return;

      const startRow = range.startRow + 1;
      const endRow = range.endRow + 1;
      const startCol = range.startCol + 1;
      const endCol = range.endCol + 1;

      if (e.key === "ArrowRight") {
        if (this.selection.type === "column") {
          this.outline.groupCols(startCol, endCol);
          this.outline.recomputeOutlineHiddenCols();
        } else {
          this.outline.groupRows(startRow, endRow);
          this.outline.recomputeOutlineHiddenRows();
        }
      } else {
        if (this.selection.type === "column") {
          this.outline.ungroupCols(startCol, endCol);
        } else {
          this.outline.ungroupRows(startRow, endRow);
        }
      }

      this.onOutlineUpdated();
      return;
    }

    // Excel-like "start typing to edit" behavior: any printable key begins edit
    // mode and replaces the cell contents.
    if (!primary && !e.altKey && e.key.length === 1) {
      e.preventDefault();
      const cell = this.selection.active;
      const bounds = this.getCellRect(cell);
      if (!bounds) return;
      this.editor.open(cell, bounds, e.key, { cursor: "end" });
      return;
    }

    const next = navigateSelectionByKey(
      this.selection,
      e.key,
      { shift: e.shiftKey, primary },
      this.usedRangeProvider(),
      this.limits
    );
    if (!next) return;

    e.preventDefault();
    this.selection = next;
    this.renderSelection();
    this.updateStatus();
  }

  private getCellDisplayValue(cell: CellCoord): string {
    const value = this.getCellComputedValue(cell);
    if (value == null) return "";
    return String(value);
  }

  private getCellInputText(cell: CellCoord): string {
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (state?.formula != null) {
      return state.formula;
    }
    if (isRichTextValue(state?.value)) return state.value.text;
    if (state?.value != null) return String(state.value);
    return "";
  }

  private getCellComputedValue(cell: CellCoord): SpreadsheetValue {
    const cacheKey = this.computedKey(this.sheetId, cellToA1(cell));
    if (this.computedValues.has(cacheKey)) {
      return this.computedValues.get(cacheKey) ?? null;
    }

    const memo = new Map<string, SpreadsheetValue>();
    const stack = new Set<string>();
    return this.computeCellValue(cell, memo, stack);
  }

  private computedKey(sheetId: string, address: string): string {
    return `${sheetId}:${address.replaceAll("$", "").toUpperCase()}`;
  }

  private invalidateComputedValues(changes: unknown): void {
    if (!Array.isArray(changes)) return;
    for (const change of changes) {
      const ref = change as EngineCellRef;
      const sheetId = typeof ref.sheetId === "string" ? ref.sheetId : this.sheetId;
      if (!Number.isInteger(ref.row) || !Number.isInteger(ref.col)) continue;
      const address = cellToA1({ row: ref.row, col: ref.col });
      this.computedValues.delete(this.computedKey(sheetId, address));
    }
  }

  private applyComputedChanges(changes: unknown): void {
    if (!Array.isArray(changes)) return;
    let updated = false;

    for (const change of changes) {
      const ref = change as EngineCellRef;

      let sheetId = typeof ref.sheet === "string" ? ref.sheet : undefined;
      if (!sheetId && typeof ref.sheetId === "string") sheetId = ref.sheetId;
      if (!sheetId) sheetId = this.sheetId;

      let address = typeof ref.address === "string" ? ref.address : undefined;
      if (!address && Number.isInteger(ref.row) && Number.isInteger(ref.col)) {
        address = cellToA1({ row: ref.row, col: ref.col });
      }
      if (!address) continue;

      // Support "Sheet1!A1" style addresses if a sheet name was embedded.
      if (address.includes("!")) {
        const [maybeSheet, cell] = address.split("!", 2);
        if (maybeSheet && cell) {
          sheetId = maybeSheet;
          address = cell;
        }
      }

      const value = ref.value;
      if (value !== null && typeof value !== "number" && typeof value !== "string" && typeof value !== "boolean") {
        continue;
      }

      this.computedValues.set(this.computedKey(sheetId, address), value);
      updated = true;
    }

    if (updated) {
      // Keep the status/formula bar in sync once computed values arrive.
      if (this.uiReady) this.updateStatus();
    }
  }

  private computeCellValue(
    cell: CellCoord,
    memo: Map<string, SpreadsheetValue>,
    stack: Set<string>
  ): SpreadsheetValue {
    const key = cellToA1(cell);
    const cached = memo.get(key);
    if (cached !== undefined || memo.has(key)) return cached ?? null;
    if (stack.has(key)) return "#REF!";

    stack.add(key);
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    let value: SpreadsheetValue;

    if (state?.formula != null) {
      value = evaluateFormula(state.formula, (ref) => {
        const normalized = ref.replaceAll("$", "");
        const coord = parseA1(normalized);
        return this.computeCellValue(coord, memo, stack);
      });
    } else if (state?.value != null) {
      value = state.value as SpreadsheetValue;
    } else {
      value = null;
    }

    stack.delete(key);
    memo.set(key, value);
    return value;
  }

  private applyEdit(cell: CellCoord, rawValue: string): void {
    const original = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (rawValue.trim() === "") {
      this.document.clearCell(this.sheetId, cell, { label: "Clear cell" });
      return;
    }

    if (rawValue.startsWith("=")) {
      this.document.setCellFormula(this.sheetId, cell, rawValue.slice(1), { label: "Edit cell" });
      return;
    }

    if (isRichTextValue(original?.value)) {
      const updated = applyPlainTextEdit(original.value, rawValue);
      if (original.formula == null && updated === original.value) {
        // No-op edit: keep rich runs without creating a history entry.
        return;
      }
      this.document.setCellValue(this.sheetId, cell, updated, { label: "Edit cell" });
      return;
    }

    this.document.setCellValue(this.sheetId, cell, rawValue, { label: "Edit cell" });
  }

  private commitFormulaBar(text: string): void {
    const target = this.formulaEditCell ?? this.selection.active;
    this.applyEdit(target, text);

    this.selection = setActiveCell(this.selection, target, this.limits);
    this.formulaEditCell = null;
    this.referencePreview = null;
    this.refresh();
    this.focus();
  }

  private cancelFormulaBar(): void {
    if (this.formulaEditCell) {
      this.selection = setActiveCell(this.selection, this.formulaEditCell, this.limits);
    }
    this.formulaEditCell = null;
    this.referencePreview = null;
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private renderReferencePreview(): void {
    const ctx = this.referenceCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.referenceCanvas.width, this.referenceCanvas.height);
    ctx.restore();

    if (!this.referencePreview) return;

    const startRow = Math.min(this.referencePreview.start.row, this.referencePreview.end.row);
    const endRow = Math.max(this.referencePreview.start.row, this.referencePreview.end.row);
    const startCol = Math.min(this.referencePreview.start.col, this.referencePreview.end.col);
    const endCol = Math.max(this.referencePreview.start.col, this.referencePreview.end.col);

    const startRect = this.getCellRect({ row: startRow, col: startCol });
    const endRect = this.getCellRect({ row: endRow, col: endCol });

    const x = startRect.x;
    const y = startRect.y;
    const width = endRect.x + endRect.width - startRect.x;
    const height = endRect.y + endRect.height - startRect.y;

    ctx.save();
    ctx.strokeStyle = resolveCssVar("--warning", { fallback: "CanvasText" });
    ctx.lineWidth = 2;
    ctx.setLineDash([4, 3]);
    ctx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);
    ctx.restore();
  }

  private usedRangeProvider() {
    return {
      getUsedRange: () => this.computeUsedRange(),
      isCellEmpty: (cell: CellCoord) => {
        const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
        return state?.value == null && state?.formula == null;
      },
      isRowHidden: (row: number) => this.isRowHidden(row),
      isColHidden: (col: number) => this.isColHidden(col),
    };
  }

  private computeUsedRange(): Range | null {
    return this.document.getUsedRange(this.sheetId);
  }

  private clearSelectionContents(): void {
    const used = this.computeUsedRange();
    if (!used) return;
    for (const range of this.selection.ranges) {
      const clipped = intersectRanges(range, used);
      if (!clipped) continue;
      this.document.clearRange(
        this.sheetId,
        {
          start: { row: clipped.startRow, col: clipped.startCol },
          end: { row: clipped.endRow, col: clipped.endCol }
        },
        { label: "Clear contents" }
      );
    }
  }
}

function isRichTextValue(
  value: unknown
): value is { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> } {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function intersectRanges(a: Range, b: Range): Range | null {
  const startRow = Math.max(a.startRow, b.startRow);
  const endRow = Math.min(a.endRow, b.endRow);
  const startCol = Math.max(a.startCol, b.startCol);
  const endCol = Math.min(a.endCol, b.endCol);
  if (startRow > endRow || startCol > endCol) return null;
  return { startRow, endRow, startCol, endCol };
}

function colToName(col: number): string {
  if (!Number.isFinite(col) || col < 0) return "";
  return colToNameA1(col);
}

function parseA1(a1: string): CellCoord {
  try {
    const { row0, col0 } = fromA1A1(a1);
    return { row: row0, col: col0 };
  } catch {
    return { row: 0, col: 0 };
  }
}
