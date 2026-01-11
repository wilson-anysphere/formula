import { CellEditorOverlay } from "../editor/cellEditorOverlay";
import { FormulaBarTabCompletionController } from "../ai/completion/formulaBarTabCompletion.js";
import { FormulaBarView } from "../formula-bar/FormulaBarView";
import { Outline, groupDetailRange, isHidden } from "../grid/outline/outline.js";
import { parseA1Range } from "../charts/a1.js";
import { emuToPx } from "../charts/overlay.js";
import { renderChartSvg } from "../charts/renderSvg.js";
import { ChartStore, type ChartRecord } from "../charts/chartStore";
import { FALLBACK_CHART_THEME, type ChartTheme } from "../charts/theme";
import { applyPlainTextEdit } from "../grid/text/rich-text/edit.js";
import { renderRichText } from "../grid/text/rich-text/render.js";
import {
  copyRangeToClipboardPayload,
  createClipboardProvider,
  parseClipboardContentToCellGrid,
} from "../clipboard/index.js";
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
import { createDesktopDlpContext } from "../dlp/desktopDlp.js";
import {
  createEngineClient,
  engineApplyDeltas,
  engineHydrateFromDocument,
  type EngineClient
} from "@formula/engine";
import { drawCommentIndicator } from "../comments/CommentIndicator";
import { evaluateFormula, type SpreadsheetValue } from "../spreadsheet/evaluateFormula";
import { AiCellFunctionEngine } from "../spreadsheet/AiCellFunctionEngine.js";
import { DocumentWorkbookAdapter } from "../search/documentWorkbookAdapter.js";
import { parseGoTo } from "../../../../packages/search/index.js";
import type { CreateChartResult, CreateChartSpec } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import { colToName as colToNameA1, fromA1 as fromA1A1 } from "@formula/spreadsheet-frontend/a1";
import { shiftA1References } from "@formula/spreadsheet-frontend";
import { InlineEditController, type InlineEditLLMClient } from "../ai/inline-edit/inlineEditController";
import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";

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
  syncNow?: () => unknown;
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

  syncNow(): void {
    const result = this.inner.syncNow?.();
    if (isThenable(result)) {
      const tracked = Promise.resolve(result).then((changes) => this.onComputedChanges(changes));
      this.idle.track(tracked);
      return;
    }

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

type DragState =
  | { pointerId: number; mode: "normal" }
  | { pointerId: number; mode: "formula" }
  | { pointerId: number; mode: "fill"; sourceRange: Range; targetRange: Range; endCell: CellCoord };

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
  private readonly aiCellFunctions: AiCellFunctionEngine;
  private limits: GridLimits;

  private wasmEngine: EngineClient | null = null;
  private wasmSyncSuspended = false;
  private wasmUnsubscribe: (() => void) | null = null;
  private wasmSyncPromise: Promise<void> = Promise.resolve();

  private gridCanvas: HTMLCanvasElement;
  private chartLayer: HTMLDivElement;
  private referenceCanvas: HTMLCanvasElement;
  private selectionCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private referenceCtx: CanvasRenderingContext2D;
  private selectionCtx: CanvasRenderingContext2D;

  private outline = new Outline();
  private outlineLayer: HTMLDivElement;
  private readonly outlineButtons = new Map<string, HTMLButtonElement>();

  private dpr = 1;
  private width = 0;
  private height = 0;

  // Scroll offsets in CSS pixels relative to the sheet data origin (A1 at 0,0).
  private scrollX = 0;
  private scrollY = 0;

  private readonly cellWidth = 100;
  private readonly cellHeight = 24;
  private readonly rowHeaderWidth = 48;
  private readonly colHeaderHeight = 24;

  // Precomputed "visual" (i.e. not hidden) row/col indices.
  // Rebuilt only when outline visibility changes.
  private rowIndexByVisual: number[] = [];
  private colIndexByVisual: number[] = [];

  // Current viewport window (virtualized render range).
  private visibleRows: number[] = [];
  private visibleCols: number[] = [];
  private visibleRowStart = 0;
  private visibleColStart = 0;

  private viewportMappingState:
    | {
        scrollX: number;
        scrollY: number;
        viewportWidth: number;
        viewportHeight: number;
        rowCount: number;
        colCount: number;
      }
    | null = null;

  // Maps from sheet row/col -> visual index (excluding hidden rows/cols).
  // These allow O(1) cell->pixel conversions for any in-bounds cell.
  private rowToVisual = new Map<number, number>();
  private colToVisual = new Map<number, number>();

  private readonly scrollbarThickness = 10;
  private vScrollbarTrack: HTMLDivElement;
  private vScrollbarThumb: HTMLDivElement;
  private hScrollbarTrack: HTMLDivElement;
  private hScrollbarThumb: HTMLDivElement;
  private scrollbarDrag:
    | { axis: "x" | "y"; pointerId: number; grabOffset: number; thumbTravel: number; trackStart: number; maxScroll: number }
    | null = null;

  private selection: SelectionState;
  private selectionRenderer = new SelectionRenderer();
  private readonly selectionListeners = new Set<(selection: SelectionState) => void>();

  private editor: CellEditorOverlay;
  private formulaBar: FormulaBarView | null = null;
  private formulaBarCompletion: FormulaBarTabCompletionController | null = null;
  private formulaEditCell: CellCoord | null = null;
  private referencePreview: { start: CellCoord; end: CellCoord } | null = null;
  private referenceHighlights: Array<{ start: CellCoord; end: CellCoord; color: string }> = [];
  private fillPreviewRange: Range | null = null;
  private showFormulas = false;

  private dragState: DragState | null = null;
  private dragPointerPos: { x: number; y: number } | null = null;
  private dragAutoScrollRaf: number | null = null;

  private resizeObserver: ResizeObserver;
  private disposed = false;
  private readonly domAbort = new AbortController();
  private commentsDocUpdateListener: (() => void) | null = null;

  private readonly inlineEditController: InlineEditController;

  private readonly currentUser: CommentAuthor;
  private readonly commentsDoc = new Y.Doc();
  private readonly commentManager = new CommentManager(this.commentsDoc);
  private commentCells = new Set<string>();
  private commentsPanelVisible = false;
  private stopCommentPersistence: (() => void) | null = null;

  private readonly chartStore: ChartStore;
  private chartTheme: ChartTheme = FALLBACK_CHART_THEME;

  private commentsPanel!: HTMLDivElement;
  private commentsPanelThreads!: HTMLDivElement;
  private commentsPanelCell!: HTMLDivElement;
  private newCommentInput!: HTMLInputElement;
  private commentTooltip!: HTMLDivElement;

  private renderScheduled = false;
  private pendingRenderMode: "full" | "scroll" = "full";
  private windowKeyDownListener: ((e: KeyboardEvent) => void) | null = null;
  private clipboardProviderPromise: ReturnType<typeof createClipboardProvider> | null = null;
  private clipboardCopyContext:
    | {
        range: Range;
        payload: { text?: string; html?: string };
        cells: Array<Array<{ value: unknown; formula: string | null; styleId: number }>>;
      }
    | null = null;
  private dlpContext: ReturnType<typeof createDesktopDlpContext> | null = null;

  private readonly chartElements = new Map<string, HTMLDivElement>();

  constructor(
    private root: HTMLElement,
    private status: SpreadsheetAppStatusElements,
    opts: {
      workbookId?: string;
      limits?: GridLimits;
      formulaBar?: HTMLElement;
      inlineEdit?: {
        llmClient?: InlineEditLLMClient;
        model?: string;
        auditStore?: AIAuditStore;
      };
    } = {}
  ) {
    this.limits = opts.limits ?? { ...DEFAULT_GRID_LIMITS, maxRows: 10_000, maxCols: 200 };
    this.selection = createSelection({ row: 0, col: 0 }, this.limits);
    this.currentUser = { id: "local", name: t("chat.role.user") };

    // Prevent DOM overlays (charts, scrollbars, outline buttons) from spilling
    // outside the grid viewport while we virtualize with negative coordinates.
    this.root.style.overflow = "hidden";

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
      onChange: () => this.renderCharts(true)
    });

    this.outlineLayer = document.createElement("div");
    this.outlineLayer.className = "outline-layer";
    this.root.appendChild(this.outlineLayer);

    // Minimal scrollbars (drawn as DOM overlays, like the React CanvasGrid).
    this.vScrollbarTrack = document.createElement("div");
    this.vScrollbarTrack.setAttribute("aria-hidden", "true");
    this.vScrollbarTrack.setAttribute("data-testid", "scrollbar-track-y");
    this.vScrollbarTrack.style.position = "absolute";
    this.vScrollbarTrack.style.background = "var(--bg-tertiary)";
    this.vScrollbarTrack.style.borderRadius = "6px";
    this.vScrollbarTrack.style.zIndex = "5";
    this.vScrollbarTrack.style.opacity = "0.9";

    this.vScrollbarThumb = document.createElement("div");
    this.vScrollbarThumb.setAttribute("aria-hidden", "true");
    this.vScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-y");
    this.vScrollbarThumb.style.position = "absolute";
    this.vScrollbarThumb.style.left = "1px";
    this.vScrollbarThumb.style.right = "1px";
    this.vScrollbarThumb.style.top = "0";
    this.vScrollbarThumb.style.height = "40px";
    this.vScrollbarThumb.style.background = "var(--text-secondary)";
    this.vScrollbarThumb.style.borderRadius = "6px";
    this.vScrollbarThumb.style.cursor = "pointer";
    this.vScrollbarTrack.appendChild(this.vScrollbarThumb);
    this.root.appendChild(this.vScrollbarTrack);

    this.hScrollbarTrack = document.createElement("div");
    this.hScrollbarTrack.setAttribute("aria-hidden", "true");
    this.hScrollbarTrack.setAttribute("data-testid", "scrollbar-track-x");
    this.hScrollbarTrack.style.position = "absolute";
    this.hScrollbarTrack.style.background = "var(--bg-tertiary)";
    this.hScrollbarTrack.style.borderRadius = "6px";
    this.hScrollbarTrack.style.zIndex = "5";
    this.hScrollbarTrack.style.opacity = "0.9";

    this.hScrollbarThumb = document.createElement("div");
    this.hScrollbarThumb.setAttribute("aria-hidden", "true");
    this.hScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-x");
    this.hScrollbarThumb.style.position = "absolute";
    this.hScrollbarThumb.style.top = "1px";
    this.hScrollbarThumb.style.bottom = "1px";
    this.hScrollbarThumb.style.left = "0";
    this.hScrollbarThumb.style.width = "40px";
    this.hScrollbarThumb.style.background = "var(--text-secondary)";
    this.hScrollbarThumb.style.borderRadius = "6px";
    this.hScrollbarThumb.style.cursor = "pointer";
    this.hScrollbarTrack.appendChild(this.hScrollbarThumb);
    this.root.appendChild(this.hScrollbarTrack);

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
        this.ensureActiveCellVisible();
        this.scrollCellIntoView(this.selection.active);
        this.refresh();
        this.focus();
      },
      onCancel: () => {
        this.renderSelection();
        this.updateStatus();
        this.focus();
      }
    });

    this.inlineEditController = new InlineEditController({
      container: this.root,
      document: this.document,
      workbookId: opts.workbookId,
      getSheetId: () => this.sheetId,
      getSelectionRange: () => this.getInlineEditSelectionRange(),
      onApplied: () => {
        this.renderGrid();
        this.renderCharts(true);
        this.renderSelection();
        this.updateStatus();
        this.focus();
      },
      onClosed: () => {
        this.focus();
      },
      llmClient: opts.inlineEdit?.llmClient,
      model: opts.inlineEdit?.model,
      auditStore: opts.inlineEdit?.auditStore
    });

    this.root.addEventListener("pointerdown", (e) => this.onPointerDown(e), { signal: this.domAbort.signal });
    this.root.addEventListener("pointermove", (e) => this.onPointerMove(e), { signal: this.domAbort.signal });
    this.root.addEventListener("pointerup", (e) => this.onPointerUp(e), { signal: this.domAbort.signal });
    this.root.addEventListener("pointercancel", (e) => this.onPointerUp(e), { signal: this.domAbort.signal });
    this.root.addEventListener("pointerleave", () => {
      this.hideCommentTooltip();
      this.root.style.cursor = "";
    }, { signal: this.domAbort.signal });
    this.root.addEventListener("keydown", (e) => this.onKeyDown(e), { signal: this.domAbort.signal });
    this.root.addEventListener("wheel", (e) => this.onWheel(e), { passive: false, signal: this.domAbort.signal });

    // If the user copies/cuts from an input/contenteditable (formula bar, comments, etc),
    // the system clipboard content changes and any prior "internal copy" context used for
    // style/formula shifting should be discarded. Grid copy/cut uses `preventDefault()`
    // and the Clipboard API, so it should not trigger these native events.
    if (typeof document !== "undefined") {
      const clearClipboardContext = () => {
        this.clipboardCopyContext = null;
      };
      document.addEventListener("copy", clearClipboardContext, { capture: true, signal: this.domAbort.signal });
      document.addEventListener("cut", clearClipboardContext, { capture: true, signal: this.domAbort.signal });

      if (typeof window !== "undefined") {
        window.addEventListener("blur", clearClipboardContext, { signal: this.domAbort.signal });
      }
      document.addEventListener(
        "visibilitychange",
        () => {
          if ((document as any).hidden) clearClipboardContext();
        },
        { signal: this.domAbort.signal }
      );
    }

    this.vScrollbarThumb.addEventListener("pointerdown", (e) => this.onScrollbarThumbPointerDown(e, "y"), {
      passive: false,
      signal: this.domAbort.signal
    });
    this.hScrollbarThumb.addEventListener("pointerdown", (e) => this.onScrollbarThumbPointerDown(e, "x"), {
      passive: false,
      signal: this.domAbort.signal
    });
    this.vScrollbarTrack.addEventListener("pointerdown", (e) => this.onScrollbarTrackPointerDown(e, "y"), {
      passive: false,
      signal: this.domAbort.signal
    });
    this.hScrollbarTrack.addEventListener("pointerdown", (e) => this.onScrollbarTrackPointerDown(e, "x"), {
      passive: false,
      signal: this.domAbort.signal
    });

    if (typeof window !== "undefined") {
      this.windowKeyDownListener = (e) => this.onWindowKeyDown(e);
      window.addEventListener("keydown", this.windowKeyDownListener);
    }

    this.resizeObserver = new ResizeObserver(() => this.onResize());
    this.resizeObserver.observe(this.root);

    // Save so we can detach cleanly in `destroy()`.
    this.commentsDocUpdateListener = () => {
      this.reindexCommentCells();
      this.refresh();
    };
    this.commentsDoc.on("update", this.commentsDocUpdateListener);

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
        },
        onReferenceHighlights: (highlights) => {
          const sheetIds = this.document.getSheetIds();
          const resolveSheetId = (name: string): string | null => {
            const trimmed = name.trim();
            if (!trimmed) return null;
            return sheetIds.find((id) => id.toLowerCase() === trimmed.toLowerCase()) ?? null;
          };

          this.referenceHighlights = highlights
            .filter((h) => {
              if (!h.range.sheet) return true;
              const resolved = resolveSheetId(h.range.sheet);
              if (!resolved) return false;
              return resolved.toLowerCase() === this.sheetId.toLowerCase();
            })
            .map((h) => ({
              start: { row: h.range.startRow, col: h.range.startCol },
              end: { row: h.range.endRow, col: h.range.endCol },
              color: h.color
            }));
          this.renderReferencePreview();
        }
      });

      this.formulaBarCompletion = new FormulaBarTabCompletionController({
        formulaBar: this.formulaBar,
        document: this.document,
        getSheetId: () => this.sheetId,
        limits: this.limits,
        schemaProvider: {
          getNamedRanges: () => {
            const formatSheetPrefix = (id: string): string => {
              const needsQuotes = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(id);
              if (!needsQuotes) return `${id}!`;
              return `'${id.replaceAll("'", "''")}'!`;
            };

            return Array.from(this.searchWorkbook.names.values())
              .map((entry: any) => {
                const name = typeof entry?.name === "string" ? entry.name : "";
                if (!name) return null;
                const sheetName = typeof entry?.sheetName === "string" ? entry.sheetName : "";
                const range = entry?.range;
                const rangeText =
                  sheetName && range ? `${formatSheetPrefix(sheetName)}${rangeToA1(range)}` : undefined;
                return { name, range: rangeText };
              })
              .filter((entry: { name: string; range?: string } | null): entry is { name: string; range?: string } =>
                Boolean(entry?.name),
              );
           },
           getTables: () =>
             Array.from(this.searchWorkbook.tables.values())
              .map((table: any) => ({
                name: typeof table?.name === "string" ? table.name : "",
                sheetName: typeof table?.sheetName === "string" ? table.sheetName : undefined,
                startRow: typeof table?.startRow === "number" ? table.startRow : undefined,
                startCol: typeof table?.startCol === "number" ? table.startCol : undefined,
                endRow: typeof table?.endRow === "number" ? table.endRow : undefined,
                endCol: typeof table?.endCol === "number" ? table.endCol : undefined,
                columns: Array.isArray(table?.columns) ? table.columns.map((c: unknown) => String(c)) : [],
              }))
              .filter((t: { name: string; columns: string[] }) => t.name.length > 0 && t.columns.length > 0),
           getCacheKey: () => `schema:${Number((this.searchWorkbook as any).schemaVersion) || 0}`,
         },
       });
     }

    // Precompute row/col visibility + mappings before any initial render work.
    this.rebuildAxisVisibilityCache();

    // Seed a demo chart using the chart store helpers so it matches the logic
    // used by AI chart creation.
    this.chartStore.createChart({
      chart_type: "bar",
      data_range: "Sheet1!A2:B5",
      title: "Example Chart"
    });

    const workbookId = opts.workbookId ?? "local-workbook";
    const dlp = createDesktopDlpContext({ documentId: workbookId });
    this.dlpContext = dlp;
    this.aiCellFunctions = new AiCellFunctionEngine({
      onUpdate: () => this.refresh(),
      workbookId,
      cache: { persistKey: "formula:ai_cell_cache" },
    });

    // Initial layout + render.
    this.onResize();
    this.uiReady = true;
  }

  destroy(): void {
    this.disposed = true;
    this.domAbort.abort();
    if (this.commentsDocUpdateListener) {
      this.commentsDoc.off("update", this.commentsDocUpdateListener);
      this.commentsDocUpdateListener = null;
    }
    this.formulaBarCompletion?.destroy();
    this.wasmUnsubscribe?.();
    this.wasmUnsubscribe = null;
    this.wasmEngine?.terminate();
    this.wasmEngine = null;
    this.stopCommentPersistence?.();
    this.resizeObserver.disconnect();
    if (this.dragAutoScrollRaf != null) {
      if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
      else globalThis.clearTimeout(this.dragAutoScrollRaf);
      this.dragAutoScrollRaf = null;
    }
    if (this.windowKeyDownListener && typeof window !== "undefined") {
      window.removeEventListener("keydown", this.windowKeyDownListener);
      this.windowKeyDownListener = null;
    }
    this.outlineButtons.clear();
    this.chartElements.clear();
    this.root.replaceChildren();
  }

  /**
   * Request a full redraw of the grid + overlays.
   *
   * This is intentionally public so external callers (e.g. AI tool execution)
   * can update the UI after mutating the DocumentController.
   */
  refresh(mode: "full" | "scroll" = "full"): void {
    if (this.disposed) return;
    if (this.renderScheduled) {
      // Upgrade a pending scroll-only render to a full render if needed.
      if (mode === "full") this.pendingRenderMode = "full";
      return;
    }

    this.renderScheduled = true;
    this.pendingRenderMode = mode;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.renderScheduled = false;
      if (this.disposed) return;
      const renderMode = this.pendingRenderMode;
      this.pendingRenderMode = "full";
      this.renderGrid();
      this.renderCharts(renderMode === "full");
      this.renderReferencePreview();
      this.renderSelection();
      if (renderMode === "full") this.updateStatus();
    });
  }

  focus(): void {
    this.root.focus();
  }

  async whenIdle(): Promise<void> {
    while (true) {
      const wasm = this.wasmSyncPromise;
      await Promise.all([this.idle.whenIdle(), wasm.catch(() => {})]);
      if (this.wasmSyncPromise === wasm) return;
    }
  }

  getRecalcCount(): number {
    return this.engine.recalcCount;
  }

  getDocument(): DocumentController {
    return this.document;
  }

  getSearchWorkbook(): DocumentWorkbookAdapter {
    return this.searchWorkbook;
  }

  getCurrentSheetId(): string {
    return this.sheetId;
  }

  getScroll(): { x: number; y: number } {
    return { x: this.scrollX, y: this.scrollY };
  }

  addChart(spec: CreateChartSpec): CreateChartResult {
    return this.chartStore.createChart(spec);
  }

  setChartTheme(theme: ChartTheme): void {
    this.chartTheme = theme;
    this.renderCharts(true);
  }

  listCharts(): readonly ChartRecord[] {
    return this.chartStore.listCharts();
  }

  private enqueueWasmSync(task: (engine: EngineClient) => Promise<void>): Promise<void> {
    const engine = this.wasmEngine;
    if (!engine) return Promise.resolve();

    const run = async () => {
      // The engine may have been replaced/terminated while the task was queued.
      if (this.wasmEngine !== engine) return;
      await task(engine);
    };

    this.wasmSyncPromise = this.wasmSyncPromise
      .catch(() => {
        // Ignore prior errors so the chain keeps flowing.
      })
      .then(run)
      .catch(() => {
        // Ignore WASM sync failures; the DocumentController remains the source of truth.
      });

    return this.wasmSyncPromise;
  }

  /**
   * Replace the DocumentController state from a snapshot, then hydrate the WASM engine in one step.
   *
   * This avoids N-per-cell RPC roundtrips during version restore by using the engine JSON load path.
   */
  async restoreDocumentState(snapshot: Uint8Array): Promise<void> {
    this.wasmSyncSuspended = true;
    try {
      // Ensure any in-flight sync operations finish before we replace the workbook.
      await this.wasmSyncPromise;
      this.computedValues.clear();
      this.document.applyState(snapshot);
      const sheetIds = this.document.getSheetIds();
      if (sheetIds.length > 0 && !sheetIds.includes(this.sheetId)) {
        this.sheetId = sheetIds[0];
        this.chartStore.setDefaultSheet(this.sheetId);
      }
      if (this.wasmEngine) {
        await this.enqueueWasmSync(async (engine) => {
          const changes = await engineHydrateFromDocument(engine, this.document);
          this.applyComputedChanges(changes);
        });
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
      const changes = await engineHydrateFromDocument(engine, this.document);
      this.applyComputedChanges(changes);

      this.wasmEngine = engine;
      this.wasmUnsubscribe = this.document.on(
        "change",
        ({ deltas, source, recalc }: { deltas: any[]; source?: string; recalc?: boolean }) => {
          if (!this.wasmEngine || this.wasmSyncSuspended) return;

          if (source === "applyState") {
            this.computedValues.clear();
            void this.enqueueWasmSync(async (worker) => {
              const changes = await engineHydrateFromDocument(worker, this.document);
              this.applyComputedChanges(changes);
            });
            return;
          }

          if (!Array.isArray(deltas) || deltas.length === 0) {
            if (recalc) {
              void this.enqueueWasmSync(async (worker) => {
                const changes = await worker.recalculate();
                this.applyComputedChanges(changes);
              });
            }
            return;
          }

          void this.enqueueWasmSync(async (worker) => {
            const changes = await engineApplyDeltas(worker, deltas, { recalculate: recalc !== false });
            this.applyComputedChanges(changes);
          });
        }
      );
    } catch {
      // Ignore initialization failures (e.g. missing WASM bundle).
      engine?.terminate();
      this.wasmEngine = null;
      this.wasmUnsubscribe?.();
      this.wasmUnsubscribe = null;
      this.computedValues.clear();
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
    this.renderCharts(true);
    this.renderSelection();
    this.updateStatus();
  }

  /**
   * Programmatically set the active cell (and optionally change sheets).
   */
  activateCell(target: { sheetId?: string; row: number; col: number }): void {
    let sheetChanged = false;
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.chartStore.setDefaultSheet(target.sheetId);
      this.renderGrid();
      this.renderCharts(true);
      sheetChanged = true;
    }
    this.selection = setActiveCell(this.selection, { row: target.row, col: target.col }, this.limits);
    this.ensureActiveCellVisible();
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (didScroll) this.ensureViewportMappingCurrent();
    this.renderSelection();
    this.updateStatus();
    if (sheetChanged) {
      // Sheet changes always require a full redraw (grid + charts may differ).
      this.refresh();
    } else if (didScroll) {
      this.refresh("scroll");
    }
    this.focus();
  }

  /**
   * Programmatically set the selection range (and optionally change sheets).
   */
  selectRange(target: { sheetId?: string; range: Range }): void {
    let sheetChanged = false;
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.chartStore.setDefaultSheet(target.sheetId);
      this.renderGrid();
      this.renderCharts(true);
      sheetChanged = true;
    }
    const active = { row: target.range.startRow, col: target.range.startCol };
    this.selection = buildSelection(
      { ranges: [target.range], active, anchor: active, activeRangeIndex: 0 },
      this.limits
    );
    this.ensureActiveCellVisible();
    const activeRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    const didScrollRange = activeRange ? this.scrollRangeIntoView(activeRange) : false;
    // Even if the range is too large to fit in the viewport, the active cell should never
    // become "lost" offscreen.
    const didScrollCell = this.scrollCellIntoView(this.selection.active);
    const didScroll = didScrollRange || didScrollCell;
    if (didScroll) this.ensureViewportMappingCurrent();
    this.renderSelection();
    this.updateStatus();
    if (sheetChanged) {
      this.refresh();
    } else if (didScroll) {
      this.refresh("scroll");
    }
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

  getCellRectA1(a1: string): { x: number; y: number; width: number; height: number } | null {
    const cell = parseA1(a1);
    return this.getCellRect(cell);
  }

  getFillHandleRect(): { x: number; y: number; width: number; height: number } | null {
    this.ensureViewportMappingCurrent();
    return this.selectionRenderer.getFillHandleRect(
      this.selection,
      {
        getCellRect: (cell) => this.getCellRect(cell),
        visibleRows: this.visibleRows,
        visibleCols: this.visibleCols,
      },
      {
        clipRect: {
          x: this.rowHeaderWidth,
          y: this.colHeaderHeight,
          width: this.viewportWidth(),
          height: this.viewportHeight(),
        },
      }
    );
  }

  async getCellDisplayValueA1(a1: string): Promise<string> {
    return this.getCellValueA1(a1);
  }

  async getCellDisplayTextForRenderA1(a1: string): Promise<string> {
    await this.whenIdle();
    const cell = parseA1(a1);
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (!state) return "";

    if (state.formula != null) {
      if (this.showFormulas) return state.formula;
      const computed = this.getCellComputedValue(cell);
      return computed == null ? "" : String(computed);
    }

    if (isRichTextValue(state.value)) return state.value.text;
    if (state.value != null) return String(state.value);
    return "";
  }

  getLastSelectionDrawn(): unknown {
    return this.selectionRenderer.getLastDebug();
  }

  getReferenceHighlightCount(): number {
    return this.referenceHighlights.length;
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

    // Charts should scroll with the grid but stay clipped under headers.
    this.chartLayer.style.left = `${this.rowHeaderWidth}px`;
    this.chartLayer.style.top = `${this.colHeaderHeight}px`;
    this.chartLayer.style.right = "0";
    this.chartLayer.style.bottom = "0";
    this.chartLayer.style.overflow = "hidden";

    this.clampScroll();
    this.syncScrollbars();

    this.renderGrid();
    this.renderCharts(true);
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
  }

  private renderGrid(): void {
    this.ensureViewportMappingCurrent();

    const ctx = this.gridCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.gridCanvas.width, this.gridCanvas.height);
    ctx.restore();

    ctx.save();
    ctx.fillStyle = resolveCssVar("--bg-primary", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.width, this.height);

    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;
    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();

    const cols = this.visibleCols.length;
    const rows = this.visibleRows.length;

    const startX = originX + this.visibleColStart * this.cellWidth - this.scrollX;
    const startY = originY + this.visibleRowStart * this.cellHeight - this.scrollY;
    const endX = startX + cols * this.cellWidth;
    const endY = startY + rows * this.cellHeight;

    ctx.strokeStyle = resolveCssVar("--grid-line", { fallback: "CanvasText" });
    ctx.lineWidth = 1;

    // Header backgrounds.
    ctx.fillStyle = resolveCssVar("--grid-header-bg", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.width, this.colHeaderHeight);
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.height);

    // Corner cell.
    ctx.fillStyle = resolveCssVar("--bg-tertiary", { fallback: "Canvas" });
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.colHeaderHeight);

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
    const errorTextColor = resolveCssVar("--error", { fallback: defaultTextColor });

    // Data region (cells) are clipped so we never paint into header chrome.
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX, originY, viewportWidth, viewportHeight);
    ctx.clip();

    // Grid lines for the data region.
    for (let r = 0; r <= rows; r++) {
      const y = startY + r * this.cellHeight + 0.5;
      ctx.beginPath();
      ctx.moveTo(startX, y);
      ctx.lineTo(endX, y);
      ctx.stroke();
    }

    for (let c = 0; c <= cols; c++) {
      const x = startX + c * this.cellWidth + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, startY);
      ctx.lineTo(x, endY);
      ctx.stroke();
    }

    for (let visualRow = 0; visualRow < rows; visualRow++) {
      const row = this.visibleRows[visualRow]!;
      for (let visualCol = 0; visualCol < cols; visualCol++) {
        const col = this.visibleCols[visualCol]!;
        const state = this.document.getCell(this.sheetId, { row, col }) as {
          value: unknown;
          formula: string | null;
        };
        if (!state) continue;

        let rich: { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> } | null =
          null;
        let color = defaultTextColor;

        if (state.formula != null) {
          if (this.showFormulas) {
            rich = { text: state.formula, runs: [] };
          } else {
            const computed = this.getCellComputedValue({ row, col });
            if (computed != null) {
              rich = { text: String(computed), runs: [] };
              if (typeof computed === "string" && computed.startsWith("#")) {
                color = errorTextColor;
              }
            }
          }
        } else if (isRichTextValue(state.value)) {
          rich = state.value;
        } else if (state.value != null) {
          rich = { text: String(state.value), runs: [] };
        }

        if (!rich || rich.text === "") continue;

        renderRichText(
          ctx,
          rich,
          {
            x: startX + visualCol * this.cellWidth,
            y: startY + visualRow * this.cellHeight,
            width: this.cellWidth,
            height: this.cellHeight
          },
          {
            padding: 4,
            align: "start",
            verticalAlign: "middle",
            fontFamily,
            fontSizePx,
            color
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
          x: startX + visualCol * this.cellWidth,
          y: startY + visualRow * this.cellHeight,
          width: this.cellWidth,
          height: this.cellHeight,
        });
      }
    }

    ctx.restore();

    // Header labels.
    ctx.fillStyle = resolveCssVar("--text-primary", { fallback: "CanvasText" });
    ctx.font = "12px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // Column header labels (horizontally scroll with the grid).
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX, 0, viewportWidth, originY);
    ctx.clip();
    for (let visualCol = 0; visualCol < cols; visualCol++) {
      const colIndex = this.visibleCols[visualCol]!;
      ctx.fillText(
        colToName(colIndex),
        startX + visualCol * this.cellWidth + this.cellWidth / 2,
        this.colHeaderHeight / 2
      );
    }
    ctx.restore();

    // Row header labels (vertically scroll with the grid).
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, originY, originX, viewportHeight);
    ctx.clip();
    for (let visualRow = 0; visualRow < rows; visualRow++) {
      const rowIndex = this.visibleRows[visualRow]!;
      ctx.fillText(
        String(rowIndex + 1),
        this.rowHeaderWidth / 2,
        startY + visualRow * this.cellHeight + this.cellHeight / 2
      );
    }
    ctx.restore();

    ctx.restore();

    this.renderOutlineControls();
  }

  private lowerBound(values: number[], target: number): number {
    let lo = 0;
    let hi = values.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if ((values[mid] ?? 0) < target) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  private visualIndexForRow(row: number): number {
    const direct = this.rowToVisual.get(row);
    if (direct !== undefined) return direct;
    return this.lowerBound(this.rowIndexByVisual, row);
  }

  private visualIndexForCol(col: number): number {
    const direct = this.colToVisual.get(col);
    if (direct !== undefined) return direct;
    return this.lowerBound(this.colIndexByVisual, col);
  }

  private chartAnchorToViewportRect(anchor: ChartRecord["anchor"]): { left: number; top: number; width: number; height: number } | null {
    if (!anchor || !("kind" in anchor)) return null;

    let left = 0;
    let top = 0;
    let width = 0;
    let height = 0;

    if (anchor.kind === "absolute") {
      left = emuToPx(anchor.xEmu);
      top = emuToPx(anchor.yEmu);
      width = emuToPx(anchor.cxEmu);
      height = emuToPx(anchor.cyEmu);
    } else if (anchor.kind === "oneCell") {
      left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu);
      top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu);
      width = emuToPx(anchor.cxEmu);
      height = emuToPx(anchor.cyEmu);
    } else if (anchor.kind === "twoCell") {
      left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu);
      top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu);
      const right = this.visualIndexForCol(anchor.toCol) * this.cellWidth + emuToPx(anchor.toColOffEmu);
      const bottom = this.visualIndexForRow(anchor.toRow) * this.cellHeight + emuToPx(anchor.toRowOffEmu);
      width = Math.max(0, right - left);
      height = Math.max(0, bottom - top);
    } else {
      return null;
    }

    if (width <= 0 || height <= 0) return null;

    return {
      left: left - this.scrollX,
      top: top - this.scrollY,
      width,
      height
    };
  }

  private renderCharts(renderContent: boolean): void {
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    const keep = new Set<string>();

    const createProvider = () => ({
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
    });

    const provider = renderContent ? createProvider() : null;

    for (const chart of charts) {
      keep.add(chart.id);
      const rect = this.chartAnchorToViewportRect(chart.anchor);
      if (!rect) continue;

      let host = this.chartElements.get(chart.id);
      const shouldRenderContent = renderContent || !host;
      if (!host) {
        host = document.createElement("div");
        host.setAttribute("data-testid", "chart-object");
        host.style.position = "absolute";
        host.style.pointerEvents = "none";
        host.style.overflow = "hidden";
        this.chartElements.set(chart.id, host);
        this.chartLayer.appendChild(host);
      }

      host.style.left = `${rect.left}px`;
      host.style.top = `${rect.top}px`;
      host.style.width = `${rect.width}px`;
      host.style.height = `${rect.height}px`;

      if (shouldRenderContent) {
        host.innerHTML = renderChartSvg(chart, provider ?? createProvider(), {
          width: rect.width,
          height: rect.height,
          theme: this.chartTheme,
        });
      }
    }

    for (const [id, el] of this.chartElements) {
      if (keep.has(id)) continue;
      el.remove();
      this.chartElements.delete(id);
    }
  }

  private renderSelection(): void {
    this.ensureViewportMappingCurrent();
    const clipRect = {
      x: this.rowHeaderWidth,
      y: this.colHeaderHeight,
      width: this.viewportWidth(),
      height: this.viewportHeight(),
    };

    this.selectionRenderer.render(
      this.selectionCtx,
      this.selection,
      {
        getCellRect: (cell) => this.getCellRect(cell),
        visibleRows: this.visibleRows,
        visibleCols: this.visibleCols,
      },
      {
        clipRect,
      }
    );

    if (this.fillPreviewRange) {
      const startRect = this.getCellRect({ row: this.fillPreviewRange.startRow, col: this.fillPreviewRange.startCol });
      const endRect = this.getCellRect({ row: this.fillPreviewRange.endRow, col: this.fillPreviewRange.endCol });
      if (startRect && endRect) {
        const x = startRect.x;
        const y = startRect.y;
        const width = endRect.x + endRect.width - startRect.x;
        const height = endRect.y + endRect.height - startRect.y;

        this.selectionCtx.save();
        this.selectionCtx.beginPath();
        this.selectionCtx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
        this.selectionCtx.clip();
        this.selectionCtx.strokeStyle = resolveCssVar("--warning", { fallback: "CanvasText" });
        this.selectionCtx.lineWidth = 2;
        this.selectionCtx.setLineDash([4, 3]);
        this.selectionCtx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);
        this.selectionCtx.restore();
      }
    }

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

  private onWindowKeyDown(e: KeyboardEvent): void {
    if (e.defaultPrevented) return;
    if (this.handleShowFormulasShortcut(e)) return;
    this.handleUndoRedoShortcut(e);
  }

  private handleShowFormulasShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return false;
    if (e.code !== "Backquote") return false;

    // Only trigger when *not* actively editing text.
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    e.preventDefault();
    this.showFormulas = !this.showFormulas;
    this.refresh();
    return true;
  }

  private handleUndoRedoShortcut(e: KeyboardEvent): boolean {
    const undo = isUndoKeyboardEvent(e);
    const redo = !undo && isRedoKeyboardEvent(e);
    if (!undo && !redo) return false;

    // Only trigger spreadsheet undo/redo when *not* actively editing text.
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    e.preventDefault();
    const did = undo ? this.document.undo() : this.document.redo();
    if (did) {
      this.syncEngineNow();
      this.refresh();
    }
    return true;
  }

  private isRowHidden(row: number): boolean {
    const entry = this.outline.rows.entry(row + 1);
    return isHidden(entry.hidden);
  }

  private isColHidden(col: number): boolean {
    const entry = this.outline.cols.entry(col + 1);
    return isHidden(entry.hidden);
  }

  private rebuildAxisVisibilityCache(): void {
    this.rowIndexByVisual = [];
    this.colIndexByVisual = [];
    this.rowToVisual.clear();
    this.colToVisual.clear();

    for (let r = 0; r < this.limits.maxRows; r += 1) {
      if (this.isRowHidden(r)) continue;
      this.rowToVisual.set(r, this.rowIndexByVisual.length);
      this.rowIndexByVisual.push(r);
    }

    for (let c = 0; c < this.limits.maxCols; c += 1) {
      if (this.isColHidden(c)) continue;
      this.colToVisual.set(c, this.colIndexByVisual.length);
      this.colIndexByVisual.push(c);
    }

    // Outline changes can affect scrollable content size.
    this.clampScroll();
    this.syncScrollbars();
  }

  private viewportWidth(): number {
    return Math.max(0, this.width - this.rowHeaderWidth);
  }

  private viewportHeight(): number {
    return Math.max(0, this.height - this.colHeaderHeight);
  }

  private contentWidth(): number {
    return this.colIndexByVisual.length * this.cellWidth;
  }

  private contentHeight(): number {
    return this.rowIndexByVisual.length * this.cellHeight;
  }

  private maxScrollX(): number {
    return Math.max(0, this.contentWidth() - this.viewportWidth());
  }

  private maxScrollY(): number {
    return Math.max(0, this.contentHeight() - this.viewportHeight());
  }

  private clampScroll(): void {
    const maxX = this.maxScrollX();
    const maxY = this.maxScrollY();
    this.scrollX = Math.min(Math.max(0, this.scrollX), maxX);
    this.scrollY = Math.min(Math.max(0, this.scrollY), maxY);
  }

  private setScroll(nextX: number, nextY: number): boolean {
    const prevX = this.scrollX;
    const prevY = this.scrollY;
    this.scrollX = nextX;
    this.scrollY = nextY;
    this.clampScroll();

    const changed = this.scrollX !== prevX || this.scrollY !== prevY;
    if (changed) {
      this.hideCommentTooltip();
      this.syncScrollbars();
    }
    return changed;
  }

  private scrollBy(deltaX: number, deltaY: number): void {
    const changed = this.setScroll(this.scrollX + deltaX, this.scrollY + deltaY);
    if (changed) this.refresh("scroll");
  }

  private updateViewportMapping(): void {
    const availableWidth = this.viewportWidth();
    const availableHeight = this.viewportHeight();

    const totalRows = this.rowIndexByVisual.length;
    const totalCols = this.colIndexByVisual.length;

    const overscan = 1;
    const firstRow = Math.max(0, Math.floor(this.scrollY / this.cellHeight) - overscan);
    const lastRow = Math.min(totalRows, Math.ceil((this.scrollY + availableHeight) / this.cellHeight) + overscan);

    const firstCol = Math.max(0, Math.floor(this.scrollX / this.cellWidth) - overscan);
    const lastCol = Math.min(totalCols, Math.ceil((this.scrollX + availableWidth) / this.cellWidth) + overscan);

    this.visibleRowStart = firstRow;
    this.visibleColStart = firstCol;
    this.visibleRows = this.rowIndexByVisual.slice(firstRow, lastRow);
    this.visibleCols = this.colIndexByVisual.slice(firstCol, lastCol);
  }

  private ensureViewportMappingCurrent(): void {
    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    const rowCount = this.rowIndexByVisual.length;
    const colCount = this.colIndexByVisual.length;

    const state = this.viewportMappingState;
    if (
      state &&
      state.scrollX === this.scrollX &&
      state.scrollY === this.scrollY &&
      state.viewportWidth === viewportWidth &&
      state.viewportHeight === viewportHeight &&
      state.rowCount === rowCount &&
      state.colCount === colCount
    ) {
      return;
    }

    this.updateViewportMapping();
    this.viewportMappingState = { scrollX: this.scrollX, scrollY: this.scrollY, viewportWidth, viewportHeight, rowCount, colCount };
  }

  private computeScrollbarThumb(options: {
    scrollPos: number;
    viewportSize: number;
    contentSize: number;
    trackSize: number;
    minThumbSize?: number;
  }): { size: number; offset: number } {
    const minThumbSize = options.minThumbSize ?? 24;
    const trackSize = Math.max(0, options.trackSize);
    const viewportSize = Math.max(0, options.viewportSize);
    const contentSize = Math.max(0, options.contentSize);
    const maxScroll = Math.max(0, contentSize - viewportSize);
    const scrollPos = Math.min(Math.max(0, options.scrollPos), maxScroll);

    if (trackSize === 0) return { size: 0, offset: 0 };
    if (contentSize === 0 || maxScroll === 0) return { size: trackSize, offset: 0 };

    const rawThumbSize = (viewportSize / contentSize) * trackSize;
    const thumbSize = Math.min(trackSize, Math.max(minThumbSize, rawThumbSize));
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    const offset = thumbTravel === 0 ? 0 : (scrollPos / maxScroll) * thumbTravel;

    return { size: thumbSize, offset };
  }

  private syncScrollbars(): void {
    const maxX = this.maxScrollX();
    const maxY = this.maxScrollY();
    const showH = maxX > 0;
    const showV = maxY > 0;

    const padding = 2;
    const thickness = this.scrollbarThickness;

    this.vScrollbarTrack.style.display = showV ? "block" : "none";
    this.hScrollbarTrack.style.display = showH ? "block" : "none";

    if (showV) {
      this.vScrollbarTrack.style.right = `${padding}px`;
      this.vScrollbarTrack.style.top = `${this.colHeaderHeight + padding}px`;
      this.vScrollbarTrack.style.bottom = `${(showH ? thickness : 0) + padding}px`;
      this.vScrollbarTrack.style.width = `${thickness}px`;

      const trackSize = Math.max(0, this.height - (this.colHeaderHeight + padding) - ((showH ? thickness : 0) + padding));
      const { size, offset } = this.computeScrollbarThumb({
        scrollPos: this.scrollY,
        viewportSize: this.viewportHeight(),
        contentSize: this.contentHeight(),
        trackSize
      });

      this.vScrollbarThumb.style.height = `${size}px`;
      this.vScrollbarThumb.style.transform = `translateY(${offset}px)`;
    }

    if (showH) {
      this.hScrollbarTrack.style.left = `${this.rowHeaderWidth + padding}px`;
      this.hScrollbarTrack.style.right = `${(showV ? thickness : 0) + padding}px`;
      this.hScrollbarTrack.style.bottom = `${padding}px`;
      this.hScrollbarTrack.style.height = `${thickness}px`;

      const trackSize = Math.max(0, this.width - (this.rowHeaderWidth + padding) - ((showV ? thickness : 0) + padding));
      const { size, offset } = this.computeScrollbarThumb({
        scrollPos: this.scrollX,
        viewportSize: this.viewportWidth(),
        contentSize: this.contentWidth(),
        trackSize
      });

      this.hScrollbarThumb.style.width = `${size}px`;
      this.hScrollbarThumb.style.transform = `translateX(${offset}px)`;
    }
  }

  private onWheel(e: WheelEvent): void {
    const target = e.target as HTMLElement | null;
    if (target?.closest('[data-testid="comments-panel"]')) return;
    if (e.ctrlKey) return;

    let deltaX = e.deltaX;
    let deltaY = e.deltaY;

    if (e.deltaMode === 1) {
      // DOM_DELTA_LINE: browsers use a "line" abstraction; normalize to CSS pixels.
      const line = 16;
      deltaX *= line;
      deltaY *= line;
    } else if (e.deltaMode === 2) {
      // DOM_DELTA_PAGE.
      deltaX *= this.viewportWidth();
      deltaY *= this.viewportHeight();
    }

    // Common UX: shift+wheel scrolls horizontally.
    if (e.shiftKey && deltaX === 0) {
      deltaX = deltaY;
      deltaY = 0;
    }

    if (deltaX === 0 && deltaY === 0) return;
    e.preventDefault();
    this.scrollBy(deltaX, deltaY);
  }

  private onScrollbarThumbPointerDown(e: PointerEvent, axis: "x" | "y"): void {
    e.preventDefault();
    e.stopPropagation();

    const thumb = axis === "y" ? this.vScrollbarThumb : this.hScrollbarThumb;
    const track = axis === "y" ? this.vScrollbarTrack : this.hScrollbarTrack;

    const trackRect = track.getBoundingClientRect();
    const thumbRect = thumb.getBoundingClientRect();

    const pointerPos = axis === "y" ? e.clientY : e.clientX;
    const thumbStart = axis === "y" ? thumbRect.top : thumbRect.left;
    const trackStart = axis === "y" ? trackRect.top : trackRect.left;
    const trackSize = axis === "y" ? trackRect.height : trackRect.width;
    const thumbSize = axis === "y" ? thumbRect.height : thumbRect.width;

    const grabOffset = pointerPos - thumbStart;
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    const maxScroll = axis === "y" ? this.maxScrollY() : this.maxScrollX();

    this.scrollbarDrag = { axis, pointerId: e.pointerId, grabOffset, thumbTravel, trackStart, maxScroll };

    (thumb as HTMLElement).setPointerCapture(e.pointerId);
  }

  private onScrollbarTrackPointerDown(e: PointerEvent, axis: "x" | "y"): void {
    // Clicking the track should scroll, but should not start a selection drag.
    e.preventDefault();
    e.stopPropagation();

    const thumb = axis === "y" ? this.vScrollbarThumb : this.hScrollbarThumb;
    const track = axis === "y" ? this.vScrollbarTrack : this.hScrollbarTrack;
    const trackRect = track.getBoundingClientRect();
    const thumbRect = thumb.getBoundingClientRect();

    const trackSize = axis === "y" ? trackRect.height : trackRect.width;
    const thumbSize = axis === "y" ? thumbRect.height : thumbRect.width;
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    if (thumbTravel === 0) return;

    const pointerPos = axis === "y" ? e.clientY - trackRect.top : e.clientX - trackRect.left;
    const targetOffset = pointerPos - thumbSize / 2;
    const clamped = Math.min(Math.max(0, targetOffset), thumbTravel);

    const maxScroll = axis === "y" ? this.maxScrollY() : this.maxScrollX();
    const nextScroll = (clamped / thumbTravel) * maxScroll;

    const changed = axis === "y" ? this.setScroll(this.scrollX, nextScroll) : this.setScroll(nextScroll, this.scrollY);
    if (changed) this.refresh("scroll");
  }

  private onScrollbarThumbPointerMove(e: PointerEvent): void {
    const drag = this.scrollbarDrag;
    if (!drag) return;
    if (e.pointerId !== drag.pointerId) return;

    e.preventDefault();

    const pointerPos = drag.axis === "y" ? e.clientY : e.clientX;
    const thumbOffset = pointerPos - drag.trackStart - drag.grabOffset;
    const clamped = Math.min(Math.max(0, thumbOffset), drag.thumbTravel);
    const nextScroll = drag.thumbTravel === 0 ? 0 : (clamped / drag.thumbTravel) * drag.maxScroll;

    const changed =
      drag.axis === "y" ? this.setScroll(this.scrollX, nextScroll) : this.setScroll(nextScroll, this.scrollY);

    if (changed) this.refresh("scroll");
  }

  private onScrollbarThumbPointerUp(e: PointerEvent): void {
    const drag = this.scrollbarDrag;
    if (!drag) return;
    if (e.pointerId !== drag.pointerId) return;
    this.scrollbarDrag = null;
  }

  private scrollCellIntoView(cell: CellCoord, paddingPx = 8): boolean {
    const visualRow = this.rowToVisual.get(cell.row);
    const visualCol = this.colToVisual.get(cell.col);
    if (visualRow === undefined || visualCol === undefined) return false;

    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    if (viewportWidth <= 0 || viewportHeight <= 0) return false;

    const left = visualCol * this.cellWidth;
    const top = visualRow * this.cellHeight;
    const right = left + this.cellWidth;
    const bottom = top + this.cellHeight;

    const pad = Math.max(0, paddingPx);
    let nextX = this.scrollX;
    let nextY = this.scrollY;

    if (left < nextX + pad) {
      nextX = left - pad;
    } else if (right > nextX + viewportWidth - pad) {
      nextX = right - viewportWidth + pad;
    }

    if (top < nextY + pad) {
      nextY = top - pad;
    } else if (bottom > nextY + viewportHeight - pad) {
      nextY = bottom - viewportHeight + pad;
    }

    return this.setScroll(nextX, nextY);
  }

  private scrollRangeIntoView(range: Range, paddingPx = 8): boolean {
    const startRow = Math.max(0, Math.min(this.limits.maxRows - 1, range.startRow));
    const endRow = Math.max(0, Math.min(this.limits.maxRows - 1, range.endRow));
    const startCol = Math.max(0, Math.min(this.limits.maxCols - 1, range.startCol));
    const endCol = Math.max(0, Math.min(this.limits.maxCols - 1, range.endCol));

    const startVisualRow = this.visualIndexForRow(startRow);
    const endVisualRow = this.visualIndexForRow(endRow);
    const startVisualCol = this.visualIndexForCol(startCol);
    const endVisualCol = this.visualIndexForCol(endCol);

    const left = Math.min(startVisualCol, endVisualCol) * this.cellWidth;
    const top = Math.min(startVisualRow, endVisualRow) * this.cellHeight;
    const right = (Math.max(startVisualCol, endVisualCol) + 1) * this.cellWidth;
    const bottom = (Math.max(startVisualRow, endVisualRow) + 1) * this.cellHeight;

    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    if (viewportWidth <= 0 || viewportHeight <= 0) return false;

    const pad = Math.max(0, paddingPx);
    let nextX = this.scrollX;
    let nextY = this.scrollY;

    // Only attempt to fully fit the range when it fits within the viewport.
    // Otherwise, fall back to keeping the active cell visible.
    if (right - left <= viewportWidth - pad * 2) {
      if (left < nextX + pad) {
        nextX = left - pad;
      } else if (right > nextX + viewportWidth - pad) {
        nextX = right - viewportWidth + pad;
      }
    }

    if (bottom - top <= viewportHeight - pad * 2) {
      if (top < nextY + pad) {
        nextY = top - pad;
      } else if (bottom > nextY + viewportHeight - pad) {
        nextY = bottom - viewportHeight + pad;
      }
    }

    return this.setScroll(nextX, nextY);
  }

  private renderOutlineControls(): void {
    if (!this.outline.pr.showOutlineSymbols) {
      for (const button of this.outlineButtons.values()) button.remove();
      this.outlineButtons.clear();
      return;
    }

    const keep = new Set<string>();
    const size = 14;
    const padding = 4;
    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;

    // Row group toggles live in the row header.
    for (let visualRow = 0; visualRow < this.visibleRows.length; visualRow++) {
      const rowIndex = this.visibleRows[visualRow]!;
      const summaryIndex = rowIndex + 1; // 1-based
      const entry = this.outline.rows.entry(summaryIndex);
      const details = groupDetailRange(this.outline.rows, summaryIndex, entry.level, this.outline.pr.summaryBelow);
      if (!details) continue;

      const key = `row:${summaryIndex}`;
      keep.add(key);
      let button = this.outlineButtons.get(key);
      if (!button) {
        button = document.createElement("button");
        button.className = "outline-toggle";
        button.type = "button";
        button.setAttribute("data-testid", `outline-toggle-row-${summaryIndex}`);
        button.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          this.outline.toggleRowGroup(summaryIndex);
          this.onOutlineUpdated();
        });
        button.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
        });
        this.outlineButtons.set(key, button);
        this.outlineLayer.appendChild(button);
      }

      button.textContent = entry.collapsed ? "+" : "-";
      button.style.left = `${padding}px`;
      const visualIndex = this.visibleRowStart + visualRow;
      const rowTop = originY + visualIndex * this.cellHeight - this.scrollY;
      button.style.top = `${rowTop + (this.cellHeight - size) / 2}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;
    }

    // Column group toggles live in the column header.
    for (let visualCol = 0; visualCol < this.visibleCols.length; visualCol++) {
      const colIndex = this.visibleCols[visualCol]!;
      const summaryIndex = colIndex + 1; // 1-based
      const entry = this.outline.cols.entry(summaryIndex);
      const details = groupDetailRange(this.outline.cols, summaryIndex, entry.level, this.outline.pr.summaryRight);
      if (!details) continue;

      const key = `col:${summaryIndex}`;
      keep.add(key);
      let button = this.outlineButtons.get(key);
      if (!button) {
        button = document.createElement("button");
        button.className = "outline-toggle";
        button.type = "button";
        button.setAttribute("data-testid", `outline-toggle-col-${summaryIndex}`);
        button.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          this.outline.toggleColGroup(summaryIndex);
          this.onOutlineUpdated();
        });
        button.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
        });
        this.outlineButtons.set(key, button);
        this.outlineLayer.appendChild(button);
      }

      button.textContent = entry.collapsed ? "+" : "-";
      const visualIndex = this.visibleColStart + visualCol;
      const colLeft = originX + visualIndex * this.cellWidth - this.scrollX;
      button.style.left = `${colLeft + (this.cellWidth - size) / 2}px`;
      button.style.top = `${padding}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;
    }

    for (const [key, button] of this.outlineButtons) {
      if (keep.has(key)) continue;
      button.remove();
      this.outlineButtons.delete(key);
    }
  }

  private onOutlineUpdated(): void {
    this.rebuildAxisVisibilityCache();
    this.ensureActiveCellVisible();
    this.scrollCellIntoView(this.selection.active);
    this.refresh();
    this.focus();
  }

  private closestVisibleIndexInRange(values: number[], target: number, start: number, end: number): number | null {
    if (values.length === 0) return null;
    const rangeStart = Math.min(start, end);
    const rangeEnd = Math.max(start, end);

    const startIdx = this.lowerBound(values, rangeStart);
    const endExclusive = this.lowerBound(values, rangeEnd + 1);
    if (startIdx >= endExclusive) return null;

    const idx = this.lowerBound(values, target);
    const clampedIdx = Math.min(Math.max(idx, startIdx), endExclusive - 1);

    let bestIdx = clampedIdx;
    let bestValue = values[bestIdx] ?? null;
    if (bestValue == null) return null;
    let bestDist = Math.abs(bestValue - target);

    const belowIdx = idx - 1;
    if (belowIdx >= startIdx && belowIdx < endExclusive) {
      const below = values[belowIdx];
      if (below != null) {
        const dist = Math.abs(below - target);
        if (dist < bestDist) {
          bestDist = dist;
          bestIdx = belowIdx;
          bestValue = below;
        }
      }
    }

    const aboveIdx = idx;
    if (aboveIdx >= startIdx && aboveIdx < endExclusive) {
      const above = values[aboveIdx];
      if (above != null) {
        const dist = Math.abs(above - target);
        if (dist < bestDist) {
          bestDist = dist;
          bestIdx = aboveIdx;
          bestValue = above;
        }
      }
    }

    return bestValue;
  }

  private ensureActiveCellVisible(): void {
    const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0] ?? null;
    let { row, col } = this.selection.active;
    let canPreserveSelection = range != null;

    if (this.isRowHidden(row)) {
      const withinRange = range
        ? this.closestVisibleIndexInRange(this.rowIndexByVisual, row, range.startRow, range.endRow)
        : null;
      if (withinRange != null) {
        row = withinRange;
      } else {
        row = this.findNextVisibleRow(row, 1) ?? this.findNextVisibleRow(row, -1) ?? row;
        canPreserveSelection = false;
      }
    }
    if (this.isColHidden(col)) {
      const withinRange = range
        ? this.closestVisibleIndexInRange(this.colIndexByVisual, col, range.startCol, range.endCol)
        : null;
      if (withinRange != null) {
        col = withinRange;
      } else {
        col = this.findNextVisibleCol(col, 1) ?? this.findNextVisibleCol(col, -1) ?? col;
        canPreserveSelection = false;
      }
    }

    if (row !== this.selection.active.row || col !== this.selection.active.col) {
      if (canPreserveSelection) {
        this.selection = buildSelection(
          {
            ranges: this.selection.ranges,
            active: { row, col },
            anchor: this.selection.anchor,
            activeRangeIndex: this.selection.activeRangeIndex,
          },
          this.limits
        );
      } else {
        // If there are no visible cells inside the current selection range (e.g. a fully-hidden row),
        // fall back to collapsing to the nearest visible cell to keep interaction predictable.
        this.selection = setActiveCell(this.selection, { row, col }, this.limits);
      }
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
    if (cell.row < 0 || cell.row >= this.limits.maxRows) return null;
    if (cell.col < 0 || cell.col >= this.limits.maxCols) return null;

    const rowDirect = this.rowToVisual.get(cell.row);
    const colDirect = this.colToVisual.get(cell.col);

    // Even when the outline hides rows/cols, downstream overlays still need a
    // stable coordinate space. Hidden rows/cols collapse to zero size and share
    // the same origin as the next visible row/col.
    const visualRow = rowDirect ?? this.lowerBound(this.rowIndexByVisual, cell.row);
    const visualCol = colDirect ?? this.lowerBound(this.colIndexByVisual, cell.col);

    return {
      x: this.rowHeaderWidth + visualCol * this.cellWidth - this.scrollX,
      y: this.colHeaderHeight + visualRow * this.cellHeight - this.scrollY,
      width: colDirect == null ? 0 : this.cellWidth,
      height: rowDirect == null ? 0 : this.cellHeight
    };
  }

  private cellFromPoint(pointX: number, pointY: number): CellCoord {
    // Pointer capture means we can receive coordinates outside the grid bounds
    // while the user is dragging a selection. Clamp to the current viewport so
    // we select the edge cell instead of snapping to the end of the sheet.
    const maxX = Math.max(this.rowHeaderWidth, this.width - 1);
    const maxY = Math.max(this.colHeaderHeight, this.height - 1);
    const clampedX = Math.min(Math.max(pointX, this.rowHeaderWidth), maxX);
    const clampedY = Math.min(Math.max(pointY, this.colHeaderHeight), maxY);

    const sheetX = this.scrollX + (clampedX - this.rowHeaderWidth);
    const sheetY = this.scrollY + (clampedY - this.colHeaderHeight);

    const colVisual = Math.floor(sheetX / this.cellWidth);
    const rowVisual = Math.floor(sheetY / this.cellHeight);

    const safeColVisual = Math.max(0, Math.min(this.colIndexByVisual.length - 1, colVisual));
    const safeRowVisual = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, rowVisual));

    const col = this.colIndexByVisual[safeColVisual] ?? 0;
    const row = this.rowIndexByVisual[safeRowVisual] ?? 0;
    return { row, col };
  }

  private maybeStartDragAutoScroll(): void {
    if (this.disposed) return;
    if (!this.dragState || !this.dragPointerPos) return;

    const margin = 8;
    const { x, y } = this.dragPointerPos;
    const left = this.rowHeaderWidth;
    const top = this.colHeaderHeight;
    const right = this.width;
    const bottom = this.height;

    const outside = x < left - margin || x > right + margin || y < top - margin || y > bottom + margin;
    if (!outside) {
      if (this.dragAutoScrollRaf != null) {
        if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
        else globalThis.clearTimeout(this.dragAutoScrollRaf);
        this.dragAutoScrollRaf = null;
      }
      return;
    }

    if (this.dragAutoScrollRaf != null) return;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 16);

    const tick = () => {
      this.dragAutoScrollRaf = null;
      if (!this.dragState || !this.dragPointerPos) return;

      const px = this.dragPointerPos.x;
      const py = this.dragPointerPos.y;

      let deltaX = 0;
      let deltaY = 0;

      if (px < left - margin) {
        deltaX = -Math.min(this.cellWidth, left - margin - px);
      } else if (px > right + margin) {
        deltaX = Math.min(this.cellWidth, px - (right + margin));
      }

      if (py < top - margin) {
        deltaY = -Math.min(this.cellHeight, top - margin - py);
      } else if (py > bottom + margin) {
        deltaY = Math.min(this.cellHeight, py - (bottom + margin));
      }

      if (deltaX === 0 && deltaY === 0) return;

      const didScroll = this.setScroll(this.scrollX + deltaX, this.scrollY + deltaY);
      if (!didScroll) return;

      const cell = this.cellFromPoint(px, py);
      if (this.dragState.mode === "fill") {
        const source = this.dragState.sourceRange;
        const target: Range = {
          startRow: Math.min(source.startRow, cell.row),
          endRow: Math.max(source.endRow, cell.row),
          startCol: Math.min(source.startCol, cell.col),
          endCol: Math.max(source.endCol, cell.col)
        };
        this.dragState.targetRange = target;
        this.dragState.endCell = cell;
        this.fillPreviewRange =
          target.startRow === source.startRow &&
          target.endRow === source.endRow &&
          target.startCol === source.startCol &&
          target.endCol === source.endCol
            ? null
            : target;
      } else {
        this.selection = extendSelectionToCell(this.selection, cell, this.limits);

        if (this.dragState.mode === "formula" && this.formulaBar) {
          const r = this.selection.ranges[0];
          if (r) {
            this.formulaBar.updateRangeSelection({
              start: { row: r.startRow, col: r.startCol },
              end: { row: r.endRow, col: r.endCol }
            });
          }
        }
      }

      // Repaint immediately (we're already in rAF / a timer tick).
      this.renderGrid();
      this.renderCharts(false);
      this.renderReferencePreview();
      this.renderSelection();
      this.updateStatus();

      this.dragAutoScrollRaf = schedule(tick) as unknown as number;
    };

    this.dragAutoScrollRaf = schedule(tick) as unknown as number;
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
      const sheetX = this.scrollX + (x - this.rowHeaderWidth);
      const visualCol = Math.floor(sheetX / this.cellWidth);
      const safeVisualCol = Math.max(0, Math.min(this.colIndexByVisual.length - 1, visualCol));
      const col = this.colIndexByVisual[safeVisualCol] ?? 0;
      this.selection = selectColumns(this.selection, col, col, { additive: primary }, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    // Row header selects entire row.
    if (x < this.rowHeaderWidth && y >= this.colHeaderHeight) {
      const sheetY = this.scrollY + (y - this.colHeaderHeight);
      const visualRow = Math.floor(sheetY / this.cellHeight);
      const safeVisualRow = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, visualRow));
      const row = this.rowIndexByVisual[safeVisualRow] ?? 0;
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
      this.dragPointerPos = { x, y };
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

    const fillHandle = this.getFillHandleRect();
    if (
      fillHandle &&
      x >= fillHandle.x &&
      x <= fillHandle.x + fillHandle.width &&
      y >= fillHandle.y &&
      y <= fillHandle.y + fillHandle.height
    ) {
      e.preventDefault();
      const sourceRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      if (sourceRange) {
        this.dragState = {
          pointerId: e.pointerId,
          mode: "fill",
          sourceRange,
          targetRange: sourceRange,
          endCell: { row: sourceRange.endRow, col: sourceRange.endCol }
        };
        this.dragPointerPos = { x, y };
        this.fillPreviewRange = null;
        this.root.setPointerCapture(e.pointerId);
        this.focus();
        return;
      }
    }

    this.dragState = { pointerId: e.pointerId, mode: "normal" };
    this.dragPointerPos = { x, y };
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
    if (this.scrollbarDrag) {
      this.onScrollbarThumbPointerMove(e);
      return;
    }

    if (!this.dragState) {
      const target = e.target as HTMLElement | null;
      if (target) {
        if (
          this.vScrollbarTrack.contains(target) ||
          this.hScrollbarTrack.contains(target) ||
          target.closest(".outline-toggle")
        ) {
          this.hideCommentTooltip();
          this.root.style.cursor = "";
          return;
        }
      }
    }

    const rect = this.root.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    if (this.dragPointerPos) {
      this.dragPointerPos.x = x;
      this.dragPointerPos.y = y;
    }

    if (this.dragState) {
      if (e.pointerId !== this.dragState.pointerId) return;
      if (this.editor.isOpen()) return;
      this.hideCommentTooltip();
      const cell = this.cellFromPoint(x, y);

      if (this.dragState.mode === "fill") {
        const source = this.dragState.sourceRange;
        const target: Range = {
          startRow: Math.min(source.startRow, cell.row),
          endRow: Math.max(source.endRow, cell.row),
          startCol: Math.min(source.startCol, cell.col),
          endCol: Math.max(source.endCol, cell.col)
        };
        this.dragState.targetRange = target;
        this.dragState.endCell = cell;
        this.fillPreviewRange =
          target.startRow === source.startRow &&
          target.endRow === source.endRow &&
          target.startCol === source.startCol &&
          target.endCol === source.endCol
            ? null
            : target;
        this.renderSelection();
        this.maybeStartDragAutoScroll();
        return;
      }

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

      this.maybeStartDragAutoScroll();
      return;
    }

    const fillHandle = this.getFillHandleRect();
    const overFillHandle =
      fillHandle &&
      x >= fillHandle.x &&
      x <= fillHandle.x + fillHandle.width &&
      y >= fillHandle.y &&
      y <= fillHandle.y + fillHandle.height;
    const nextCursor = overFillHandle ? "crosshair" : "";
    if (this.root.style.cursor !== nextCursor) {
      this.root.style.cursor = nextCursor;
    }

    if (this.commentsPanelVisible) {
      // Don't show tooltips while the panel is open; it obscures the grid anyway.
      this.hideCommentTooltip();
      return;
    }

    if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
      this.hideCommentTooltip();
      this.root.style.cursor = "";
      return;
    }

    if (x < this.rowHeaderWidth || y < this.colHeaderHeight) {
      this.hideCommentTooltip();
      this.root.style.cursor = "";
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
    if (this.scrollbarDrag) {
      this.onScrollbarThumbPointerUp(e);
      return;
    }

    if (!this.dragState) return;
    if (e.pointerId !== this.dragState.pointerId) return;
    const state = this.dragState;
    this.dragState = null;
    this.dragPointerPos = null;
    if (this.dragAutoScrollRaf != null) {
      if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
      else globalThis.clearTimeout(this.dragAutoScrollRaf);
    }
    this.dragAutoScrollRaf = null;

    if (state.mode === "fill") {
      this.fillPreviewRange = null;
      const { sourceRange, targetRange, endCell } = state;
      const changed =
        targetRange.startRow !== sourceRange.startRow ||
        targetRange.endRow !== sourceRange.endRow ||
        targetRange.startCol !== sourceRange.startCol ||
        targetRange.endCol !== sourceRange.endCol;

      if (changed) {
        this.applyFill(sourceRange, targetRange);
        this.selection = buildSelection(
          {
            ranges: [targetRange],
            active: endCell,
            anchor: endCell,
            activeRangeIndex: 0
          },
          this.limits
        );
        this.refresh();
        this.focus();
      } else {
        // Clear any preview overlay.
        this.renderSelection();
      }
      return;
    }

    if (state.mode === "formula" && this.formulaBar) {
      this.formulaBar.endRangeSelection();
      // Restore focus to the formula bar without clearing its insertion state mid-drag.
      this.formulaBar.focus();
    }
  }

  private applyFill(sourceRange: Range, targetRange: Range): void {
    const sheetId = this.sheetId;
    const source = sourceRange;
    const target = targetRange;

    const sourceHeight = source.endRow - source.startRow + 1;
    const sourceWidth = source.endCol - source.startCol + 1;

    const isVertical1d =
      sourceWidth === 1 && sourceHeight >= 1 && target.startCol === source.startCol && target.endCol === source.endCol;
    const isHorizontal1d =
      sourceHeight === 1 && sourceWidth >= 1 && target.startRow === source.startRow && target.endRow === source.endRow;

    this.document.beginBatch({ label: "Fill" });
    try {
      // Excel-style numeric series for simple 1D numeric inputs (e.g. 1,2 -> 3,4).
      const didSeries = (isVertical1d || isHorizontal1d) && this.applyNumericSeriesFill(sheetId, source, target);
      if (!didSeries) {
        this.applyPatternFill(sheetId, source, target, sourceHeight, sourceWidth);
      }
    } finally {
      this.document.endBatch();
    }
  }

  private applyNumericSeriesFill(sheetId: string, source: Range, target: Range): boolean {
    const sourceHeight = source.endRow - source.startRow + 1;
    const sourceWidth = source.endCol - source.startCol + 1;
    const isVertical = sourceWidth === 1 && target.startCol === source.startCol && target.endCol === source.endCol;
    const isHorizontal = sourceHeight === 1 && target.startRow === source.startRow && target.endRow === source.endRow;
    if (!isVertical && !isHorizontal) return false;

    const len = isVertical ? sourceHeight : sourceWidth;
    if (len <= 0) return false;

    const nums: number[] = [];
    let outputAsString = true;

    for (let i = 0; i < len; i += 1) {
      const row = isVertical ? source.startRow + i : source.startRow;
      const col = isVertical ? source.startCol : source.startCol + i;
      const state = this.document.getCell(sheetId, { row, col }) as { value: unknown; formula: string | null };
      if (state?.formula != null) return false;

      const value = state?.value ?? null;
      if (typeof value !== "string") outputAsString = false;
      const num = coerceNumber(value);
      if (num == null) return false;
      nums.push(num);
    }

    let step = 0;
    if (nums.length >= 2) {
      step = nums[1]! - nums[0]!;
      for (let i = 2; i < nums.length; i += 1) {
        const expected = nums[0]! + step * i;
        if (nums[i] !== expected) return false;
      }
    }

    const start = nums[0]!;

    for (let row = target.startRow; row <= target.endRow; row += 1) {
      for (let col = target.startCol; col <= target.endCol; col += 1) {
        if (row >= source.startRow && row <= source.endRow && col >= source.startCol && col <= source.endCol) continue;
        const offset = isVertical ? row - source.startRow : col - source.startCol;
        const next = start + step * offset;
        this.document.setCellInput(sheetId, { row, col }, outputAsString ? String(next) : next);
      }
    }

    return true;
  }

  private applyPatternFill(sheetId: string, source: Range, target: Range, sourceHeight: number, sourceWidth: number): void {
    for (let row = target.startRow; row <= target.endRow; row += 1) {
      for (let col = target.startCol; col <= target.endCol; col += 1) {
        if (row >= source.startRow && row <= source.endRow && col >= source.startCol && col <= source.endCol) continue;

        const sourceRow = source.startRow + mod(row - source.startRow, sourceHeight);
        const sourceCol = source.startCol + mod(col - source.startCol, sourceWidth);
        const state = this.document.getCell(sheetId, { row: sourceRow, col: sourceCol }) as {
          value: unknown;
          formula: string | null;
        };

        const deltaRow = row - sourceRow;
        const deltaCol = col - sourceCol;

        if (state?.formula != null) {
          this.document.setCellInput(sheetId, { row, col }, shiftA1References(state.formula, deltaRow, deltaCol));
        } else {
          this.document.setCellInput(sheetId, { row, col }, state?.value ?? null);
        }
      }
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (this.inlineEditController.isOpen()) {
      return;
    }
    if (this.editor.isOpen()) {
      // The editor handles Enter/Tab/Escape itself. We keep focus on the textarea.
      return;
    }

    if (this.handleUndoRedoShortcut(e)) return;
    if (this.handleShowFormulasShortcut(e)) return;
    if (this.handleClipboardShortcut(e)) return;

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
    if (primary && (e.key === "k" || e.key === "K")) {
      // Inline edit (Cmd/Ctrl+K) should not trigger while the formula bar is actively editing.
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.inlineEditController.open();
      return;
    }
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

    // Page navigation (Excel-style): PageUp/PageDown move by approximately one viewport.
    // Alt+PageUp/PageDown scroll horizontally.
    if (!primary && (e.key === "PageDown" || e.key === "PageUp")) {
      e.preventDefault();
      this.ensureActiveCellVisible();
      const dir = e.key === "PageDown" ? 1 : -1;
      const visualRow = this.rowToVisual.get(this.selection.active.row);
      const visualCol = this.colToVisual.get(this.selection.active.col);
      if (visualRow === undefined || visualCol === undefined) return;

      if (e.altKey) {
        const pageCols = Math.max(1, Math.floor(this.viewportWidth() / this.cellWidth));
        const nextColVisual = Math.min(
          Math.max(0, visualCol + dir * pageCols),
          Math.max(0, this.colIndexByVisual.length - 1)
        );
        const col = this.colIndexByVisual[nextColVisual] ?? 0;
        this.selection = e.shiftKey
          ? extendSelectionToCell(this.selection, { row: this.selection.active.row, col }, this.limits)
          : setActiveCell(this.selection, { row: this.selection.active.row, col }, this.limits);
      } else {
        const pageRows = Math.max(1, Math.floor(this.viewportHeight() / this.cellHeight));
        const nextRowVisual = Math.min(
          Math.max(0, visualRow + dir * pageRows),
          Math.max(0, this.rowIndexByVisual.length - 1)
        );
        const row = this.rowIndexByVisual[nextRowVisual] ?? 0;
        this.selection = e.shiftKey
          ? extendSelectionToCell(this.selection, { row, col: this.selection.active.col }, this.limits)
          : setActiveCell(this.selection, { row, col: this.selection.active.col }, this.limits);
      }

      this.ensureActiveCellVisible();
      const didScroll = this.scrollCellIntoView(this.selection.active);
      if (didScroll) this.ensureViewportMappingCurrent();
      this.renderSelection();
      this.updateStatus();
      if (didScroll) this.refresh("scroll");
      return;
    }

    // Home/End without Ctrl/Cmd.
    // Home: first visible column in the current row.
    // End: last visible column in the current row.
    if (!primary && !e.altKey && (e.key === "Home" || e.key === "End")) {
      e.preventDefault();
      this.ensureActiveCellVisible();
      const row = this.selection.active.row;
      const col =
        e.key === "Home"
          ? (this.colIndexByVisual[0] ?? 0)
          : (this.colIndexByVisual[this.colIndexByVisual.length - 1] ?? 0);
      this.selection = e.shiftKey
        ? extendSelectionToCell(this.selection, { row, col }, this.limits)
        : setActiveCell(this.selection, { row, col }, this.limits);
      this.ensureActiveCellVisible();
      const didScroll = this.scrollCellIntoView(this.selection.active);
      if (didScroll) this.ensureViewportMappingCurrent();
      this.renderSelection();
      this.updateStatus();
      if (didScroll) this.refresh("scroll");
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
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (didScroll) this.ensureViewportMappingCurrent();
    this.renderSelection();
    this.updateStatus();
    if (didScroll) this.refresh("scroll");
  }

  private handleClipboardShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || e.altKey || e.shiftKey) return false;

    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    const key = e.key.toLowerCase();

    if (key === "c") {
      e.preventDefault();
      this.idle.track(this.copySelectionToClipboard());
      return true;
    }

    if (key === "x") {
      e.preventDefault();
      this.idle.track(this.cutSelectionToClipboard());
      return true;
    }

    if (key === "v") {
      e.preventDefault();
      this.idle.track(this.pasteClipboardToSelection());
      return true;
    }

    return false;
  }

  private async getClipboardProvider(): Promise<Awaited<ReturnType<typeof createClipboardProvider>>> {
    if (!this.clipboardProviderPromise) {
      this.clipboardProviderPromise = createClipboardProvider();
    }
    return this.clipboardProviderPromise;
  }

  private snapshotClipboardCells(range: Range): Array<Array<{ value: unknown; formula: string | null; styleId: number }>> {
    const cells: Array<Array<{ value: unknown; formula: string | null; styleId: number }>> = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const outRow: Array<{ value: unknown; formula: string | null; styleId: number }> = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.document.getCell(this.sheetId, { row, col }) as {
          value: unknown;
          formula: string | null;
          styleId: number;
        };
        outRow.push({ value: cell.value ?? null, formula: cell.formula ?? null, styleId: cell.styleId ?? 0 });
      }
      cells.push(outRow);
    }
    return cells;
  }

  private getClipboardCopyRange(): Range {
    const activeRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    const activeCellFallback: Range = {
      startRow: this.selection.active.row,
      endRow: this.selection.active.row,
      startCol: this.selection.active.col,
      endCol: this.selection.active.col
    };

    if (!activeRange) return activeCellFallback;

    if (this.selection.type === "row" || this.selection.type === "column" || this.selection.type === "all") {
      const used = this.computeUsedRange();
      const clipped = used ? intersectRanges(activeRange, used) : null;
      if (clipped) return clipped;

      // If a user selects an entire row/column with no used-range overlap, still copy a
      // bounded slice (the selected row/column within the UI's grid limits). For an
      // entirely empty sheet (`all`), fall back to the active cell to avoid generating
      // a massive empty payload.
      return this.selection.type === "all" ? activeCellFallback : activeRange;
    }

    return activeRange;
  }

  private async copySelectionToClipboard(): Promise<void> {
    try {
      const range = this.getClipboardCopyRange();
      const cells = this.snapshotClipboardCells(range);
      const dlp = this.dlpContext;
      const payload = copyRangeToClipboardPayload(
        this.document,
        this.sheetId,
        {
          start: { row: range.startRow, col: range.startCol },
          end: { row: range.endRow, col: range.endCol }
        },
        dlp
          ? {
              dlp: {
                documentId: dlp.documentId,
                classificationStore: dlp.classificationStore,
                policy: dlp.policy
              }
            }
          : undefined
      );
      const provider = await this.getClipboardProvider();
      await provider.write(payload);
      this.clipboardCopyContext = { range, payload, cells };
    } catch {
      // Ignore clipboard failures (permissions, platform restrictions).
    }
  }

  private async pasteClipboardToSelection(): Promise<void> {
    try {
      const provider = await this.getClipboardProvider();
      const content = await provider.read();
      const start = { ...this.selection.active };
      const ctx = this.clipboardCopyContext;
      let deltaRow = 0;
      let deltaCol = 0;

      const normalizeClipboardText = (text: string): string =>
        text
          .replace(/\r\n/g, "\n")
          .replace(/\r/g, "\n")
          // Some clipboard implementations add a trailing newline; ignore it when
          // detecting "internal" pastes for formula shifting.
          .replace(/\n+$/g, "");

      const isInternalPaste =
        ctx &&
        ((typeof content.text === "string" &&
          typeof ctx.payload.text === "string" &&
          normalizeClipboardText(content.text) === normalizeClipboardText(ctx.payload.text)) ||
          (typeof content.html === "string" &&
            typeof ctx.payload.html === "string" &&
            content.html === ctx.payload.html));

      const externalGrid = isInternalPaste ? null : parseClipboardContentToCellGrid(content);
      const internalCells = isInternalPaste ? ctx.cells : null;
      const rowCount = internalCells ? internalCells.length : externalGrid?.length ?? 0;
      const colCount = Math.max(
        0,
        ...(internalCells ? internalCells.map((row) => row.length) : []),
        ...(externalGrid ? externalGrid.map((row) => row.length) : [])
      );
      if (rowCount === 0 || colCount === 0) return;

      if (
        isInternalPaste
      ) {
        deltaRow = start.row - ctx.range.startRow;
        deltaCol = start.col - ctx.range.startCol;
      }

      const values = isInternalPaste
        ? internalCells!.map((row) =>
            row.map((cell) => {
              const rawFormula = cell.formula;
              const formula =
                rawFormula != null && (deltaRow !== 0 || deltaCol !== 0)
                  ? shiftA1References(rawFormula, deltaRow, deltaCol)
                  : rawFormula;
              if (formula != null) {
                return { formula, styleId: cell.styleId };
              }
              return { value: cell.value ?? null, styleId: cell.styleId };
            })
          )
        : externalGrid!.map((row) =>
            row.map((cell) => {
              const format = cell.format ?? null;
              const rawFormula = cell.formula;
              const formula =
                rawFormula != null && (deltaRow !== 0 || deltaCol !== 0)
                  ? shiftA1References(rawFormula, deltaRow, deltaCol)
                  : rawFormula;

              if (formula != null) {
                return { formula, format };
              }

              return { value: cell.value ?? null, format };
            })
          );

      this.document.setRangeValues(this.sheetId, start, values, { label: t("clipboard.paste") });

      const range: Range = {
        startRow: start.row,
        endRow: start.row + rowCount - 1,
        startCol: start.col,
        endCol: start.col + colCount - 1
      };
      this.selection = buildSelection({ ranges: [range], active: start, anchor: start, activeRangeIndex: 0 }, this.limits);

      this.syncEngineNow();
      this.refresh();
      this.focus();
    } catch {
      // Ignore clipboard failures (permissions, platform restrictions).
    }
  }

  private async cutSelectionToClipboard(): Promise<void> {
    try {
      const range = this.getClipboardCopyRange();
      const cells = this.snapshotClipboardCells(range);
      const dlp = this.dlpContext;
      const payload = copyRangeToClipboardPayload(
        this.document,
        this.sheetId,
        {
          start: { row: range.startRow, col: range.startCol },
          end: { row: range.endRow, col: range.endCol }
        },
        dlp
          ? {
              dlp: {
                documentId: dlp.documentId,
                classificationStore: dlp.classificationStore,
                policy: dlp.policy
              }
            }
          : undefined
      );

      const provider = await this.getClipboardProvider();
      await provider.write(payload);
      this.clipboardCopyContext = { range, payload, cells };

      const label = (() => {
        const translated = t("clipboard.cut");
        return translated === "clipboard.cut" ? "Cut" : translated;
      })();

      this.document.beginBatch({ label });
      this.document.clearRange(
        this.sheetId,
        {
          start: { row: range.startRow, col: range.startCol },
          end: { row: range.endRow, col: range.endCol }
        },
        { label }
      );
      this.document.endBatch();

      this.syncEngineNow();
      this.refresh();
      this.focus();
    } catch {
      // Ignore clipboard failures (permissions, platform restrictions).
    }
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
    return this.computeCellValue(this.sheetId, cell, memo, stack);
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

      let value = ref.value;
      if (value === undefined) value = null;
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
    sheetId: string,
    cell: CellCoord,
    memo: Map<string, SpreadsheetValue>,
    stack: Set<string>
  ): SpreadsheetValue {
    const address = cellToA1(cell);
    const computedKey = this.computedKey(sheetId, address);
    if (this.computedValues.has(computedKey)) {
      return this.computedValues.get(computedKey) ?? null;
    }
    const cached = memo.get(computedKey);
    if (cached !== undefined || memo.has(computedKey)) return cached ?? null;
    if (stack.has(computedKey)) return "#REF!";

    stack.add(computedKey);
    const state = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };
    let value: SpreadsheetValue;

    if (state?.formula != null) {
      const resolveSheetId = (name: string): string | null => {
        const trimmed = name.trim();
        if (!trimmed) return null;
        const knownSheets = this.document.getSheetIds();
        return knownSheets.find((id) => id.toLowerCase() === trimmed.toLowerCase()) ?? null;
      };

      value = evaluateFormula(state.formula, (ref) => {
        const normalized = ref.replaceAll("$", "").trim();
        let targetSheet = sheetId;
        let targetAddress = normalized;
        if (normalized.includes("!")) {
          const [maybeSheet, addr] = normalized.split("!", 2);
          if (maybeSheet && addr) {
            const resolved = resolveSheetId(maybeSheet);
            if (!resolved) return "#REF!";
            targetSheet = resolved;
            targetAddress = addr.trim();
          }
        }
        const coord = parseA1(targetAddress);
        return this.computeCellValue(targetSheet, coord, memo, stack);
      }, {
        ai: this.aiCellFunctions,
        cellAddress: `${sheetId}!${address}`,
      });
    } else if (state?.value != null) {
      value = isRichTextValue(state.value) ? state.value.text : (state.value as SpreadsheetValue);
    } else {
      value = null;
    }

    stack.delete(computedKey);
    memo.set(computedKey, value);
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
    this.ensureActiveCellVisible();
    this.scrollCellIntoView(this.selection.active);
    this.formulaEditCell = null;
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.refresh();
    this.focus();
  }

  private cancelFormulaBar(): void {
    if (this.formulaEditCell) {
      this.selection = setActiveCell(this.selection, this.formulaEditCell, this.limits);
    }
    this.formulaEditCell = null;
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.ensureActiveCellVisible();
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (didScroll) this.ensureViewportMappingCurrent();
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
    if (didScroll) this.refresh("scroll");
    this.focus();
  }

  private renderReferencePreview(): void {
    const ctx = this.referenceCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.referenceCanvas.width, this.referenceCanvas.height);
    ctx.restore();

    if (this.referenceHighlights.length === 0 && !this.referencePreview) return;
    this.ensureViewportMappingCurrent();

    const drawDashedRange = (startRow: number, endRow: number, startCol: number, endCol: number, color: string) => {
      // Clip preview rendering to the visible viewport so dragging a range that
      // extends offscreen doesn't crash (and still provides visual feedback).
      const visibleStartRow = this.visibleRows.find((row) => row >= startRow && row <= endRow) ?? null;
      const visibleEndRow = (() => {
        for (let i = this.visibleRows.length - 1; i >= 0; i -= 1) {
          const row = this.visibleRows[i]!;
          if (row >= startRow && row <= endRow) return row;
        }
        return null;
      })();
      const visibleStartCol = this.visibleCols.find((col) => col >= startCol && col <= endCol) ?? null;
      const visibleEndCol = (() => {
        for (let i = this.visibleCols.length - 1; i >= 0; i -= 1) {
          const col = this.visibleCols[i]!;
          if (col >= startCol && col <= endCol) return col;
        }
        return null;
      })();
      if (visibleStartRow == null || visibleEndRow == null || visibleStartCol == null || visibleEndCol == null) return;

      const startRect = this.getCellRect({ row: visibleStartRow, col: visibleStartCol });
      const endRect = this.getCellRect({ row: visibleEndRow, col: visibleEndCol });
      if (!startRect || !endRect) return;

      const x = startRect.x;
      const y = startRect.y;
      const width = endRect.x + endRect.width - startRect.x;
      const height = endRect.y + endRect.height - startRect.y;
      if (width <= 0 || height <= 0) return;

      ctx.save();
      ctx.beginPath();
      ctx.rect(this.rowHeaderWidth, this.colHeaderHeight, this.viewportWidth(), this.viewportHeight());
      ctx.clip();
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);
      ctx.restore();
    };

    for (const highlight of this.referenceHighlights) {
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);
      drawDashedRange(startRow, endRow, startCol, endCol, highlight.color);
    }

    if (this.referencePreview) {
      const startRow = Math.min(this.referencePreview.start.row, this.referencePreview.end.row);
      const endRow = Math.max(this.referencePreview.start.row, this.referencePreview.end.row);
      const startCol = Math.min(this.referencePreview.start.col, this.referencePreview.end.col);
      const endCol = Math.max(this.referencePreview.start.col, this.referencePreview.end.col);
      drawDashedRange(startRow, endRow, startCol, endCol, resolveCssVar("--warning", { fallback: "CanvasText" }));
    }
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

  private getInlineEditSelectionRange(): Range | null {
    const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    return range ? { ...range } : null;
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

function mod(n: number, m: number): number {
  if (!Number.isFinite(n) || !Number.isFinite(m) || m === 0) return 0;
  return ((n % m) + m) % m;
}

function coerceNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
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
