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
import { AuditingOverlayRenderer } from "../grid/auditing-overlays/AuditingOverlayRenderer";
import { computeAuditingOverlays, type AuditingMode, type AuditingOverlays } from "../grid/auditing-overlays/overlays";
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
import { createSchemaProviderFromSearchWorkbook } from "../ai/context/searchWorkbookSchemaProvider.js";
import { InlineEditController, type InlineEditLLMClient } from "../ai/inline-edit/inlineEditController";
import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import type { CellRange as GridCellRange, GridAxisSizeChange, GridViewportState } from "@formula/grid";
import { resolveDesktopGridMode, type DesktopGridMode } from "../grid/shared/desktopGridMode.js";
import { DocumentCellProvider } from "../grid/shared/documentCellProvider.js";
import { DesktopSharedGrid } from "../grid/shared/desktopSharedGrid.js";
import { applyFillCommitToDocumentController } from "../fill/applyFillCommit";
import type { CellRange as FillEngineRange, FillMode as FillHandleMode } from "@formula/fill-engine";

import * as Y from "yjs";
import { CommentManager, bindDocToStorage } from "@formula/collab-comments";
import type { Comment, CommentAuthor } from "@formula/collab-comments";

type EngineCellRef = { sheetId?: string; sheet?: string; row?: number; col?: number; address?: string; value?: unknown };
type AuditingCacheEntry = {
  precedents: string[];
  dependents: string[];
  precedentsError: string | null;
  dependentsError: string | null;
};

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return typeof (value as { then?: unknown } | null)?.then === "function";
}

function isInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value);
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
  applyChanges?(changes: unknown): unknown;
  recalculate?(): unknown;
  syncNow?(): unknown;
  beginBatch?(): unknown;
  endBatch?(): unknown;
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
  | {
      pointerId: number;
      mode: "fill";
      sourceRange: Range;
      targetRange: Range;
      endCell: CellCoord;
      fillMode: FillHandleMode;
      activeRangeIndex: number;
    };

export interface SpreadsheetAppStatusElements {
  activeCell: HTMLElement;
  selectionRange: HTMLElement;
  activeValue: HTMLElement;
  selectionSum?: HTMLElement;
  selectionAverage?: HTMLElement;
  selectionCount?: HTMLElement;
}

export type UndoRedoState = {
  canUndo: boolean;
  canRedo: boolean;
  undoLabel: string | null;
  redoLabel: string | null;
};

export interface SpreadsheetSelectionSummary {
  /**
   * Sum of numeric values in the current selection.
   *
   * - Uses computed values for formula cells.
   * - Ignores non-numeric values (text, booleans, errors).
   * - `null` when there are no numeric values selected.
   */
  sum: number | null;
  /**
   * Average of numeric values in the current selection.
   *
   * - Uses computed values for formula cells.
   * - Ignores non-numeric values (text, booleans, errors).
   * - `null` when there are no numeric values selected.
   */
  average: number | null;
  /**
   * "Count" as shown by default in Excel's status bar: number of *non-empty* cells.
   *
   * This counts cells with a value or a formula, and ignores format-only cells
   * (styleId-only entries).
   */
  count: number;
  /**
   * Number of numeric values in the selection (Excel "Numerical Count").
   */
  numericCount: number;
  /**
   * Alias for `count` (explicit name for consumers that also expose `numericCount`).
   */
  countNonEmpty: number;
}

export class SpreadsheetApp {
  private sheetId = "Sheet1";
  private readonly idle = new IdleTracker();
  private readonly computedValues = new Map<string, SpreadsheetValue>();
  private uiReady = false;
  private readonly gridMode: DesktopGridMode;
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
  private sharedGrid: DesktopSharedGrid | null = null;
  private sharedProvider: DocumentCellProvider | null = null;
  private readonly commentMeta = new Map<string, { resolved: boolean }>();
  private sharedGridSelectionSyncInProgress = false;
  private readonly sharedGridAxisCols = new Set<number>();
  private readonly sharedGridAxisRows = new Set<number>();

  private wasmEngine: EngineClient | null = null;
  private wasmSyncSuspended = false;
  private wasmUnsubscribe: (() => void) | null = null;
  private wasmSyncPromise: Promise<void> = Promise.resolve();
  private auditingUnsubscribe: (() => void) | null = null;

  private gridCanvas: HTMLCanvasElement;
  private chartLayer: HTMLDivElement;
  private sharedChartPanes:
    | {
        topLeft: HTMLDivElement;
        topRight: HTMLDivElement;
        bottomLeft: HTMLDivElement;
        bottomRight: HTMLDivElement;
      }
    | null = null;
  private sharedChartPaneLayout: { frozenContentWidth: number; frozenContentHeight: number } | null = null;
  private referenceCanvas: HTMLCanvasElement;
  private auditingCanvas: HTMLCanvasElement;
  private selectionCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private referenceCtx: CanvasRenderingContext2D;
  private auditingCtx: CanvasRenderingContext2D;
  private selectionCtx: CanvasRenderingContext2D;

  private auditingRenderer = new AuditingOverlayRenderer();
  private auditingMode: "off" | AuditingMode = "off";
  private auditingTransitive = false;
  private auditingHighlights: AuditingOverlays = { precedents: new Set(), dependents: new Set() };
  private auditingErrors: { precedents: string | null; dependents: string | null } = { precedents: null, dependents: null };
  private readonly auditingCache = new Map<string, AuditingCacheEntry>();
  private auditingLegend!: HTMLDivElement;
  private auditingUpdateScheduled = false;
  private auditingNeedsUpdateAfterDrag = false;
  private auditingRequestId = 0;
  private auditingIdlePromise: Promise<void> = Promise.resolve();
  private auditingLastCellKey: string | null = null;

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

  // Frozen panes for the active sheet. `frozenRows`/`frozenCols` are sheet
  // indices (0-based counts); `frozenWidth`/`frozenHeight` are viewport pixels
  // for the visible (i.e. not hidden) frozen rows/cols.
  private frozenRows = 0;
  private frozenCols = 0;
  private frozenVisibleRows: number[] = [];
  private frozenVisibleCols: number[] = [];
  private frozenWidth = 0;
  private frozenHeight = 0;

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
  private formulaSelectionRenderer = new SelectionRenderer({
    fillColor: "transparent",
    borderColor: "transparent",
    activeBorderColor: resolveCssVar("--selection-border", { fallback: "transparent" }),
    borderWidth: 2,
    activeBorderWidth: 3,
    fillHandleSize: 0,
  });
  private readonly selectionListeners = new Set<(selection: SelectionState) => void>();

  private editState = false;
  private readonly editStateListeners = new Set<(isEditing: boolean) => void>();

  private editor: CellEditorOverlay;
  private formulaBar: FormulaBarView | null = null;
  private formulaBarCompletion: FormulaBarTabCompletionController | null = null;
  private formulaEditCell: { sheetId: string; cell: CellCoord } | null = null;
  private referencePreview: { start: CellCoord; end: CellCoord } | null = null;
  private referenceHighlights: Array<{ start: CellCoord; end: CellCoord; color: string; active: boolean }> = [];
  private referenceHighlightsSource: Array<{
    range: { startRow: number; startCol: number; endRow: number; endCol: number; sheet?: string };
    color: string;
    active?: boolean;
  }> = [];
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
    this.gridMode = resolveDesktopGridMode();
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
    this.auditingCanvas = document.createElement("canvas");
    this.auditingCanvas.className = "grid-canvas";
    this.auditingCanvas.setAttribute("aria-hidden", "true");
    this.selectionCanvas = document.createElement("canvas");
    this.selectionCanvas.className = "grid-canvas";
    this.selectionCanvas.setAttribute("aria-hidden", "true");

    this.root.appendChild(this.gridCanvas);
    this.root.appendChild(this.chartLayer);
    this.root.appendChild(this.referenceCanvas);
    this.root.appendChild(this.auditingCanvas);
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
    this.vScrollbarTrack.className = "grid-scrollbar-track grid-scrollbar-track--vertical";

    this.vScrollbarThumb = document.createElement("div");
    this.vScrollbarThumb.setAttribute("aria-hidden", "true");
    this.vScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-y");
    this.vScrollbarThumb.className = "grid-scrollbar-thumb";
    this.vScrollbarTrack.appendChild(this.vScrollbarThumb);
    this.root.appendChild(this.vScrollbarTrack);

    this.hScrollbarTrack = document.createElement("div");
    this.hScrollbarTrack.setAttribute("aria-hidden", "true");
    this.hScrollbarTrack.setAttribute("data-testid", "scrollbar-track-x");
    this.hScrollbarTrack.className = "grid-scrollbar-track grid-scrollbar-track--horizontal";

    this.hScrollbarThumb = document.createElement("div");
    this.hScrollbarThumb.setAttribute("aria-hidden", "true");
    this.hScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-x");
    this.hScrollbarThumb.className = "grid-scrollbar-thumb";
    this.hScrollbarTrack.appendChild(this.hScrollbarThumb);
    this.root.appendChild(this.hScrollbarTrack);

    this.commentsPanel = this.createCommentsPanel();
    this.root.appendChild(this.commentsPanel);

    this.commentTooltip = this.createCommentTooltip();
    this.root.appendChild(this.commentTooltip);

    this.auditingLegend = this.createAuditingLegend();
    this.root.appendChild(this.auditingLegend);

    const gridCtx = this.gridCanvas.getContext("2d");
    const referenceCtx = this.referenceCanvas.getContext("2d");
    const auditingCtx = this.auditingCanvas.getContext("2d");
    const selectionCtx = this.selectionCanvas.getContext("2d");
    if (!gridCtx || !referenceCtx || !auditingCtx || !selectionCtx) {
      throw new Error("Canvas 2D context not available");
    }
    this.gridCtx = gridCtx;
    this.referenceCtx = referenceCtx;
    this.auditingCtx = auditingCtx;
    this.selectionCtx = selectionCtx;

    this.editor = new CellEditorOverlay(this.root, {
      onCommit: (commit) => {
        this.updateEditState();
        this.applyEdit(this.sheetId, commit.cell, commit.value);

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
        if (this.sharedGrid) this.syncSharedGridSelectionFromState();
        this.refresh();
        this.focus();
      },
      onCancel: () => {
        this.updateEditState();
        this.renderSelection();
        this.updateStatus();
        this.focus();
      }
    });

    this.inlineEditController = new InlineEditController({
      container: this.root,
      document: this.document,
      workbookId: opts.workbookId,
      schemaProvider: createSchemaProviderFromSearchWorkbook(this.searchWorkbook),
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

    if (this.gridMode === "shared") {
      const headerRows = 1;
      const headerCols = 1;
      this.sharedProvider = new DocumentCellProvider({
        document: this.document,
        getSheetId: () => this.sheetId,
        headerRows,
        headerCols,
        rowCount: this.limits.maxRows + headerRows,
        colCount: this.limits.maxCols + headerCols,
        showFormulas: () => this.showFormulas,
        getComputedValue: (cell) => this.getCellComputedValue(cell),
        getCommentMeta: (cellRef) => this.commentMeta.get(cellRef) ?? null
      });

      this.sharedGrid = new DesktopSharedGrid({
        container: this.root,
        provider: this.sharedProvider,
        rowCount: this.limits.maxRows + headerRows,
        colCount: this.limits.maxCols + headerCols,
        frozenRows: headerRows,
        frozenCols: headerCols,
        defaultRowHeight: this.cellHeight,
        defaultColWidth: this.cellWidth,
        enableResize: true,
        enableKeyboard: false,
        canvases: { grid: this.gridCanvas, content: this.referenceCanvas, selection: this.selectionCanvas },
        scrollbars: {
          vTrack: this.vScrollbarTrack,
          vThumb: this.vScrollbarThumb,
          hTrack: this.hScrollbarTrack,
          hThumb: this.hScrollbarThumb
        },
        callbacks: {
          onScroll: (scroll, viewport) => {
            this.scrollX = scroll.x;
            this.scrollY = scroll.y;
            this.syncSharedChartPanes(viewport);
            this.hideCommentTooltip();
            this.renderCharts(false);
            this.renderAuditing();
            this.renderSelection();
          },
          onSelectionChange: () => {
            if (this.sharedGridSelectionSyncInProgress) return;
            this.syncSelectionFromSharedGrid();
            this.updateStatus();
          },
          onSelectionRangeChange: () => {
            if (this.sharedGridSelectionSyncInProgress) return;
            this.syncSelectionFromSharedGrid();
            this.updateStatus();
          },
          onRequestCellEdit: (request) => {
            this.openEditorFromSharedGrid(request);
          },
          onAxisSizeChange: (change) => {
            this.onSharedGridAxisSizeChange(change);
          },
          onRangeSelectionStart: (range) => this.onSharedRangeSelectionStart(range),
          onRangeSelectionChange: (range) => this.onSharedRangeSelectionChange(range),
          onRangeSelectionEnd: () => this.onSharedRangeSelectionEnd(),
          onFillCommit: ({ sourceRange, targetRange, mode }) => {
            const headerRows = this.sharedHeaderRows();
            const headerCols = this.sharedHeaderCols();

            const toFillRange = (range: GridCellRange): FillEngineRange | null => {
              const startRow = Math.max(0, range.startRow - headerRows);
              const endRow = Math.max(0, range.endRow - headerRows);
              const startCol = Math.max(0, range.startCol - headerCols);
              const endCol = Math.max(0, range.endCol - headerCols);
              if (endRow <= startRow || endCol <= startCol) return null;
              return { startRow, endRow, startCol, endCol };
            };

            const source = toFillRange(sourceRange);
            const target = toFillRange(targetRange);
            if (!source || !target) return;

            applyFillCommitToDocumentController({
              document: this.document,
              sheetId: this.sheetId,
              sourceRange: source,
              targetRange: target,
              mode,
              getCellComputedValue: (row, col) => this.getCellComputedValue({ row, col }) as any
            });

            // Ensure non-grid overlays (charts, auditing) refresh after the mutation.
            this.refresh();
            this.focus();
          }
        }
      });

      // Match the legacy header sizing so existing click offsets and overlays stay aligned.
      this.sharedGrid.renderer.setColWidth(0, this.rowHeaderWidth);
      this.sharedGrid.renderer.setRowHeight(0, this.colHeaderHeight);

      // Keep legacy overlay ordering: charts above cells, selection above charts.
      this.chartLayer.style.zIndex = "2";
      this.selectionCanvas.style.zIndex = "3";

      this.initSharedChartPanes();
    }

    if (this.gridMode === "shared") {
      // Shared-grid mode uses the CanvasGridRenderer selection layer, but we still
      // need pointer movement for comment tooltips.
      this.root.addEventListener("pointermove", (e) => this.onSharedPointerMove(e), { signal: this.domAbort.signal });
      this.root.addEventListener("pointerleave", () => this.hideCommentTooltip(), { signal: this.domAbort.signal });
      this.root.addEventListener("keydown", (e) => this.onKeyDown(e), { signal: this.domAbort.signal });
    } else {
      this.root.addEventListener("pointerdown", (e) => this.onPointerDown(e), { signal: this.domAbort.signal });
      this.root.addEventListener("pointermove", (e) => this.onPointerMove(e), { signal: this.domAbort.signal });
      this.root.addEventListener("pointerup", (e) => this.onPointerUp(e), { signal: this.domAbort.signal });
      this.root.addEventListener("pointercancel", (e) => this.onPointerUp(e), { signal: this.domAbort.signal });
      this.root.addEventListener(
        "pointerleave",
        () => {
          this.hideCommentTooltip();
          this.root.style.cursor = "";
        },
        { signal: this.domAbort.signal }
      );
      this.root.addEventListener("keydown", (e) => this.onKeyDown(e), { signal: this.domAbort.signal });
      this.root.addEventListener("wheel", (e) => this.onWheel(e), { passive: false, signal: this.domAbort.signal });

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
    }

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

    this.auditingUnsubscribe = this.document.on("change", (payload: any) => {
      this.auditingCache.clear();
      this.auditingLastCellKey = null;
      if (this.auditingMode !== "off") {
        this.scheduleAuditingUpdate();
      }

      // DocumentController changes can also include sheet-level view deltas
      // (e.g. frozen panes). In shared-grid mode, frozen panes must be pushed
      // down to the CanvasGridRenderer explicitly.
      if (
        Array.isArray(payload?.sheetViewDeltas) &&
        payload.sheetViewDeltas.some((delta: any) => delta?.sheetId === this.sheetId)
      ) {
        this.syncFrozenPanes();
      }
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
          this.formulaEditCell = { sheetId: this.sheetId, cell: { ...this.selection.active } };
          this.syncSharedGridInteractionMode();
          this.updateEditState();
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
          if (this.sharedGrid) {
            this.syncSharedGridInteractionMode();
            if (range) {
              const gridRange = this.gridRangeFromDocRange({
                startRow: range.start.row,
                endRow: range.end.row,
                startCol: range.start.col,
                endCol: range.end.col
              });
              this.sharedGrid.setRangeSelection(gridRange);
            } else {
              this.sharedGrid.clearRangeSelection();
            }
          } else {
            this.renderReferencePreview();
          }
        },
        onReferenceHighlights: (highlights) => {
          this.referenceHighlightsSource = highlights;
          this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
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

            const out: Array<{ name: string; range?: string }> = [];
            for (const entry of this.searchWorkbook.names.values()) {
              const e: any = entry as any;
              const name = typeof e?.name === "string" ? (e.name as string) : "";
              if (!name) continue;
              const sheetName = typeof e?.sheetName === "string" ? (e.sheetName as string) : "";
              const range = e?.range;
              const rangeText = sheetName && range ? `${formatSheetPrefix(sheetName)}${rangeToA1(range)}` : undefined;
              out.push({ name, range: rangeText });
            }
            return out;
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
    if (this.sharedGrid) {
      // Ensure the shared renderer selection layer matches the app selection model.
      this.syncSharedGridSelectionFromState();
    }
    this.uiReady = true;
    this.editState = this.isEditing();
  }

  private dispatchViewChanged(): void {
    // Used by the Ribbon to sync pressed state for view toggles (e.g. when toggled
    // via keyboard shortcuts).
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:view-changed"));
  }

  destroy(): void {
    this.disposed = true;
    this.domAbort.abort();
    if (this.commentsDocUpdateListener) {
      this.commentsDoc.off("update", this.commentsDocUpdateListener);
      this.commentsDocUpdateListener = null;
    }
    this.formulaBarCompletion?.destroy();
    this.sharedGrid?.destroy();
    this.sharedGrid = null;
    this.sharedProvider = null;
    this.wasmUnsubscribe?.();
    this.wasmUnsubscribe = null;
    this.auditingUnsubscribe?.();
    this.auditingUnsubscribe = null;
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
      this.renderAuditing();
      this.renderSelection();
      if (renderMode === "full") this.updateStatus();
    });
  }

  focus(): void {
    this.root.focus();
  }

  isCellEditorOpen(): boolean {
    return this.editor.isOpen();
  }

  isFormulaBarEditing(): boolean {
    return Boolean(this.formulaBar?.isEditing() || this.formulaEditCell);
  }

  isEditing(): boolean {
    return this.isCellEditorOpen() || this.isFormulaBarEditing();
  }

  onEditStateChange(listener: (isEditing: boolean) => void): () => void {
    this.editStateListeners.add(listener);
    listener(this.isEditing());
    return () => {
      this.editStateListeners.delete(listener);
    };
  }

  private updateEditState(): void {
    const next = this.isEditing();
    if (next === this.editState) return;
    this.editState = next;
    for (const listener of this.editStateListeners) {
      listener(next);
    }
  }

  async clipboardCopy(): Promise<void> {
    await this.copySelectionToClipboard();
  }

  async clipboardCut(): Promise<void> {
    await this.cutSelectionToClipboard();
  }

  async clipboardPaste(): Promise<void> {
    await this.pasteClipboardToSelection();
  }

  async clipboardPasteSpecial(mode: "all" | "values" | "formulas" | "formats" = "all"): Promise<void> {
    const normalized: "all" | "values" | "formulas" | "formats" =
      mode === "values" || mode === "formulas" || mode === "formats" ? mode : "all";

    try {
      const provider = await this.getClipboardProvider();
      const content = await provider.read();
      const grid = parseClipboardContentToCellGrid(content);
      if (!grid) return;

      const rowCount = grid.length;
      const colCount = Math.max(0, ...grid.map((row) => row.length));
      if (rowCount === 0 || colCount === 0) return;

      const start = { ...this.selection.active };

      const values = Array.from({ length: rowCount }, (_, r) => {
        const row = grid[r] ?? [];
        return Array.from({ length: colCount }, (_, c) => {
          const cell = (row[c] ?? null) as any;

          if (normalized === "values") return cell?.value ?? null;

          if (normalized === "formulas") {
            return cell?.formula != null ? { formula: cell.formula } : cell?.value ?? null;
          }

          if (normalized === "formats") return { format: cell?.format ?? null };

          // normalized === "all"
          if (cell?.formula != null) return { formula: cell.formula, value: null, format: cell?.format ?? null };
          return { value: cell?.value ?? null, format: cell?.format ?? null };
        });
      });

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

  clearContents(): void {
    this.clearSelectionContents();
    this.refresh();
    this.focus();
  }

  async whenIdle(): Promise<void> {
    // `wasmSyncPromise` and `auditingIdlePromise` are growing promise chains (see `enqueueWasmSync`
    // and `scheduleAuditingUpdate`). While awaiting them, additional work can be appended
    // (e.g. a batched `setCells` update queued after an engine apply). Loop until both
    // chains stabilize so callers (notably Playwright) can reliably await the
    // "no more background work pending" condition.
    while (true) {
      const wasm = this.wasmSyncPromise;
      const auditing = this.auditingIdlePromise;
      await Promise.all([this.idle.whenIdle(), wasm.catch(() => {}), auditing.catch(() => {})]);
      if (this.wasmSyncPromise === wasm && this.auditingIdlePromise === auditing) return;
    }
  }

  getRecalcCount(): number {
    return this.engine.recalcCount;
  }

  getDocument(): DocumentController {
    return this.document;
  }

  undo(): boolean {
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;
    return this.applyUndoRedo("undo");
  }

  redo(): boolean {
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;
    return this.applyUndoRedo("redo");
  }

  getUndoRedoState(): UndoRedoState {
    return {
      canUndo: Boolean(this.document.canUndo),
      canRedo: Boolean(this.document.canRedo),
      undoLabel: (this.document.undoLabel as string | null) ?? null,
      redoLabel: (this.document.redoLabel as string | null) ?? null,
    };
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

  getGridMode(): DesktopGridMode {
    return this.gridMode;
  }

  supportsZoom(): boolean {
    return this.sharedGrid != null;
  }

  getZoom(): number {
    return this.sharedGrid ? this.sharedGrid.renderer.getZoom() : 1;
  }

  setZoom(nextZoom: number): void {
    if (!this.sharedGrid) return;
    this.sharedGrid.renderer.setZoom(nextZoom);
    // Document sheet views store base axis sizes at zoom=1; apply zoom scaling and resync.
    this.syncSharedGridAxisSizesFromDocument();
    // Re-emit scroll so overlays (charts, auditing, etc) can re-layout at the new zoom level.
    const scroll = this.sharedGrid.getScroll();
    this.sharedGrid.scrollTo(scroll.x, scroll.y);
  }

  getShowFormulas(): boolean {
    return this.showFormulas;
  }

  setShowFormulas(enabled: boolean): void {
    if (enabled === this.showFormulas) return;
    this.showFormulas = enabled;
    this.sharedProvider?.invalidateAll();
    this.refresh();
    this.dispatchViewChanged();
  }

  toggleShowFormulas(): void {
    this.setShowFormulas(!this.showFormulas);
  }

  /**
   * Test-only helper: returns the viewport-relative rect for a cell address.
   */
  getCellRectA1(a1: string): { x: number; y: number; width: number; height: number } | null {
    const cell = parseA1(a1);
    const rect = this.getCellRect(cell);
    if (!rect) return null;
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  }

  /**
   * Returns grid renderer perf stats (shared grid only).
   */
  getGridPerfStats(): unknown {
    if (!this.sharedGrid) return null;
    const stats = this.sharedGrid.getPerfStats();
    return {
      enabled: stats.enabled,
      lastFrameMs: stats.lastFrameMs,
      cellsPainted: stats.cellsPainted,
      cellFetches: stats.cellFetches,
      dirtyRects: { ...stats.dirtyRects },
      blitUsed: stats.blitUsed
    };
  }

  setGridPerfStatsEnabled(enabled: boolean): void {
    this.sharedGrid?.setPerfStatsEnabled(enabled);
    this.dispatchViewChanged();
  }

  getFrozen(): { frozenRows: number; frozenCols: number } {
    const view = this.document.getSheetView(this.sheetId) as { frozenRows?: number; frozenCols?: number } | null;
    const normalize = (value: unknown, max: number): number => {
      const num = Number(value);
      if (!Number.isFinite(num)) return 0;
      return Math.max(0, Math.min(Math.trunc(num), max));
    };
    return {
      frozenRows: normalize(view?.frozenRows, this.limits.maxRows),
      frozenCols: normalize(view?.frozenCols, this.limits.maxCols),
    };
  }

  private syncFrozenPanes(): void {
    const { frozenRows, frozenCols } = this.getFrozen();
    if (this.sharedGrid) {
      // Shared-grid mode uses frozen panes for headers + user freezes.
      const headerRows = 1;
      const headerCols = 1;
      this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
      this.syncSharedGridAxisSizesFromDocument();
      // Force scrollbars + overlay layers to re-measure frozen extents.
      this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
      return;
    }

    this.clampScroll();
    this.syncScrollbars();
    this.refresh();
  }

  private syncSharedGridAxisSizesFromDocument(): void {
    if (!this.sharedGrid) return;

    const view = this.document.getSheetView(this.sheetId) as {
      colWidths?: Record<string, number>;
      rowHeights?: Record<string, number>;
    } | null;

    const zoom = this.sharedGrid.renderer.getZoom();
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();

    const nextCols = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.colWidths ?? {})) {
      const col = Number(key);
      if (!Number.isInteger(col) || col < 0) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      nextCols.set(col, size);
    }

    const nextRows = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.rowHeights ?? {})) {
      const row = Number(key);
      if (!Number.isInteger(row) || row < 0) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      nextRows.set(row, size);
    }

    const allCols = new Set<number>();
    for (const col of this.sharedGridAxisCols) allCols.add(col);
    for (const col of nextCols.keys()) allCols.add(col);
    for (const col of allCols) {
      const base = nextCols.get(col);
      const gridCol = col + headerCols;
      if (base === undefined) this.sharedGrid.renderer.resetColWidth(gridCol);
      else this.sharedGrid.renderer.setColWidth(gridCol, base * zoom);
    }

    const allRows = new Set<number>();
    for (const row of this.sharedGridAxisRows) allRows.add(row);
    for (const row of nextRows.keys()) allRows.add(row);
    for (const row of allRows) {
      const base = nextRows.get(row);
      const gridRow = row + headerRows;
      if (base === undefined) this.sharedGrid.renderer.resetRowHeight(gridRow);
      else this.sharedGrid.renderer.setRowHeight(gridRow, base * zoom);
    }

    this.sharedGridAxisCols.clear();
    for (const col of nextCols.keys()) this.sharedGridAxisCols.add(col);
    this.sharedGridAxisRows.clear();
    for (const row of nextRows.keys()) this.sharedGridAxisRows.add(row);
  }

  freezePanes(): void {
    const active = this.selection.active;
    this.document.setFrozen(this.sheetId, active.row, active.col, { label: "Freeze Panes" });
  }

  freezeTopRow(): void {
    this.document.setFrozen(this.sheetId, 1, 0, { label: "Freeze Top Row" });
  }

  freezeFirstColumn(): void {
    this.document.setFrozen(this.sheetId, 0, 1, { label: "Freeze First Column" });
  }

  unfreezePanes(): void {
    this.document.setFrozen(this.sheetId, 0, 0, { label: "Unfreeze Panes" });
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
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      this.syncFrozenPanes();
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

    let postInitHydrate: Promise<void> | null = null;

    // Include WASM initialization in the idle tracker. This ensures e2e tests (and any UI that
    // awaits `whenIdle()`) don't race the initial engine hydration.
    const prior = this.wasmSyncPromise;
    this.wasmSyncPromise = prior
      .catch(() => {})
      .then(async () => {
        // Another init call may have completed while we were waiting on the prior promise.
        if (this.wasmEngine) return;

        const env = (import.meta as any)?.env as Record<string, unknown> | undefined;
        const wasmModuleUrl =
          typeof env?.VITE_FORMULA_WASM_MODULE_URL === "string" ? env.VITE_FORMULA_WASM_MODULE_URL : undefined;
        const wasmBinaryUrl =
          typeof env?.VITE_FORMULA_WASM_BINARY_URL === "string" ? env.VITE_FORMULA_WASM_BINARY_URL : undefined;

        let engine: EngineClient | null = null;
        try {
          engine = createEngineClient({ wasmModuleUrl, wasmBinaryUrl });
          let changedDuringInit = false;
          const unsubscribeInit = this.document.on("change", () => {
            changedDuringInit = true;
          });
          try {
            await engine.init();

            // `engineHydrateFromDocument` is relatively expensive, but we need to ensure we don't miss
            // edits that happen while the WASM worker is booting. If the user edits the document
            // between hydrating the worker and subscribing to incremental deltas, the engine can get
            // permanently out of sync (until a full reload). Track any DocumentController changes
            // during the hydrate step and retry once to converge.
            //
            // This is especially important for fast e2e runs that start editing immediately after
            // navigation.
            for (let attempt = 0; attempt < 2; attempt += 1) {
              changedDuringInit = false;
              this.computedValues.clear();
              const changes = await engineHydrateFromDocument(engine, this.document);
              this.applyComputedChanges(changes);
              if (!changedDuringInit) break;
            }
          } finally {
            unsubscribeInit();
          }

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

           // `initWasmEngine` runs asynchronously and can overlap with early user edits (or e2e
           // interactions) before the `document.on("change")` listener is installed. If the
           // DocumentController changed while `engineHydrateFromDocument` was in-flight, those
           // deltas could be missed, leaving the worker with an incomplete view of inputs.
           //
           // Re-hydrate once through the serialized WASM queue to guarantee the worker matches the
           // latest DocumentController state before incremental deltas begin flowing.
           //
           // Note: do not `await` inside this init chain (it would deadlock by waiting on the
           // promise chain we're currently building).
           postInitHydrate = this.enqueueWasmSync(async (worker) => {
             const changes = await engineHydrateFromDocument(worker, this.document);
             this.applyComputedChanges(changes);
           });
         } catch {
           // Ignore initialization failures (e.g. missing WASM bundle).
           engine?.terminate();
           this.wasmEngine = null;
           this.wasmUnsubscribe?.();
           this.wasmUnsubscribe = null;
           this.computedValues.clear();
         }
      })
      .catch(() => {
        // Ignore WASM init failures; the app continues to function using the in-process mock engine.
      });

    await this.wasmSyncPromise;
    if (postInitHydrate) {
      try {
        await postInitHydrate;
      } catch {
        // ignore
      }
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
    this.referencePreview = null;
    this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
    if (this.sharedGrid) {
      const { frozenRows, frozenCols } = this.getFrozen();
      const headerRows = 1;
      const headerCols = 1;
      this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
      this.syncSharedGridAxisSizesFromDocument();
      this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
    } else {
      this.clampScroll();
      this.syncScrollbars();
    }
    this.renderGrid();
    this.renderCharts(true);
    this.renderReferencePreview();
    if (this.sharedGrid) {
      // Switching sheets updates the provider data source but does not emit document
      // changes. Force a full redraw so the CanvasGridRenderer pulls from the new
      // sheet's data.
      this.sharedProvider?.invalidateAll();
      this.syncSharedGridSelectionFromState();
    }
    this.renderReferencePreview();
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
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      if (this.sharedGrid) {
        const { frozenRows, frozenCols } = this.getFrozen();
        const headerRows = 1;
        const headerCols = 1;
        this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
        this.syncSharedGridAxisSizesFromDocument();
        this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
      } else {
        this.clampScroll();
        this.syncScrollbars();
      }
      this.renderGrid();
      this.renderCharts(true);
      this.sharedProvider?.invalidateAll();
      sheetChanged = true;
    }
    this.selection = setActiveCell(this.selection, { row: target.row, col: target.col }, this.limits);
    this.ensureActiveCellVisible();
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
    else if (didScroll) this.ensureViewportMappingCurrent();
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
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      if (this.sharedGrid) {
        const { frozenRows, frozenCols } = this.getFrozen();
        const headerRows = 1;
        const headerCols = 1;
        this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
        this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
      } else {
        this.clampScroll();
        this.syncScrollbars();
      }
      this.renderGrid();
      this.renderCharts(true);
      this.sharedProvider?.invalidateAll();
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
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
    else if (didScroll) this.ensureViewportMappingCurrent();
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

  /**
   * Compute Excel-like status bar stats (Sum / Average / Count) for the current selection.
   *
   * Performance note: this is implemented by iterating only the *stored* cells in the
   * DocumentController's sparse cell map (not every coordinate in the rectangular ranges).
   */
  getSelectionSummary(): SpreadsheetSelectionSummary {
    const ranges = this.selection.ranges;
    let countNonEmpty = 0;
    let numericCount = 0;
    let numericSum = 0;

    // Iterate only stored cells (value/formula/format-only), then filter by selection.
    this.document.forEachCellInSheet(this.sheetId, ({ row, col, cell }) => {
      // Ignore cells outside the current selection ranges.
      let inSelection = false;
      for (const r of ranges) {
        if (row >= r.startRow && row <= r.endRow && col >= r.startCol && col <= r.endCol) {
          inSelection = true;
          break;
        }
      }
      if (!inSelection) return;

      // Ignore format-only cells (styleId-only).
      const hasContent = cell.value != null || cell.formula != null;
      if (!hasContent) return;

      countNonEmpty += 1;

      // Sum/average operate on numeric values only (computed values for formulas).
      if (cell.formula != null) {
        const computed = this.getCellComputedValue({ row, col });
        if (typeof computed === "number" && Number.isFinite(computed)) {
          numericCount += 1;
          numericSum += computed;
        }
        return;
      }

      if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
        numericCount += 1;
        numericSum += cell.value;
      }
    });

    const sum = numericCount > 0 ? numericSum : null;
    const average = numericCount > 0 ? numericSum / numericCount : null;

    return {
      sum,
      average,
      count: countNonEmpty,
      numericCount,
      countNonEmpty,
    };
  }

  getActiveCell(): CellCoord {
    return { ...this.selection.active };
  }

  /**
   * Return the active cell's bounding box in viewport coordinates.
   *
   * This is useful for anchoring floating UI (e.g. context menus) to the active cell
   * without changing selection.
   */
  getActiveCellRect(): { x: number; y: number; width: number; height: number } | null {
    if (!this.sharedGrid) {
      this.ensureViewportMappingCurrent();
    }

    const rect = this.getCellRect(this.selection.active);
    if (!rect) return null;

    const rootRect = (this.sharedGrid ? this.selectionCanvas : this.root).getBoundingClientRect();
    return {
      x: rootRect.left + rect.x,
      y: rootRect.top + rect.y,
      width: rect.width,
      height: rect.height,
    };
  }

  /**
   * Hit-test the grid and return the document cell (0-based row/col) under the
   * provided client (viewport) coordinates.
   *
   * This is intentionally limited to sheet cells for now (it returns null when
   * the point lands on row/column headers).
   */
  pickCellAtClientPoint(clientX: number, clientY: number): CellCoord | null {
    const rootRect = this.root.getBoundingClientRect();
    const x = clientX - rootRect.left;
    const y = clientY - rootRect.top;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
    if (x < 0 || y < 0 || x > rootRect.width || y > rootRect.height) return null;

    // Only treat hits in the cell grid (exclude row/col headers).
    if (!this.sharedGrid) {
      if (x < this.rowHeaderWidth || y < this.colHeaderHeight) return null;
      return this.cellFromPoint(x, y);
    }

    // Shared grid uses its own internal coordinate space anchored on the
    // selection canvas.
    const canvasRect = this.selectionCanvas.getBoundingClientRect();
    const vx = clientX - canvasRect.left;
    const vy = clientY - canvasRect.top;
    if (!Number.isFinite(vx) || !Number.isFinite(vy)) return null;
    if (vx < 0 || vy < 0 || vx > canvasRect.width || vy > canvasRect.height) return null;

    const picked = this.sharedGrid.renderer.pickCellAt(vx, vy);
    if (!picked) return null;

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    if (picked.row < headerRows || picked.col < headerCols) return null;

    return { row: picked.row - headerRows, col: picked.col - headerCols };
  }

  subscribeSelection(listener: (selection: SelectionState) => void): () => void {
    this.selectionListeners.add(listener);
    listener(this.selection);
    return () => this.selectionListeners.delete(listener);
  }

  private sharedHeaderRows(): number {
    return this.sharedGrid ? 1 : 0;
  }

  private sharedHeaderCols(): number {
    return this.sharedGrid ? 1 : 0;
  }

  private docCellFromGridCell(cell: { row: number; col: number }): CellCoord {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return { row: Math.max(0, cell.row - headerRows), col: Math.max(0, cell.col - headerCols) };
  }

  private gridCellFromDocCell(cell: CellCoord): { row: number; col: number } {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return { row: cell.row + headerRows, col: cell.col + headerCols };
  }

  private gridRangeFromDocRange(range: Range): GridCellRange {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return {
      startRow: range.startRow + headerRows,
      endRow: range.endRow + headerRows + 1,
      startCol: range.startCol + headerCols,
      endCol: range.endCol + headerCols + 1
    };
  }

  private docRangeFromGridRange(range: GridCellRange): Range {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return {
      startRow: Math.max(0, range.startRow - headerRows),
      endRow: Math.max(0, range.endRow - headerRows - 1),
      startCol: Math.max(0, range.startCol - headerCols),
      endCol: Math.max(0, range.endCol - headerCols - 1)
    };
  }

  private syncSelectionFromSharedGrid(): void {
    if (!this.sharedGrid) return;
    const selection = this.sharedGrid.renderer.getSelection();
    const ranges = this.sharedGrid.renderer.getSelectionRanges();
    const activeIndex = this.sharedGrid.renderer.getActiveSelectionIndex();

    if (!selection || ranges.length === 0) {
      this.selection = createSelection({ row: 0, col: 0 }, this.limits);
      return;
    }

    const docActive = this.docCellFromGridCell(selection);
    const docRanges = ranges.map((r) => this.docRangeFromGridRange(r));
    const activeRangeIndex = Math.max(0, Math.min(docRanges.length - 1, activeIndex));
    const anchor = { ...docActive };

    this.selection = buildSelection(
      {
        ranges: docRanges,
        active: docActive,
        anchor,
        activeRangeIndex
      },
      this.limits
    );
  }

  private syncSharedGridSelectionFromState(): void {
    if (!this.sharedGrid) return;
    const gridRanges = this.selection.ranges.map((r) => this.gridRangeFromDocRange(r));
    const gridActive = this.gridCellFromDocCell(this.selection.active);
    this.sharedGridSelectionSyncInProgress = true;
    try {
      this.sharedGrid.setSelectionRanges(gridRanges, { activeIndex: this.selection.activeRangeIndex, activeCell: gridActive });
      this.scrollX = this.sharedGrid.getScroll().x;
      this.scrollY = this.sharedGrid.getScroll().y;
    } finally {
      this.sharedGridSelectionSyncInProgress = false;
    }
  }

  private openEditorFromSharedGrid(request: { row: number; col: number; initialKey?: string }): void {
    if (!this.sharedGrid) return;
    if (this.editor.isOpen()) return;
    const docCell = this.docCellFromGridCell({ row: request.row, col: request.col });
    const rect = this.sharedGrid.getCellRect(request.row, request.col);
    if (!rect) return;
    const initialValue = request.initialKey ?? this.getCellInputText(docCell);
    this.editor.open(docCell, rect, initialValue, { cursor: "end" });
    this.updateEditState();
  }

  private onSharedGridAxisSizeChange(change: GridAxisSizeChange): void {
    if (!this.sharedGrid) return;

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;

    if (change.kind === "col") {
      const docCol = change.index - headerCols;
      if (docCol < 0) return;
      const label = change.source === "autoFit" ? "Autofit Column Width" : "Resize Column";
      if (isDefault) {
        this.document.resetColWidth(this.sheetId, docCol, { label });
        this.sharedGridAxisCols.delete(docCol);
      } else {
        this.document.setColWidth(this.sheetId, docCol, baseSize, { label });
        this.sharedGridAxisCols.add(docCol);
      }
      return;
    }

    const docRow = change.index - headerRows;
    if (docRow < 0) return;
    const label = change.source === "autoFit" ? "Autofit Row Height" : "Resize Row";
    if (isDefault) {
      this.document.resetRowHeight(this.sheetId, docRow, { label });
      this.sharedGridAxisRows.delete(docRow);
    } else {
      this.document.setRowHeight(this.sheetId, docRow, baseSize, { label });
      this.sharedGridAxisRows.add(docRow);
    }
  }

  private syncSharedGridInteractionMode(): void {
    if (!this.sharedGrid) return;
    const mode = this.formulaBar?.isFormulaEditing() ? "rangeSelection" : "default";
    this.sharedGrid.setInteractionMode(mode);
  }

  private onSharedRangeSelectionStart(range: GridCellRange): void {
    if (!this.formulaBar) return;
    this.syncSharedGridInteractionMode();
    const docRange = this.docRangeFromGridRange(range);
    const rangeSheetId = this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
    this.formulaBar.beginRangeSelection({
      start: { row: docRange.startRow, col: docRange.startCol },
      end: { row: docRange.endRow, col: docRange.endCol }
    }, rangeSheetId);
  }

  private onSharedRangeSelectionChange(range: GridCellRange): void {
    if (!this.formulaBar) return;
    this.syncSharedGridInteractionMode();
    const docRange = this.docRangeFromGridRange(range);
    const rangeSheetId = this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
    this.formulaBar.updateRangeSelection({
      start: { row: docRange.startRow, col: docRange.startCol },
      end: { row: docRange.endRow, col: docRange.endCol }
    }, rangeSheetId);
  }

  private onSharedRangeSelectionEnd(): void {
    if (!this.formulaBar) return;
    this.formulaBar.endRangeSelection();
    this.formulaBar.focus();
  }

  private onSharedPointerMove(e: PointerEvent): void {
    if (!this.sharedGrid) return;
    if (this.commentsPanelVisible) {
      this.hideCommentTooltip();
      return;
    }
    if (this.editor.isOpen()) {
      this.hideCommentTooltip();
      return;
    }
    if (e.buttons !== 0) {
      this.hideCommentTooltip();
      return;
    }

    const rootRect = this.root.getBoundingClientRect();
    const x = e.clientX - rootRect.left;
    const y = e.clientY - rootRect.top;
    if (x < 0 || y < 0 || x > rootRect.width || y > rootRect.height) {
      this.hideCommentTooltip();
      return;
    }

    const canvasRect = this.selectionCanvas.getBoundingClientRect();
    const vx = e.clientX - canvasRect.left;
    const vy = e.clientY - canvasRect.top;
    const picked = this.sharedGrid.renderer.pickCellAt(vx, vy);
    if (!picked) {
      this.hideCommentTooltip();
      return;
    }

    const docCell = this.docCellFromGridCell(picked);
    const cellRef = cellToA1(docCell);
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
    this.commentTooltip.classList.add("comment-tooltip--visible");
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

  getFillHandleRect(): { x: number; y: number; width: number; height: number } | null {
    if (this.formulaBar?.isFormulaEditing()) return null;
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.getFillHandleRect();
    }
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

  getAuditingHighlights(): {
    mode: "off" | AuditingMode;
    transitive: boolean;
    precedents: string[];
    dependents: string[];
    errors: { precedents: string | null; dependents: string | null };
  } {
    return {
      mode: this.auditingMode,
      transitive: this.auditingTransitive,
      precedents: Array.from(this.auditingHighlights.precedents).sort(),
      dependents: Array.from(this.auditingHighlights.dependents).sort(),
      errors: { ...this.auditingErrors },
    };
  }

  toggleAuditingPrecedents(): void {
    this.toggleAuditingComponent("precedents");
  }

  toggleAuditingDependents(): void {
    this.toggleAuditingComponent("dependents");
  }

  toggleAuditingTransitive(): void {
    this.auditingTransitive = !this.auditingTransitive;
    this.updateAuditingLegend();
    this.scheduleAuditingUpdate();
  }

  clearAuditing(): void {
    this.auditingMode = "off";
    this.auditingHighlights = { precedents: new Set(), dependents: new Set() };
    this.auditingErrors = { precedents: null, dependents: null };
    this.updateAuditingLegend();
    this.renderAuditing();
  }

  toggleCommentsPanel(): void {
    this.commentsPanelVisible = !this.commentsPanelVisible;
    this.commentsPanel.classList.toggle("comments-panel--visible", this.commentsPanelVisible);
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
    panel.className = "comments-panel";

    panel.addEventListener("pointerdown", (e) => e.stopPropagation());
    panel.addEventListener("dblclick", (e) => e.stopPropagation());
    panel.addEventListener("keydown", (e) => e.stopPropagation());

    const header = document.createElement("div");
    header.className = "comments-panel__header";

    const title = document.createElement("div");
    title.textContent = t("comments.title");
    title.className = "comments-panel__title";

    const closeButton = document.createElement("button");
    closeButton.textContent = "×";
    closeButton.type = "button";
    closeButton.className = "comments-panel__close-button";
    closeButton.setAttribute("aria-label", t("comments.closePanel"));
    closeButton.addEventListener("click", () => this.toggleCommentsPanel());

    header.appendChild(title);
    header.appendChild(closeButton);
    panel.appendChild(header);

    this.commentsPanelCell = document.createElement("div");
    this.commentsPanelCell.dataset.testid = "comments-active-cell";
    this.commentsPanelCell.className = "comments-panel__active-cell";
    panel.appendChild(this.commentsPanelCell);

    this.commentsPanelThreads = document.createElement("div");
    this.commentsPanelThreads.className = "comments-panel__threads";
    panel.appendChild(this.commentsPanelThreads);

    const footer = document.createElement("div");
    footer.className = "comments-panel__footer";

    this.newCommentInput = document.createElement("input");
    this.newCommentInput.dataset.testid = "new-comment-input";
    this.newCommentInput.type = "text";
    this.newCommentInput.placeholder = t("comments.new.placeholder");
    this.newCommentInput.className = "comments-panel__new-comment-input";

    const submit = document.createElement("button");
    submit.dataset.testid = "submit-comment";
    submit.textContent = t("comments.new.submit");
    submit.type = "button";
    submit.className = "comments-panel__submit-button";
    submit.addEventListener("click", () => this.submitNewComment());

    footer.appendChild(this.newCommentInput);
    footer.appendChild(submit);
    panel.appendChild(footer);

    return panel;
  }

  private createCommentTooltip(): HTMLDivElement {
    const tooltip = document.createElement("div");
    tooltip.dataset.testid = "comment-tooltip";
    tooltip.className = "comment-tooltip";
    return tooltip;
  }

  private createAuditingLegend(): HTMLDivElement {
    const legend = document.createElement("div");
    legend.dataset.testid = "auditing-legend";
    legend.className = "auditing-legend";
    legend.hidden = true;
    return legend;
  }

  private updateAuditingLegend(): void {
    const legend = this.auditingLegend;

    if (this.auditingMode === "off") {
      legend.hidden = true;
      legend.replaceChildren();
      return;
    }

    const wantsPrecedents = this.auditingMode === "precedents" || this.auditingMode === "both";
    const wantsDependents = this.auditingMode === "dependents" || this.auditingMode === "both";

    legend.replaceChildren();

    const makeModeItem = (opts: { kind: "precedents" | "dependents"; label: string }): HTMLElement => {
      const item = document.createElement("span");
      item.className = "auditing-legend__item";

      const swatch = document.createElement("span");
      swatch.className = `auditing-legend__swatch auditing-legend__swatch--${opts.kind}`;

      const text = document.createElement("span");
      text.textContent = opts.label;

      item.appendChild(swatch);
      item.appendChild(text);
      return item;
    };

    const transitive = this.auditingTransitive ? "Transitive" : "Direct";

    const errors: string[] = [];
    if (wantsPrecedents && this.auditingErrors.precedents) errors.push(`Precedents: ${this.auditingErrors.precedents}`);
    if (wantsDependents && this.auditingErrors.dependents) errors.push(`Dependents: ${this.auditingErrors.dependents}`);

    const row = document.createElement("div");
    row.className = "auditing-legend__row";

    const modeRow = document.createElement("div");
    modeRow.className = "auditing-legend__modes";
    if (wantsPrecedents) modeRow.appendChild(makeModeItem({ kind: "precedents", label: "Precedents" }));
    if (wantsDependents) modeRow.appendChild(makeModeItem({ kind: "dependents", label: "Dependents" }));
    row.appendChild(modeRow);

    const transitiveEl = document.createElement("span");
    transitiveEl.className = "auditing-legend__transitive";
    transitiveEl.textContent = `(${transitive})`;
    row.appendChild(transitiveEl);

    legend.appendChild(row);

    if (errors.length > 0) {
      const errorsEl = document.createElement("div");
      errorsEl.className = "auditing-legend__errors";
      errorsEl.textContent = errors.join("\n");
      legend.appendChild(errorsEl);
    }

    legend.hidden = false;
  }

  private hideCommentTooltip(): void {
    this.commentTooltip.classList.remove("comment-tooltip--visible");
  }

  private reindexCommentCells(): void {
    this.commentCells.clear();
    this.commentMeta.clear();
    for (const comment of this.commentManager.listAll()) {
      this.commentCells.add(comment.cellRef);
      const existing = this.commentMeta.get(comment.cellRef);
      if (!existing) {
        this.commentMeta.set(comment.cellRef, { resolved: Boolean(comment.resolved) });
      } else {
        // Treat the cell as resolved only if all comment threads on the cell are resolved.
        existing.resolved = existing.resolved && Boolean(comment.resolved);
      }
    }

    // The shared renderer caches cell metadata, so comment indicator updates require a provider invalidation.
    this.sharedProvider?.invalidateAll();
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
      empty.className = "comments-panel__empty";
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
    container.className = "comment-thread";

    const header = document.createElement("div");
    header.className = "comment-thread__header";

    const author = document.createElement("div");
    author.textContent = comment.author.name || t("presence.anonymous");
    author.className = "comment-thread__author";

    const resolve = document.createElement("button");
    resolve.dataset.testid = "resolve-comment";
    resolve.textContent = comment.resolved ? t("comments.unresolve") : t("comments.resolve");
    resolve.type = "button";
    resolve.className = "comment-thread__resolve-button";
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
    body.className = "comment-thread__body";
    container.appendChild(body);

    for (const reply of comment.replies) {
      const replyEl = document.createElement("div");
      replyEl.className = "comment-thread__reply";

      const replyAuthor = document.createElement("div");
      replyAuthor.textContent = reply.author.name || t("presence.anonymous");
      replyAuthor.className = "comment-thread__reply-author";

      const replyBody = document.createElement("div");
      replyBody.textContent = reply.content;
      replyBody.className = "comment-thread__reply-body";

      replyEl.appendChild(replyAuthor);
      replyEl.appendChild(replyBody);
      container.appendChild(replyEl);
    }

    const replyRow = document.createElement("div");
    replyRow.className = "comment-thread__reply-row";

    const replyInput = document.createElement("input");
    replyInput.dataset.testid = "reply-input";
    replyInput.type = "text";
    replyInput.placeholder = t("comments.reply.placeholder");
    replyInput.className = "comment-thread__reply-input";

    const submitReply = document.createElement("button");
    submitReply.dataset.testid = "submit-reply";
    submitReply.textContent = t("comments.reply.send");
    submitReply.type = "button";
    submitReply.className = "comment-thread__submit-reply-button";
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

    if (this.sharedGrid) {
      // The shared grid owns the main canvas layers, but we still render auditing overlays
      // on a separate canvas (and keep chart DOM overlays positioned relative to the frozen headers).
      this.auditingCanvas.width = Math.floor(this.width * this.dpr);
      this.auditingCanvas.height = Math.floor(this.height * this.dpr);
      this.auditingCanvas.style.width = `${this.width}px`;
      this.auditingCanvas.style.height = `${this.height}px`;
      this.auditingCtx.setTransform(1, 0, 0, 1, 0, 0);
      this.auditingCtx.scale(this.dpr, this.dpr);

       this.sharedGrid.resize(this.width, this.height, this.dpr);
       const viewport = this.sharedGrid.renderer.scroll.getViewportState();
      this.syncSharedChartPanes(viewport);

      // Keep our legacy scroll coordinates in sync for chart positioning helpers.
      const scroll = this.sharedGrid.getScroll();
      this.scrollX = scroll.x;
      this.scrollY = scroll.y;

      this.renderCharts(true);
      this.renderAuditing();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    for (const canvas of [this.gridCanvas, this.referenceCanvas, this.auditingCanvas, this.selectionCanvas]) {
      canvas.width = Math.floor(this.width * this.dpr);
      canvas.height = Math.floor(this.height * this.dpr);
      canvas.style.width = `${this.width}px`;
      canvas.style.height = `${this.height}px`;
    }

    // Reset transforms and apply DPR scaling so drawing code uses CSS pixels.
    for (const ctx of [this.gridCtx, this.referenceCtx, this.auditingCtx, this.selectionCtx]) {
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
    this.renderAuditing();
    this.renderSelection();
    this.updateStatus();
  }

  private renderGrid(): void {
    if (this.sharedGrid) {
      // Shared-grid mode renders via CanvasGridRenderer. The legacy renderer must
      // not clear or paint over the canvases.
      return;
    }

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

    const frozenWidth = Math.min(viewportWidth, this.frozenWidth);
    const frozenHeight = Math.min(viewportHeight, this.frozenHeight);
    const scrollableWidth = Math.max(0, viewportWidth - frozenWidth);
    const scrollableHeight = Math.max(0, viewportHeight - frozenHeight);

    const frozenRows = this.frozenVisibleRows;
    const frozenCols = this.frozenVisibleCols;
    const scrollRows = this.visibleRows;
    const scrollCols = this.visibleCols;

    const startXScroll = originX + this.visibleColStart * this.cellWidth - this.scrollX;
    const startYScroll = originY + this.visibleRowStart * this.cellHeight - this.scrollY;

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

    const renderCellRegion = (options: {
      clipX: number;
      clipY: number;
      clipWidth: number;
      clipHeight: number;
      rows: number[];
      cols: number[];
      startX: number;
      startY: number;
    }): void => {
      const { clipX, clipY, clipWidth, clipHeight, rows, cols, startX, startY } = options;
      if (clipWidth <= 0 || clipHeight <= 0) return;
      if (rows.length === 0 || cols.length === 0) return;

      const endX = startX + cols.length * this.cellWidth;
      const endY = startY + rows.length * this.cellHeight;

      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipWidth, clipHeight);
      ctx.clip();

      // Grid lines for this quadrant.
      for (let r = 0; r <= rows.length; r++) {
        const y = startY + r * this.cellHeight + 0.5;
        ctx.beginPath();
        ctx.moveTo(startX, y);
        ctx.lineTo(endX, y);
        ctx.stroke();
      }

      for (let c = 0; c <= cols.length; c++) {
        const x = startX + c * this.cellWidth + 0.5;
        ctx.beginPath();
        ctx.moveTo(x, startY);
        ctx.lineTo(x, endY);
        ctx.stroke();
      }

      for (let visualRow = 0; visualRow < rows.length; visualRow++) {
        const row = rows[visualRow]!;
        for (let visualCol = 0; visualCol < cols.length; visualCol++) {
          const col = cols[visualCol]!;
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
      for (let visualRow = 0; visualRow < rows.length; visualRow++) {
        const row = rows[visualRow]!;
        for (let visualCol = 0; visualCol < cols.length; visualCol++) {
          const col = cols[visualCol]!;
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
    };

    // --- Data quadrants ---
    renderCellRegion({
      clipX: originX,
      clipY: originY,
      clipWidth: frozenWidth,
      clipHeight: frozenHeight,
      rows: frozenRows,
      cols: frozenCols,
      startX: originX,
      startY: originY
    });

    renderCellRegion({
      clipX: originX + frozenWidth,
      clipY: originY,
      clipWidth: scrollableWidth,
      clipHeight: frozenHeight,
      rows: frozenRows,
      cols: scrollCols,
      startX: startXScroll,
      startY: originY
    });

    renderCellRegion({
      clipX: originX,
      clipY: originY + frozenHeight,
      clipWidth: frozenWidth,
      clipHeight: scrollableHeight,
      rows: scrollRows,
      cols: frozenCols,
      startX: originX,
      startY: startYScroll
    });

    renderCellRegion({
      clipX: originX + frozenWidth,
      clipY: originY + frozenHeight,
      clipWidth: scrollableWidth,
      clipHeight: scrollableHeight,
      rows: scrollRows,
      cols: scrollCols,
      startX: startXScroll,
      startY: startYScroll
    });

    // Freeze lines.
    ctx.save();
    ctx.strokeStyle = resolveCssVar("--formula-grid-freeze-line", {
      fallback: resolveCssVar("--grid-line", { fallback: "CanvasText" })
    });
    ctx.lineWidth = 2;
    if (frozenWidth > 0 && frozenWidth < viewportWidth) {
      const x = originX + frozenWidth + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, this.height);
      ctx.stroke();
    }
    if (frozenHeight > 0 && frozenHeight < viewportHeight) {
      const y = originY + frozenHeight + 0.5;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(this.width, y);
      ctx.stroke();
    }
    ctx.restore();

    // Header labels.
    ctx.fillStyle = resolveCssVar("--text-primary", { fallback: "CanvasText" });
    ctx.font = "12px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // Column header labels: frozen columns (pinned).
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX, 0, frozenWidth, originY);
    ctx.clip();
    for (let visualCol = 0; visualCol < frozenCols.length; visualCol++) {
      const colIndex = frozenCols[visualCol]!;
      ctx.fillText(colToName(colIndex), originX + visualCol * this.cellWidth + this.cellWidth / 2, originY / 2);
    }
    ctx.restore();

    // Column header labels: scrollable columns.
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX + frozenWidth, 0, scrollableWidth, originY);
    ctx.clip();
    for (let visualCol = 0; visualCol < scrollCols.length; visualCol++) {
      const colIndex = scrollCols[visualCol]!;
      ctx.fillText(colToName(colIndex), startXScroll + visualCol * this.cellWidth + this.cellWidth / 2, originY / 2);
    }
    ctx.restore();

    // Row header labels: frozen rows (pinned).
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, originY, originX, frozenHeight);
    ctx.clip();
    for (let visualRow = 0; visualRow < frozenRows.length; visualRow++) {
      const rowIndex = frozenRows[visualRow]!;
      ctx.fillText(String(rowIndex + 1), originX / 2, originY + visualRow * this.cellHeight + this.cellHeight / 2);
    }
    ctx.restore();

    // Row header labels: scrollable rows.
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, originY + frozenHeight, originX, scrollableHeight);
    ctx.clip();
    for (let visualRow = 0; visualRow < scrollRows.length; visualRow++) {
      const rowIndex = scrollRows[visualRow]!;
      ctx.fillText(
        String(rowIndex + 1),
        originX / 2,
        startYScroll + visualRow * this.cellHeight + this.cellHeight / 2
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

  private initSharedChartPanes(): void {
    if (!this.sharedGrid) return;
    if (this.sharedChartPanes) return;

    const createPane = (testId: string) => {
      const pane = document.createElement("div");
      pane.dataset.testid = testId;
      pane.style.position = "absolute";
      pane.style.pointerEvents = "none";
      pane.style.overflow = "hidden";
      pane.style.left = "0";
      pane.style.top = "0";
      pane.style.width = "0";
      pane.style.height = "0";
      return pane;
    };

    const topLeft = createPane("chart-pane-top-left");
    const topRight = createPane("chart-pane-top-right");
    const bottomLeft = createPane("chart-pane-bottom-left");
    const bottomRight = createPane("chart-pane-bottom-right");

    this.chartLayer.appendChild(topLeft);
    this.chartLayer.appendChild(topRight);
    this.chartLayer.appendChild(bottomLeft);
    this.chartLayer.appendChild(bottomRight);

    this.sharedChartPanes = { topLeft, topRight, bottomLeft, bottomRight };

    // Establish an initial pane layout before the first render pass.
    this.syncSharedChartPanes(this.sharedGrid.renderer.scroll.getViewportState());
  }

  private syncSharedChartPanes(viewport: GridViewportState): void {
    if (!this.sharedGrid) return;
    if (!this.sharedChartPanes) return;

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    const headerWidth =
      headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
    const headerHeight =
      headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;

    const headerWidthClamped = Math.min(headerWidth, viewport.width);
    const headerHeightClamped = Math.min(headerHeight, viewport.height);

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

    const frozenContentWidth = Math.max(0, frozenWidthClamped - headerWidthClamped);
    const frozenContentHeight = Math.max(0, frozenHeightClamped - headerHeightClamped);

    const cellAreaWidth = Math.max(0, viewport.width - headerWidthClamped);
    const cellAreaHeight = Math.max(0, viewport.height - headerHeightClamped);

    const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
    const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

    // Charts are rendered as DOM overlays. In shared-grid mode, the column/row
    // headers are implemented as frozen panes, so we position the outer chart
    // layer just under those headers, then clip charts into the four pane
    // quadrants to mimic Excel behavior (objects are constrained to their pane).
    this.chartLayer.style.left = `${headerWidthClamped}px`;
    this.chartLayer.style.top = `${headerHeightClamped}px`;
    this.chartLayer.style.right = "0";
    this.chartLayer.style.bottom = "0";
    this.chartLayer.style.overflow = "hidden";

    const applyPaneRect = (pane: HTMLDivElement, rect: { x: number; y: number; width: number; height: number }) => {
      pane.style.left = `${rect.x}px`;
      pane.style.top = `${rect.y}px`;
      pane.style.width = `${rect.width}px`;
      pane.style.height = `${rect.height}px`;
      pane.style.display = rect.width > 0 && rect.height > 0 ? "block" : "none";
    };

    const { topLeft, topRight, bottomLeft, bottomRight } = this.sharedChartPanes;
    applyPaneRect(topLeft, { x: 0, y: 0, width: frozenContentWidth, height: frozenContentHeight });
    applyPaneRect(topRight, { x: frozenContentWidth, y: 0, width: scrollableWidth, height: frozenContentHeight });
    applyPaneRect(bottomLeft, { x: 0, y: frozenContentHeight, width: frozenContentWidth, height: scrollableHeight });
    applyPaneRect(bottomRight, {
      x: frozenContentWidth,
      y: frozenContentHeight,
      width: scrollableWidth,
      height: scrollableHeight,
    });

    this.sharedChartPaneLayout = { frozenContentWidth, frozenContentHeight };
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
      if (this.sharedGrid) {
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
        const gridRow = anchor.fromRow + headerRows;
        const gridCol = anchor.fromCol + headerCols;
        const { rowCount, colCount } = this.sharedGrid.renderer.scroll.getCounts();
        if (gridRow < 0 || gridRow >= rowCount || gridCol < 0 || gridCol >= colCount) return null;

        left =
          this.sharedGrid.renderer.scroll.cols.positionOf(gridCol) -
          headerWidth +
          emuToPx(anchor.fromColOffEmu ?? 0);
        top =
          this.sharedGrid.renderer.scroll.rows.positionOf(gridRow) -
          headerHeight +
          emuToPx(anchor.fromRowOffEmu ?? 0);
        width = emuToPx(anchor.cxEmu ?? 0);
        height = emuToPx(anchor.cyEmu ?? 0);
      } else {
        left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu);
        top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu);
        width = emuToPx(anchor.cxEmu);
        height = emuToPx(anchor.cyEmu);
      }
    } else if (anchor.kind === "twoCell") {
      if (this.sharedGrid) {
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
        const fromRow = anchor.fromRow + headerRows;
        const fromCol = anchor.fromCol + headerCols;
        const toRow = anchor.toRow + headerRows;
        const toCol = anchor.toCol + headerCols;
        const { rowCount, colCount } = this.sharedGrid.renderer.scroll.getCounts();
        if (fromRow < 0 || fromRow >= rowCount || toRow < 0 || toRow >= rowCount) return null;
        if (fromCol < 0 || fromCol >= colCount || toCol < 0 || toCol >= colCount) return null;

        left =
          this.sharedGrid.renderer.scroll.cols.positionOf(fromCol) -
          headerWidth +
          emuToPx(anchor.fromColOffEmu ?? 0);
        top =
          this.sharedGrid.renderer.scroll.rows.positionOf(fromRow) -
          headerHeight +
          emuToPx(anchor.fromRowOffEmu ?? 0);
        const right =
          this.sharedGrid.renderer.scroll.cols.positionOf(toCol) -
          headerWidth +
          emuToPx(anchor.toColOffEmu ?? 0);
        const bottom =
          this.sharedGrid.renderer.scroll.rows.positionOf(toRow) -
          headerHeight +
          emuToPx(anchor.toRowOffEmu ?? 0);

        width = Math.max(0, right - left);
        height = Math.max(0, bottom - top);
      } else {
        left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu);
        top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu);
        const right = this.visualIndexForCol(anchor.toCol) * this.cellWidth + emuToPx(anchor.toColOffEmu);
        const bottom = this.visualIndexForRow(anchor.toRow) * this.cellHeight + emuToPx(anchor.toRowOffEmu);
        width = Math.max(0, right - left);
        height = Math.max(0, bottom - top);
      }
    } else {
      return null;
    }

    if (width <= 0 || height <= 0) return null;

    const { frozenRows, frozenCols } = this.getFrozen();
    const scrollX = anchor.kind === "absolute" ? this.scrollX : (anchor.fromCol < frozenCols ? 0 : this.scrollX);
    const scrollY = anchor.kind === "absolute" ? this.scrollY : (anchor.fromRow < frozenRows ? 0 : this.scrollY);

    return {
      left: left - scrollX,
      top: top - scrollY,
      width,
      height
    };
  }

  private renderCharts(renderContent: boolean): void {
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    const keep = new Set<string>();

    const sharedPanes = this.sharedGrid ? this.sharedChartPanes : null;
    const sharedLayout = this.sharedGrid ? this.sharedChartPaneLayout : null;

    const resolveHostContainer = (anchor: ChartRecord["anchor"]): { container: HTMLDivElement; originX: number; originY: number } => {
      if (!sharedPanes || !sharedLayout) return { container: this.chartLayer, originX: 0, originY: 0 };

      const { frozenRows, frozenCols } = this.getFrozen();
      const fromRow = anchor.kind === "oneCell" || anchor.kind === "twoCell" ? anchor.fromRow : Number.POSITIVE_INFINITY;
      const fromCol = anchor.kind === "oneCell" || anchor.kind === "twoCell" ? anchor.fromCol : Number.POSITIVE_INFINITY;
      const inFrozenRows = fromRow < frozenRows;
      const inFrozenCols = fromCol < frozenCols;

      const { frozenContentWidth, frozenContentHeight } = sharedLayout;

      if (inFrozenRows && inFrozenCols) return { container: sharedPanes.topLeft, originX: 0, originY: 0 };
      if (inFrozenRows && !inFrozenCols) return { container: sharedPanes.topRight, originX: frozenContentWidth, originY: 0 };
      if (!inFrozenRows && inFrozenCols) return { container: sharedPanes.bottomLeft, originX: 0, originY: frozenContentHeight };
      return { container: sharedPanes.bottomRight, originX: frozenContentWidth, originY: frozenContentHeight };
    };

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
      const { container, originX, originY } = resolveHostContainer(chart.anchor);
      if (!host) {
        host = document.createElement("div");
        host.setAttribute("data-testid", "chart-object");
        host.dataset.chartId = chart.id;
        host.style.position = "absolute";
        host.style.pointerEvents = "none";
        host.style.overflow = "hidden";
        this.chartElements.set(chart.id, host);
        container.appendChild(host);
      } else if (host.parentElement !== container) {
        container.appendChild(host);
      }

      host.style.left = `${rect.left - originX}px`;
      host.style.top = `${rect.top - originY}px`;
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

  private renderAuditing(): void {
    this.auditingRenderer.clear(this.auditingCtx);

    if (this.auditingMode === "off") return;

    const clipRect = this.sharedGrid
      ? (() => {
          const viewport = this.sharedGrid.renderer.scroll.getViewportState();
          const headerWidth = Math.max(0, this.sharedGrid.renderer.getColWidth(0));
          const headerHeight = Math.max(0, this.sharedGrid.renderer.getRowHeight(0));
          return {
            x: headerWidth,
            y: headerHeight,
            width: Math.max(0, viewport.width - headerWidth),
            height: Math.max(0, viewport.height - headerHeight),
          };
        })()
      : {
          x: this.rowHeaderWidth,
          y: this.colHeaderHeight,
          width: this.viewportWidth(),
          height: this.viewportHeight(),
        };

    this.auditingCtx.save();
    this.auditingCtx.beginPath();
    this.auditingCtx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    this.auditingCtx.clip();

    this.auditingRenderer.render(this.auditingCtx, this.auditingHighlights, {
      getCellRect: (row, col) => this.getCellRect({ row, col }),
    });

    this.auditingCtx.restore();
  }

  private renderSelection(): void {
    if (this.sharedGrid) {
      // Selection rendering is handled by the shared CanvasGridRenderer selection layer.
      // We still need to keep the in-place editor aligned when the viewport scrolls or resizes.
      if (this.editor.isOpen()) {
        const gridCell = this.gridCellFromDocCell(this.selection.active);
        const rect = this.sharedGrid.getCellRect(gridCell.row, gridCell.col);
        if (rect) this.editor.reposition(rect);
      }
      return;
    }

    this.ensureViewportMappingCurrent();
    const clipRect = {
      x: this.rowHeaderWidth,
      y: this.colHeaderHeight,
      width: this.viewportWidth(),
      height: this.viewportHeight(),
    };

    const renderer = this.formulaBar?.isFormulaEditing() ? this.formulaSelectionRenderer : this.selectionRenderer;

    const visibleRows = this.frozenVisibleRows.length
      ? [...this.frozenVisibleRows, ...this.visibleRows]
      : this.visibleRows;
    const visibleCols = this.frozenVisibleCols.length
      ? [...this.frozenVisibleCols, ...this.visibleCols]
      : this.visibleCols;

    renderer.render(
      this.selectionCtx,
      this.selection,
      {
        getCellRect: (cell) => this.getCellRect(cell),
        visibleRows,
        visibleCols,
      },
      {
        clipRect,
      }
    );

    if (!this.formulaBar?.isFormulaEditing() && this.fillPreviewRange) {
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
    this.updateSelectionStats();

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

    this.maybeScheduleAuditingUpdate();
  }

  private updateSelectionStats(): void {
    const sumEl = this.status.selectionSum;
    const avgEl = this.status.selectionAverage;
    const countEl = this.status.selectionCount;
    if (!sumEl && !avgEl && !countEl) return;

    const { sum, avg, count } = this.computeSelectionStats();
    const formatter = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });

    if (sumEl) sumEl.textContent = `Sum: ${formatter.format(sum)}`;
    if (avgEl) avgEl.textContent = `Avg: ${formatter.format(avg)}`;
    if (countEl) countEl.textContent = `Count: ${formatter.format(count)}`;
  }

  private computeSelectionStats(): { sum: number; avg: number; count: number } {
    const ranges = this.selection.ranges;
    if (ranges.length === 0) return { sum: 0, avg: 0, count: 0 };

    const normalized = ranges.map((range) => {
      const startRow = Math.min(range.startRow, range.endRow);
      const endRow = Math.max(range.startRow, range.endRow);
      const startCol = Math.min(range.startCol, range.endCol);
      const endCol = Math.max(range.startCol, range.endCol);
      return { startRow, endRow, startCol, endCol };
    });

    let selectedCellCount = 0;
    for (const range of normalized) {
      const rows = Math.max(0, range.endRow - range.startRow + 1);
      const cols = Math.max(0, range.endCol - range.startCol + 1);
      selectedCellCount += rows * cols;
    }

    const useEngineCache = this.document.getSheetIds().length <= 1;
    const memo = new Map<string, SpreadsheetValue>();
    const stack = new Set<string>();

    let sum = 0;
    let count = 0;

    const addValue = (value: SpreadsheetValue): void => {
      const num = coerceNumber(value);
      if (num == null) return;
      sum += num;
      count += 1;
    };

    // For small selections, iterate the selection window directly.
    // For large selections (e.g. select-all), iterate sparse sheet data to avoid O(rows*cols) scans.
    const SPARSE_SELECTION_THRESHOLD = 10_000;
    if (selectedCellCount <= SPARSE_SELECTION_THRESHOLD) {
      for (const range of normalized) {
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let col = range.startCol; col <= range.endCol; col += 1) {
            const coord = { row, col };
            const state = this.document.getCell(this.sheetId, coord) as { value: unknown; formula: string | null };
            if (state?.formula != null) {
              addValue(this.computeCellValue(this.sheetId, coord, memo, stack, { useEngineCache }));
            } else if (isRichTextValue(state?.value)) {
              // Rich text is treated as text (non-numeric) in the status bar quick stats.
            } else {
              addValue((state?.value as SpreadsheetValue) ?? null);
            }
          }
        }
      }
    } else {
      const sheetModel = (this.document as any)?.model?.sheets?.get(this.sheetId) as { cells?: Map<string, any> } | undefined;
      const cells: Map<string, { value: unknown; formula: string | null }> | undefined = sheetModel?.cells;
      if (cells) {
        const inSelection = (row: number, col: number): boolean =>
          normalized.some((range) => row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol);

        for (const [key, cell] of cells.entries()) {
          const [rowStr, colStr] = key.split(",");
          const row = Number(rowStr);
          const col = Number(colStr);
          if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
          if (!inSelection(row, col)) continue;
          const coord = { row, col };
          if (cell.formula != null) {
            addValue(this.computeCellValue(this.sheetId, coord, memo, stack, { useEngineCache }));
          } else if (isRichTextValue(cell.value)) {
            // Ignore rich text (non-numeric).
          } else {
            addValue((cell.value as SpreadsheetValue) ?? null);
          }
        }
      }
    }

    const avg = count === 0 ? 0 : sum / count;
    return { sum, avg, count };
  }

  private syncEngineNow(): void {
    (this.engine as unknown as { syncNow?: () => void }).syncNow?.();
  }

  private onWindowKeyDown(e: KeyboardEvent): void {
    if (e.defaultPrevented) return;
    if (this.handleShowFormulasShortcut(e)) return;
    if (this.handleAuditingShortcut(e)) return;
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
    this.toggleShowFormulas();
    return true;
  }

  private handleAuditingShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return false;
    if (e.code !== "BracketLeft" && e.code !== "BracketRight") return false;

    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    e.preventDefault();
    this.toggleAuditingComponent(e.code === "BracketLeft" ? "precedents" : "dependents");
    return true;
  }

  private toggleAuditingComponent(component: "precedents" | "dependents"): void {
    const hasPrecedents = this.auditingMode === "precedents" || this.auditingMode === "both";
    const hasDependents = this.auditingMode === "dependents" || this.auditingMode === "both";

    const nextPrecedents = component === "precedents" ? !hasPrecedents : hasPrecedents;
    const nextDependents = component === "dependents" ? !hasDependents : hasDependents;

    const nextMode: "off" | AuditingMode =
      nextPrecedents && nextDependents
        ? "both"
        : nextPrecedents
          ? "precedents"
          : nextDependents
            ? "dependents"
            : "off";

    if (nextMode === this.auditingMode) return;
    this.auditingMode = nextMode;
    this.auditingErrors = { precedents: null, dependents: null };

    if (nextMode === "off") {
      this.auditingHighlights = { precedents: new Set(), dependents: new Set() };
      this.updateAuditingLegend();
      this.renderAuditing();
      return;
    }

    this.updateAuditingLegend();
    this.scheduleAuditingUpdate();
    this.renderAuditing();
  }

  private auditingCacheKey(sheetId: string, row: number, col: number, transitive: boolean): string {
    return `${sheetId}:${row},${col}:${transitive ? "t" : "d"}`;
  }

  private getTauriInvoke(): ((cmd: string, args?: Record<string, unknown>) => Promise<unknown>) | null {
    const queued = (globalThis as any).__formulaQueuedInvoke;
    if (typeof queued === "function") {
      return queued;
    }

    const invoke = (globalThis as any).__TAURI__?.core?.invoke;
    if (typeof invoke !== "function") return null;
    return invoke;
  }

  private maybeScheduleAuditingUpdate(): void {
    if (this.auditingMode === "off") return;

    if (this.dragState) {
      this.auditingNeedsUpdateAfterDrag = true;
      return;
    }

    const cell = this.selection.active;
    const key = this.auditingCacheKey(this.sheetId, cell.row, cell.col, this.auditingTransitive);
    if (key === this.auditingLastCellKey) return;

    this.scheduleAuditingUpdate();
  }

  private scheduleAuditingUpdate(): void {
    if (this.auditingMode === "off") return;
    if (this.auditingUpdateScheduled) return;
    this.auditingUpdateScheduled = true;

    const update = new Promise<void>((resolve) => {
      queueMicrotask(() => {
        this.auditingUpdateScheduled = false;
        this.updateAuditingForActiveCell()
          .catch(() => {})
          .finally(() => resolve());
      });
    });
    this.auditingIdlePromise = update;
  }

  private async updateAuditingForActiveCell(): Promise<void> {
    if (this.auditingMode === "off") return;

    const invoke = this.getTauriInvoke();
    const cell = this.selection.active;
    const key = this.auditingCacheKey(this.sheetId, cell.row, cell.col, this.auditingTransitive);
    const cellA1 = cellToA1(cell);

    this.auditingLastCellKey = key;

    const cached = this.auditingCache.get(key);
    if (cached) {
      this.applyAuditingEntry(cached, cellA1);
      return;
    }

    if (!invoke) {
      const entry: AuditingCacheEntry = {
        precedents: [],
        dependents: [],
        precedentsError: "Auditing requires the desktop engine.",
        dependentsError: "Auditing requires the desktop engine.",
      };
      this.auditingCache.set(key, entry);
      this.applyAuditingEntry(entry, cellA1);
      return;
    }

    const requestId = ++this.auditingRequestId;

    const args = {
      sheet_id: this.sheetId,
      row: cell.row,
      col: cell.col,
      transitive: this.auditingTransitive,
    };

    let precedents: string[] = [];
    let dependents: string[] = [];
    let precedentsError: string | null = null;
    let dependentsError: string | null = null;

    try {
      const payload = await invoke("get_precedents", args);
      precedents = Array.isArray(payload) ? payload.map(String) : [];
    } catch (err) {
      precedentsError = String(err);
      precedents = [];
    }

    try {
      const payload = await invoke("get_dependents", args);
      dependents = Array.isArray(payload) ? payload.map(String) : [];
    } catch (err) {
      dependentsError = String(err);
      dependents = [];
    }

    const entry: AuditingCacheEntry = { precedents, dependents, precedentsError, dependentsError };
    this.auditingCache.set(key, entry);

    if (requestId !== this.auditingRequestId) return;
    if (key !== this.auditingLastCellKey) return;

    this.applyAuditingEntry(entry, cellA1);
  }

  private applyAuditingEntry(entry: AuditingCacheEntry, cellA1: string): void {
    const engine = {
      precedents: () => entry.precedents,
      dependents: () => entry.dependents,
    };

    const mode = this.auditingMode === "off" ? "both" : this.auditingMode;
    this.auditingHighlights = computeAuditingOverlays(engine, cellA1, mode, { transitive: this.auditingTransitive });
    this.auditingErrors = { precedents: entry.precedentsError, dependents: entry.dependentsError };

    this.updateAuditingLegend();
    this.renderAuditing();
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
    this.applyUndoRedo(undo ? "undo" : "redo");
    return true;
  }

  private applyUndoRedo(kind: "undo" | "redo"): boolean {
    const did = kind === "undo" ? this.document.undo() : this.document.redo();
    if (!did) return false;

    this.syncEngineNow();
    // Undo/redo can affect sheet view state (e.g. frozen panes). Keep renderer + scrollbars in sync.
    this.syncFrozenPanes();
    if (this.sharedGrid) {
      // Shared grid rendering is driven by CanvasGridRenderer, but we still need to refresh
      // overlays (charts, auditing, etc) after changes.
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
    if (this.sharedGrid) {
      const viewport = this.sharedGrid.renderer.scroll.getViewportState();
      const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);
      return Math.max(0, viewport.width - frozenWidth);
    }
    return Math.max(0, this.width - this.rowHeaderWidth);
  }

  private viewportHeight(): number {
    if (this.sharedGrid) {
      const viewport = this.sharedGrid.renderer.scroll.getViewportState();
      const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);
      return Math.max(0, viewport.height - frozenHeight);
    }
    return Math.max(0, this.height - this.colHeaderHeight);
  }

  private contentWidth(): number {
    return this.colIndexByVisual.length * this.cellWidth;
  }

  private contentHeight(): number {
    return this.rowIndexByVisual.length * this.cellHeight;
  }

  private maxScrollX(): number {
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.scroll.getViewportState().maxScrollX;
    }
    return Math.max(0, this.contentWidth() - this.viewportWidth());
  }

  private maxScrollY(): number {
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.scroll.getViewportState().maxScrollY;
    }
    return Math.max(0, this.contentHeight() - this.viewportHeight());
  }

  private clampScroll(): void {
    if (this.sharedGrid) {
      const scroll = this.sharedGrid.getScroll();
      this.scrollX = scroll.x;
      this.scrollY = scroll.y;
      return;
    }
    const maxX = this.maxScrollX();
    const maxY = this.maxScrollY();
    this.scrollX = Math.min(Math.max(0, this.scrollX), maxX);
    this.scrollY = Math.min(Math.max(0, this.scrollY), maxY);
  }

  private setScroll(nextX: number, nextY: number): boolean {
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      this.sharedGrid.scrollTo(nextX, nextY);
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

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

    const view = this.document.getSheetView(this.sheetId) as { frozenRows?: number; frozenCols?: number } | null;
    const nextFrozenRows = Number.isFinite(view?.frozenRows) ? Math.max(0, Math.trunc(view?.frozenRows ?? 0)) : 0;
    const nextFrozenCols = Number.isFinite(view?.frozenCols) ? Math.max(0, Math.trunc(view?.frozenCols ?? 0)) : 0;

    this.frozenRows = Math.min(nextFrozenRows, this.limits.maxRows);
    this.frozenCols = Math.min(nextFrozenCols, this.limits.maxCols);

    const frozenRowCount = this.lowerBound(this.rowIndexByVisual, this.frozenRows);
    const frozenColCount = this.lowerBound(this.colIndexByVisual, this.frozenCols);
    this.frozenVisibleRows = this.rowIndexByVisual.slice(0, frozenRowCount);
    this.frozenVisibleCols = this.colIndexByVisual.slice(0, frozenColCount);
    this.frozenHeight = frozenRowCount * this.cellHeight;
    this.frozenWidth = frozenColCount * this.cellWidth;

    const overscan = 1;
    const firstRow = Math.max(
      frozenRowCount,
      Math.floor((this.scrollY + this.frozenHeight) / this.cellHeight) - overscan
    );
    const lastRow = Math.min(totalRows, Math.ceil((this.scrollY + availableHeight) / this.cellHeight) + overscan);

    const firstCol = Math.max(
      frozenColCount,
      Math.floor((this.scrollX + this.frozenWidth) / this.cellWidth) - overscan
    );
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
    if (this.sharedGrid) {
      this.sharedGrid.syncScrollbars();
      return;
    }

    // Keep frozen metrics up-to-date so scrollbars reflect the scrollable region.
    this.updateViewportMapping();
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
        viewportSize: Math.max(0, this.viewportHeight() - this.frozenHeight),
        contentSize: Math.max(0, this.contentHeight() - this.frozenHeight),
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
        viewportSize: Math.max(0, this.viewportWidth() - this.frozenWidth),
        contentSize: Math.max(0, this.contentWidth() - this.frozenWidth),
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
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      const gridCell = this.gridCellFromDocCell(cell);
      this.sharedGrid.scrollToCell(gridCell.row, gridCell.col, { align: "auto", padding: paddingPx });
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    // Ensure frozen pane metrics are current even when called outside `renderGrid()`.
    this.updateViewportMapping();

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

    if (cell.col >= this.frozenCols) {
      const minX = nextX + this.frozenWidth + pad;
      const maxX = nextX + viewportWidth - pad;
      if (left < minX) {
        nextX = left - this.frozenWidth - pad;
      } else if (right > maxX) {
        nextX = right - viewportWidth + pad;
      }
    }

    if (cell.row >= this.frozenRows) {
      const minY = nextY + this.frozenHeight + pad;
      const maxY = nextY + viewportHeight - pad;
      if (top < minY) {
        nextY = top - this.frozenHeight - pad;
      } else if (bottom > maxY) {
        nextY = bottom - viewportHeight + pad;
      }
    }

    return this.setScroll(nextX, nextY);
  }

  private scrollRangeIntoView(range: Range, paddingPx = 8): boolean {
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      const start = this.gridCellFromDocCell({ row: range.startRow, col: range.startCol });
      const end = this.gridCellFromDocCell({ row: range.endRow, col: range.endCol });
      const pad = Math.max(0, paddingPx);
      // Best-effort: attempt to bring both the start and end cells into view.
      // If the range is larger than the viewport this will still keep the active cell visible.
      this.sharedGrid.scrollToCell(start.row, start.col, { align: "start", padding: pad });
      this.sharedGrid.scrollToCell(end.row, end.col, { align: "end", padding: pad });
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    // Ensure frozen pane metrics are current even when called outside `renderGrid()`.
    this.updateViewportMapping();

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

    const affectsX = Math.max(startCol, endCol) >= this.frozenCols;
    const affectsY = Math.max(startRow, endRow) >= this.frozenRows;

    const scrollableViewportWidth = Math.max(0, viewportWidth - this.frozenWidth);
    const scrollableViewportHeight = Math.max(0, viewportHeight - this.frozenHeight);
    const scrollableLeft = Math.max(left, this.frozenWidth);
    const scrollableTop = Math.max(top, this.frozenHeight);

    // Only attempt to fully fit the range when it fits within the viewport.
    // Otherwise, fall back to keeping the active cell visible.
    if (affectsX && right - scrollableLeft <= scrollableViewportWidth - pad * 2) {
      const minX = nextX + this.frozenWidth + pad;
      const maxX = nextX + viewportWidth - pad;
      if (scrollableLeft < minX) {
        nextX = scrollableLeft - this.frozenWidth - pad;
      } else if (right > maxX) {
        nextX = right - viewportWidth + pad;
      }
    }

    if (affectsY && bottom - scrollableTop <= scrollableViewportHeight - pad * 2) {
      const minY = nextY + this.frozenHeight + pad;
      const maxY = nextY + viewportHeight - pad;
      if (scrollableTop < minY) {
        nextY = scrollableTop - this.frozenHeight - pad;
      } else if (bottom > maxY) {
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
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
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
    if (this.sharedGrid) {
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const gridRow = cell.row + headerRows;
      const gridCol = cell.col + headerCols;
      return this.sharedGrid.getCellRect(gridRow, gridCol);
    }

    if (cell.row < 0 || cell.row >= this.limits.maxRows) return null;
    if (cell.col < 0 || cell.col >= this.limits.maxCols) return null;

    const rowDirect = this.rowToVisual.get(cell.row);
    const colDirect = this.colToVisual.get(cell.col);

    // Even when the outline hides rows/cols, downstream overlays still need a
    // stable coordinate space. Hidden rows/cols collapse to zero size and share
    // the same origin as the next visible row/col.
    const visualRow = rowDirect ?? this.lowerBound(this.rowIndexByVisual, cell.row);
    const visualCol = colDirect ?? this.lowerBound(this.colIndexByVisual, cell.col);

    const frozenRows = this.frozenRows;
    const frozenCols = this.frozenCols;

    return {
      x: this.rowHeaderWidth + visualCol * this.cellWidth - (cell.col < frozenCols ? 0 : this.scrollX),
      y: this.colHeaderHeight + visualRow * this.cellHeight - (cell.row < frozenRows ? 0 : this.scrollY),
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

    const frozenEdgeX = this.rowHeaderWidth + this.frozenWidth;
    const frozenEdgeY = this.colHeaderHeight + this.frozenHeight;
    const localX = clampedX - this.rowHeaderWidth;
    const localY = clampedY - this.colHeaderHeight;

    const sheetX = clampedX < frozenEdgeX ? localX : this.scrollX + localX;
    const sheetY = clampedY < frozenEdgeY ? localY : this.scrollY + localY;

    const colVisual = Math.floor(sheetX / this.cellWidth);
    const rowVisual = Math.floor(sheetY / this.cellHeight);

    const safeColVisual = Math.max(0, Math.min(this.colIndexByVisual.length - 1, colVisual));
    const safeRowVisual = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, rowVisual));

    const col = this.colIndexByVisual[safeColVisual] ?? 0;
    const row = this.rowIndexByVisual[safeRowVisual] ?? 0;
    return { row, col };
  }

  private computeFillDragTarget(
    source: Range,
    cell: CellCoord
  ): { targetRange: Range; endCell: CellCoord } {
    const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));
    const srcTop = source.startRow;
    const srcBottom = source.endRow;
    const srcLeft = source.startCol;
    const srcRight = source.endCol;

    const rowExtension = cell.row < srcTop ? cell.row - srcTop : cell.row > srcBottom ? cell.row - srcBottom : 0;
    const colExtension = cell.col < srcLeft ? cell.col - srcLeft : cell.col > srcRight ? cell.col - srcRight : 0;

    if (rowExtension === 0 && colExtension === 0) {
      const endCell: CellCoord = {
        row: clamp(cell.row, srcTop, srcBottom),
        col: clamp(cell.col, srcLeft, srcRight)
      };
      return { targetRange: source, endCell };
    }

    const axis =
      rowExtension !== 0 && colExtension !== 0
        ? Math.abs(rowExtension) >= Math.abs(colExtension)
          ? "vertical"
          : "horizontal"
        : rowExtension !== 0
          ? "vertical"
          : "horizontal";

    if (axis === "vertical") {
      const targetRange: Range =
        rowExtension > 0
          ? {
              startRow: source.startRow,
              endRow: Math.max(source.endRow, cell.row),
              startCol: source.startCol,
              endCol: source.endCol
            }
          : {
              startRow: Math.min(source.startRow, cell.row),
              endRow: source.endRow,
              startCol: source.startCol,
              endCol: source.endCol
            };

      const endCell: CellCoord = {
        row: clamp(cell.row, targetRange.startRow, targetRange.endRow),
        col: clamp(cell.col, targetRange.startCol, targetRange.endCol)
      };

      return { targetRange, endCell };
    }

    const targetRange: Range =
      colExtension > 0
        ? {
            startRow: source.startRow,
            endRow: source.endRow,
            startCol: source.startCol,
            endCol: Math.max(source.endCol, cell.col)
          }
        : {
            startRow: source.startRow,
            endRow: source.endRow,
            startCol: Math.min(source.startCol, cell.col),
            endCol: source.endCol
          };

    const endCell: CellCoord = {
      row: clamp(cell.row, targetRange.startRow, targetRange.endRow),
      col: clamp(cell.col, targetRange.startCol, targetRange.endCol)
    };

    return { targetRange, endCell };
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
        const state = this.dragState;
        const source = state.sourceRange;
        const { targetRange, endCell } = this.computeFillDragTarget(source, cell);
        state.targetRange = targetRange;
        state.endCell = endCell;
        this.fillPreviewRange =
          targetRange.startRow === source.startRow &&
          targetRange.endRow === source.endRow &&
          targetRange.startCol === source.startCol &&
          targetRange.endCol === source.endCol
            ? null
            : targetRange;
      } else {
        this.selection = extendSelectionToCell(this.selection, cell, this.limits);

        if (this.dragState.mode === "formula" && this.formulaBar) {
          const r = this.selection.ranges[0];
          if (r) {
            const rangeSheetId =
              this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
            this.formulaBar.updateRangeSelection(
              {
                start: { row: r.startRow, col: r.startCol },
                end: { row: r.endRow, col: r.endCol }
              },
              rangeSheetId
            );
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

    // Right/middle clicks should not start drag selection, but we still want right-click to
    // apply to the cell under the cursor when the click happens outside the current selection.
    // (Excel keeps selection when right-clicking inside it, but moves selection when right-clicking
    // outside.)
    if (e.pointerType === "mouse" && e.button !== 0) {
      if (x >= this.rowHeaderWidth && y >= this.colHeaderHeight) {
        const cell = this.cellFromPoint(x, y);
        const inSelection = this.selection.ranges.some(
          (range) =>
            cell.row >= range.startRow &&
            cell.row <= range.endRow &&
            cell.col >= range.startCol &&
            cell.col <= range.endCol
        );
        if (!inSelection) {
          this.selection = setActiveCell(this.selection, cell, this.limits);
          this.renderSelection();
          this.updateStatus();
        }
      }
      this.focus();
      return;
    }

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
      const frozenEdgeX = this.rowHeaderWidth + this.frozenWidth;
      const localX = x - this.rowHeaderWidth;
      const sheetX = x < frozenEdgeX ? localX : this.scrollX + localX;
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
      const frozenEdgeY = this.colHeaderHeight + this.frozenHeight;
      const localY = y - this.colHeaderHeight;
      const sheetY = y < frozenEdgeY ? localY : this.scrollY + localY;
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
      const rangeSheetId =
        this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
      this.formulaBar.beginRangeSelection(
        {
          start: { row: cell.row, col: cell.col },
          end: { row: cell.row, col: cell.col }
        },
        rangeSheetId
      );
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
        const fillMode: FillHandleMode = e.altKey ? "formulas" : e.ctrlKey || e.metaKey ? "copy" : "series";
        this.dragState = {
          pointerId: e.pointerId,
          mode: "fill",
          sourceRange,
          targetRange: sourceRange,
          endCell: { row: sourceRange.endRow, col: sourceRange.endCol },
          fillMode,
          activeRangeIndex: this.selection.activeRangeIndex
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
        const state = this.dragState;
        const source = state.sourceRange;
        const { targetRange, endCell } = this.computeFillDragTarget(source, cell);
        state.targetRange = targetRange;
        state.endCell = endCell;
        this.fillPreviewRange =
          targetRange.startRow === source.startRow &&
          targetRange.endRow === source.endRow &&
          targetRange.startCol === source.startCol &&
          targetRange.endCol === source.endCol
            ? null
            : targetRange;
        this.renderSelection();
        this.maybeStartDragAutoScroll();
        return;
      }

      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
      this.renderSelection();
      this.updateStatus();

      if (this.dragState.mode === "formula" && this.formulaBar) {
        const r = this.selection.ranges[0];
        const rangeSheetId =
          this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
        this.formulaBar.updateRangeSelection(
          {
            start: { row: r.startRow, col: r.startCol },
            end: { row: r.endRow, col: r.endCol }
          },
          rangeSheetId
        );
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
    this.commentTooltip.classList.add("comment-tooltip--visible");
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
      const shouldCommit = e.type === "pointerup";
      const { sourceRange, targetRange, endCell, fillMode, activeRangeIndex } = state;
      const changed =
        targetRange.startRow !== sourceRange.startRow ||
        targetRange.endRow !== sourceRange.endRow ||
        targetRange.startCol !== sourceRange.startCol ||
        targetRange.endCol !== sourceRange.endCol;

      if (changed && shouldCommit) {
        this.applyFill(sourceRange, targetRange, fillMode);

        const existingRanges = this.selection.ranges;
        const nextActiveIndex = Math.max(0, Math.min(existingRanges.length - 1, activeRangeIndex));
        const updatedRanges = existingRanges.length === 0 ? [targetRange] : [...existingRanges];
        updatedRanges[nextActiveIndex] = targetRange;

        this.selection = buildSelection(
          {
            ranges: updatedRanges,
            active: endCell,
            anchor: endCell,
            activeRangeIndex: nextActiveIndex
          },
          this.limits
        );

        this.refresh();
        this.focus();
      } else {
        // Clear any preview overlay.
        this.renderSelection();
      }

      if (this.auditingNeedsUpdateAfterDrag) {
        this.auditingNeedsUpdateAfterDrag = false;
        this.scheduleAuditingUpdate();
      }
      return;
    }

    if (this.auditingNeedsUpdateAfterDrag) {
      this.auditingNeedsUpdateAfterDrag = false;
      this.scheduleAuditingUpdate();
    }

    if (state.mode === "formula" && this.formulaBar) {
      this.formulaBar.endRangeSelection();
      // Restore focus to the formula bar without clearing its insertion state mid-drag.
      this.formulaBar.focus();
    }
  }

  private applyFill(sourceRange: Range, targetRange: Range, mode: FillHandleMode): void {
    const toFillRange = (range: Range): FillEngineRange => ({
      startRow: range.startRow,
      endRow: range.endRow + 1,
      startCol: range.startCol,
      endCol: range.endCol + 1
    });

    const source = toFillRange(sourceRange);
    const union = toFillRange(targetRange);

    const deltaRange: FillEngineRange | null = (() => {
      const sameCols = source.startCol === union.startCol && source.endCol === union.endCol;
      const sameRows = source.startRow === union.startRow && source.endRow === union.endRow;

      if (sameCols) {
        if (union.endRow > source.endRow) {
          return { startRow: source.endRow, endRow: union.endRow, startCol: source.startCol, endCol: source.endCol };
        }
        if (union.startRow < source.startRow) {
          return { startRow: union.startRow, endRow: source.startRow, startCol: source.startCol, endCol: source.endCol };
        }
      }

      if (sameRows) {
        if (union.endCol > source.endCol) {
          return { startRow: source.startRow, endRow: source.endRow, startCol: source.endCol, endCol: union.endCol };
        }
        if (union.startCol < source.startCol) {
          return { startRow: source.startRow, endRow: source.endRow, startCol: union.startCol, endCol: source.startCol };
        }
      }

      return null;
    })();

    if (!deltaRange) return;

    applyFillCommitToDocumentController({
      document: this.document,
      sheetId: this.sheetId,
      sourceRange: source,
      targetRange: deltaRange,
      mode,
      getCellComputedValue: (row, col) => this.getCellComputedValue({ row, col }) as any
    });
  }

  private applyFillShortcut(direction: "down" | "right"): void {
    const clampInt = (value: number, min: number, max: number): number => {
      const n = Math.trunc(value);
      return Math.min(max, Math.max(min, n));
    };

    const maxRowInclusive = Math.max(0, this.limits.maxRows - 1);
    const maxColInclusive = Math.max(0, this.limits.maxCols - 1);
    const maxRowExclusive = Math.max(0, this.limits.maxRows);
    const maxColExclusive = Math.max(0, this.limits.maxCols);

    const operations: Array<{ sourceRange: FillEngineRange; targetRange: FillEngineRange }> = [];

    for (const range of this.selection.ranges) {
      const startRow = clampInt(Math.min(range.startRow, range.endRow), 0, maxRowInclusive);
      const endRow = clampInt(Math.max(range.startRow, range.endRow), 0, maxRowInclusive);
      const startCol = clampInt(Math.min(range.startCol, range.endCol), 0, maxColInclusive);
      const endCol = clampInt(Math.max(range.startCol, range.endCol), 0, maxColInclusive);

      if (direction === "down") {
        if (endRow <= startRow) continue; // height === 1

        const sourceRange: FillEngineRange = {
          startRow,
          endRow: Math.min(startRow + 1, maxRowExclusive),
          startCol,
          endCol: Math.min(endCol + 1, maxColExclusive)
        };
        const targetRange: FillEngineRange = {
          startRow: Math.min(startRow + 1, maxRowExclusive),
          endRow: Math.min(endRow + 1, maxRowExclusive),
          startCol,
          endCol: Math.min(endCol + 1, maxColExclusive)
        };

        if (targetRange.endRow <= targetRange.startRow) continue;
        if (sourceRange.endCol <= sourceRange.startCol) continue;
        operations.push({ sourceRange, targetRange });
        continue;
      }

      if (endCol <= startCol) continue; // width === 1

      const sourceRange: FillEngineRange = {
        startRow,
        endRow: Math.min(endRow + 1, maxRowExclusive),
        startCol,
        endCol: Math.min(startCol + 1, maxColExclusive)
      };
      const targetRange: FillEngineRange = {
        startRow,
        endRow: Math.min(endRow + 1, maxRowExclusive),
        startCol: Math.min(startCol + 1, maxColExclusive),
        endCol: Math.min(endCol + 1, maxColExclusive)
      };

      if (targetRange.endCol <= targetRange.startCol) continue;
      if (sourceRange.endRow <= sourceRange.startRow) continue;
      operations.push({ sourceRange, targetRange });
    }

    if (operations.length === 0) return;

    // Explicit batch so multi-range selections become a single undo step.
    this.document.beginBatch({ label: "Fill" });
    try {
      for (const op of operations) {
        applyFillCommitToDocumentController({
          document: this.document,
          sheetId: this.sheetId,
          sourceRange: op.sourceRange,
          targetRange: op.targetRange,
          mode: "formulas",
          getCellComputedValue: (row, col) => this.getCellComputedValue({ row, col }) as any
        });
      }
    } finally {
      this.document.endBatch();
    }

    this.refresh();
    this.focus();
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (this.inlineEditController.isOpen()) {
      return;
    }
    if (this.editor.isOpen()) {
      // The editor handles Enter/Tab/Escape itself. We keep focus on the textarea.
      return;
    }

    if (e.key === "Escape") {
      if (this.dragState?.mode === "fill") {
        e.preventDefault();
        const state = this.dragState;
        this.dragState = null;
        this.dragPointerPos = null;
        this.fillPreviewRange = null;
        if (this.dragAutoScrollRaf != null) {
          if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
          else globalThis.clearTimeout(this.dragAutoScrollRaf);
        }
        this.dragAutoScrollRaf = null;

        try {
          this.root.releasePointerCapture(state.pointerId);
        } catch {
          // ignore
        }

        this.renderSelection();

        if (this.auditingNeedsUpdateAfterDrag) {
          this.auditingNeedsUpdateAfterDrag = false;
          this.scheduleAuditingUpdate();
        }
        return;
      }

      if (this.sharedGrid?.cancelFillHandleDrag()) {
        e.preventDefault();
        return;
      }
    }

    if (this.handleUndoRedoShortcut(e)) return;
    if (this.handleShowFormulasShortcut(e)) return;
    if (this.handleAuditingShortcut(e)) return;
    if (this.handleClipboardShortcut(e)) return;
    if (this.handleAutoSumShortcut(e)) return;

    // Editing
    if (e.key === "F2") {
      e.preventDefault();
      const cell = this.selection.active;
      const bounds = this.getCellRect(cell);
      if (!bounds) return;
      const initialValue = this.getCellInputText(cell);
      this.editor.open(cell, bounds, initialValue, { cursor: "end" });
      this.updateEditState();
      return;
    }

    const primary = e.ctrlKey || e.metaKey;

    // Excel-style fill shortcuts:
    // - Ctrl/Cmd+D: Fill Down
    // - Ctrl/Cmd+R: Fill Right
    // Only trigger when *not* editing text (editor, formula bar, inline edit).
    if (primary && !e.altKey && !e.shiftKey && (e.key === "d" || e.key === "D")) {
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.applyFillShortcut("down");
      return;
    }

    if (primary && !e.altKey && !e.shiftKey && (e.key === "r" || e.key === "R")) {
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.applyFillShortcut("right");
      return;
    }

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
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (primary && e.code === "Space") {
      // Ctrl+Space selects entire column.
      e.preventDefault();
      this.selection = selectColumns(this.selection, this.selection.active.col, this.selection.active.col, {}, this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (!primary && e.shiftKey && e.code === "Space") {
      // Shift+Space selects entire row.
      e.preventDefault();
      this.selection = selectRows(this.selection, this.selection.active.row, this.selection.active.row, {}, this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
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
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      else if (didScroll) this.ensureViewportMappingCurrent();
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
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      else if (didScroll) this.ensureViewportMappingCurrent();
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
      this.updateEditState();
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
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
    else if (didScroll) this.ensureViewportMappingCurrent();
    this.renderSelection();
    this.updateStatus();
    if (didScroll) this.refresh("scroll");
  }

  private handleAutoSumShortcut(e: KeyboardEvent): boolean {
    if (!e.altKey) return false;
    if (e.code !== "Equal") return false;
    // Avoid hijacking Ctrl/Cmd-modified shortcuts.
    if (e.ctrlKey || e.metaKey) return false;

    // Only trigger when not actively editing.
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;

    e.preventDefault();
    this.autoSumSelection();
    return true;
  }

  private autoSumSelection(): void {
    const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    if (!range) return;

    const target = this.autoSumTargetCell(range);
    if (!target) return;

    const formula = `=SUM(${rangeToA1(range)})`;
    this.document.setCellInput(this.sheetId, target, formula, { label: "AutoSum" });
    this.activateCell({ row: target.row, col: target.col });
    this.refresh();
  }

  private autoSumTargetCell(range: Range): CellCoord | null {
    const inBounds = (cell: CellCoord): boolean =>
      cell.row >= 0 && cell.col >= 0 && cell.row < this.limits.maxRows && cell.col < this.limits.maxCols;

    const isSingleCol = range.startCol === range.endCol;
    const isSingleRow = range.startRow === range.endRow;

    if (isSingleCol) {
      const below = { row: range.endRow + 1, col: range.startCol };
      return inBounds(below) ? below : null;
    }

    if (isSingleRow) {
      const right = { row: range.startRow, col: range.endCol + 1 };
      return inBounds(right) ? right : null;
    }

    const diag = { row: range.endRow + 1, col: range.endCol + 1 };
    if (inBounds(diag)) return diag;

    // Fall back to a cell directly below the selected range.
    const below = { row: range.endRow + 1, col: range.endCol };
    return inBounds(below) ? below : null;
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
      const cells = this.snapshotClipboardCells(range);
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
            row.map((cell: any) => {
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

      const cells = this.snapshotClipboardCells(range);
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
    const address = cellToA1(cell);
    const cacheKey = this.computedKey(this.sheetId, address);

    // The WASM engine currently cannot resolve sheet-qualified references (e.g. `Sheet2!A1`),
    // so when multiple sheets exist we fall back to the in-process evaluator for *all* formulas
    // to keep dependent values consistent.
    const useEngineCache = this.document.getSheetIds().length <= 1;
    if (useEngineCache && this.computedValues.has(cacheKey)) {
      return this.computedValues.get(cacheKey) ?? null;
    }

    const memo = new Map<string, SpreadsheetValue>();
    const stack = new Set<string>();
    return this.computeCellValue(this.sheetId, cell, memo, stack, { useEngineCache });
  }

  private computedKey(sheetId: string, address: string): string {
    return `${sheetId}:${address.replaceAll("$", "").toUpperCase()}`;
  }

  private invalidateComputedValues(changes: unknown): void {
    if (!Array.isArray(changes)) return;
    for (const change of changes) {
      const ref = change as EngineCellRef;
      const sheetId = typeof ref.sheetId === "string" ? ref.sheetId : this.sheetId;
      if (!isInteger(ref.row) || !isInteger(ref.col)) continue;
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
      if (!address && isInteger(ref.row) && isInteger(ref.col)) {
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
      // Some engine implementations omit `value` entirely to represent an empty cell.
      // Treat missing values as null so we don't keep stale computed results around.
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
    stack: Set<string>,
    options: { useEngineCache: boolean }
  ): SpreadsheetValue {
    const address = cellToA1(cell);
    const computedKey = this.computedKey(sheetId, address);
    if (options.useEngineCache && this.computedValues.has(computedKey)) {
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
         return this.computeCellValue(targetSheet, coord, memo, stack, options);
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

  private applyEdit(sheetId: string, cell: CellCoord, rawValue: string): void {
    const original = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };
    if (rawValue.trim() === "") {
      this.document.clearCell(sheetId, cell, { label: "Clear cell" });
      return;
    }

    // Preserve rich-text formatting runs when editing a rich-text cell with plain text
    // (but still allow formulas / leading apostrophes to override rich-text semantics).
    const trimmedStart = rawValue.trimStart();
    if (!trimmedStart.startsWith("=") && !rawValue.startsWith("'") && isRichTextValue(original?.value)) {
      const updated = applyPlainTextEdit(original.value, rawValue);
      if (original.formula == null && updated === original.value) {
        // No-op edit: keep rich runs without creating a history entry.
        return;
      }
      this.document.setCellValue(sheetId, cell, updated, { label: "Edit cell" });
      return;
    }
    this.document.setCellInput(sheetId, cell, rawValue, { label: "Edit cell" });
  }

  private commitFormulaBar(text: string): void {
    this.updateEditState();
    const target = this.formulaEditCell ?? { sheetId: this.sheetId, cell: { ...this.selection.active } };
    this.applyEdit(target.sheetId, target.cell, text);

    this.formulaEditCell = null;
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.referenceHighlightsSource = [];

    if (this.sharedGrid) {
      this.syncSharedGridInteractionMode();
      this.sharedGrid.clearRangeSelection();
    }

    // Restore focus + selection to the original edit cell, even if the user
    // navigated to another sheet while picking ranges.
    this.activateCell({ sheetId: target.sheetId, row: target.cell.row, col: target.cell.col });
    this.refresh();
    this.focus();
  }

  private cancelFormulaBar(): void {
    this.updateEditState();
    const target = this.formulaEditCell;
    this.formulaEditCell = null;
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.referenceHighlightsSource = [];

    if (this.sharedGrid) {
      this.syncSharedGridInteractionMode();
      this.sharedGrid.clearRangeSelection();
    }

    if (target) {
      // Restore the original edit location (sheet + cell).
      this.activateCell({ sheetId: target.sheetId, row: target.cell.row, col: target.cell.col });
      this.renderReferencePreview();
      return;
    }

    this.ensureActiveCellVisible();
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
    else if (didScroll) this.ensureViewportMappingCurrent();
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private computeReferenceHighlightsForSheet(
    sheetId: string,
    highlights: typeof this.referenceHighlightsSource
  ): Array<{ start: CellCoord; end: CellCoord; color: string; active: boolean }> {
    if (!highlights || highlights.length === 0) return [];

    const sheetIds = this.document.getSheetIds();
    const resolveSheetId = (name: string): string | null => {
      const trimmed = name.trim();
      if (!trimmed) return null;
      return sheetIds.find((id) => id.toLowerCase() === trimmed.toLowerCase()) ?? null;
    };

    return highlights
      .filter((h) => {
        const sheet = h.range.sheet;
        if (!sheet) return true;
        const resolved = resolveSheetId(sheet);
        if (!resolved) return false;
        return resolved.toLowerCase() === sheetId.toLowerCase();
      })
      .map((h) => ({
        start: { row: h.range.startRow, col: h.range.startCol },
        end: { row: h.range.endRow, col: h.range.endCol },
        color: h.color,
        active: Boolean(h.active),
      }));
  }

  private renderReferencePreview(): void {
    if (this.sharedGrid) {
      // Reference previews are rendered via the shared grid's `rangeSelection` overlay.
      // The legacy implementation paints onto the content canvas, which would clobber cell text.
      return;
    }

    const ctx = this.referenceCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.referenceCanvas.width, this.referenceCanvas.height);
    ctx.restore();

    if (this.referenceHighlights.length === 0 && !this.referencePreview) return;
    this.ensureViewportMappingCurrent();

    const drawRangeOutline = (
      startRow: number,
      endRow: number,
      startCol: number,
      endCol: number,
      options: { color: string; lineWidth?: number; dash?: number[] }
    ) => {
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
      ctx.strokeStyle = options.color;
      ctx.lineWidth = options.lineWidth ?? 2;
      ctx.setLineDash(options.dash ?? []);
      const inset = ctx.lineWidth / 2;
      if (width > ctx.lineWidth && height > ctx.lineWidth) {
        ctx.strokeRect(x + inset, y + inset, width - ctx.lineWidth, height - ctx.lineWidth);
      }
      ctx.restore();
    };

    // In formula-editing mode, the formula bar provides `referenceHighlights` with
    // per-reference colors and an `active` flag. Render all non-active refs as
    // dashed outlines, and the active one (if any) as a thicker solid outline.
    //
    // When not editing, only the single `referencePreview` (hover) is shown.
    for (const highlight of this.referenceHighlights) {
      if (highlight.active) continue;
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, { color: highlight.color, dash: [4, 3] });
    }

    for (const highlight of this.referenceHighlights) {
      if (!highlight.active) continue;
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, { color: highlight.color, lineWidth: 3 });
    }

    if (this.referenceHighlights.length === 0 && this.referencePreview) {
      const startRow = Math.min(this.referencePreview.start.row, this.referencePreview.end.row);
      const endRow = Math.max(this.referencePreview.start.row, this.referencePreview.end.row);
      const startCol = Math.min(this.referencePreview.start.col, this.referencePreview.end.col);
      const endCol = Math.max(this.referencePreview.start.col, this.referencePreview.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, {
        color: resolveCssVar("--warning", { fallback: "CanvasText" }),
        dash: [4, 3],
      });
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
