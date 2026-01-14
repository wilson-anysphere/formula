import type {
  CellProvider,
  CellProviderUpdate,
  CellRange,
  FillCommitEvent,
  FillDragPreview,
  FillMode,
  GridAxisSizeChange,
  CanvasGridImageResolver,
  GridPerfStats,
  GridViewportState,
  ScrollToCellAlign
} from "@formula/grid";
import {
  applySrOnlyStyle,
  alignScrollToDevicePixels,
  CanvasGridRenderer,
  computeFillPreview,
  computeScrollbarThumb,
  DEFAULT_GRID_FONT_FAMILY,
  DEFAULT_GRID_MONOSPACE_FONT_FAMILY,
  describeActiveCellLabel,
  describeCellForA11y,
  hitTestSelectionHandle,
  resolveGridThemeFromCssVars,
  wheelDeltaToPixels,
} from "@formula/grid";

import { openExternalHyperlink } from "../../hyperlinks/openExternal.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import { shellOpen } from "../../tauri/shellOpen.js";
import { looksLikeExternalHyperlink } from "./looksLikeExternalHyperlink.js";
import { resolveCssVar } from "../../theme/cssVars.js";

export type DesktopGridInteractionMode = "default" | "rangeSelection";

export interface DesktopSharedGridCallbacks {
  onScroll?: (scroll: { x: number; y: number }, viewport: GridViewportState) => void;
  onSelectionChange?: (selection: { row: number; col: number } | null) => void;
  onSelectionRangeChange?: (range: CellRange | null) => void;
  onRequestCellEdit?: (request: { row: number; col: number; initialKey?: string }) => void;
  onAxisSizeChange?: (change: GridAxisSizeChange) => void;

  onRangeSelectionStart?: (range: CellRange) => void;
  onRangeSelectionChange?: (range: CellRange) => void;
  onRangeSelectionEnd?: (range: CellRange) => void;

  /**
   * Called when the user commits a fill handle drag.
   *
   * The `targetRange` excludes the `sourceRange` (it only includes the cells that
   * should be written).
   */
  onFillCommit?: (event: FillCommitEvent) => void | Promise<void>;
}

type ResizeHit = { kind: "col"; index: number } | { kind: "row"; index: number };

type ResizeDragState =
  | { kind: "col"; index: number; startClient: number; startSize: number }
  | { kind: "row"; index: number; startClient: number; startSize: number };

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return clamp(Math.trunc(value), min, max);
}

function rangesEqual(a: CellRange | null, b: CellRange | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.startRow === b.startRow && a.endRow === b.endRow && a.startCol === b.startCol && a.endCol === b.endCol;
}

export class DesktopSharedGrid {
  readonly renderer: CanvasGridRenderer;

  private readonly container: HTMLElement;
  private readonly provider: CellProvider;
  private readonly callbacks: DesktopSharedGridCallbacks;

  private readonly gridCanvas: HTMLCanvasElement;
  private readonly contentCanvas: HTMLCanvasElement;
  private readonly selectionCanvas: HTMLCanvasElement;

  private readonly vTrack: HTMLDivElement;
  private readonly vThumb: HTMLDivElement;
  private readonly hTrack: HTMLDivElement;
  private readonly hThumb: HTMLDivElement;

  private lastScrollbarLayout:
    | { showV: boolean; showH: boolean; frozenWidth: number; frozenHeight: number }
    | null = null;

  private lastScrollbarThumb: { vSize: number | null; vOffset: number | null; hSize: number | null; hOffset: number | null } = {
    vSize: null,
    vOffset: null,
    hSize: null,
    hOffset: null
  };
  private readonly scrollbarThumbScratch = {
    v: { size: 0, offset: 0 },
    h: { size: 0, offset: 0 }
  };

  private readonly frozenRows: number;
  private readonly frozenCols: number;
  private readonly headerRows: number;
  private readonly headerCols: number;

  private interactionMode: DesktopGridInteractionMode = "default";

  private selectionAnchor: { row: number; col: number } | null = null;
  private keyboardAnchor: { row: number; col: number } | null = null;
  private selectionPointerId: number | null = null;
  /**
   * Cached selection canvas origin (client-space), used to avoid layout reads in
   * hot pointer-move paths.
   */
  private selectionCanvasViewportOrigin: { left: number; top: number } | null = null;
  private transientRange: CellRange | null = null;
  private lastPointerViewport: { x: number; y: number } | null = null;
  private readonly dragViewportPointScratch = { x: 0, y: 0 };
  private readonly hoverViewportPointScratch = { x: 0, y: 0 };
  private readonly pickCellScratch = { row: 0, col: 0 };
  private readonly fillHandlePointerCellScratch = { row: 0, col: 0 };
  private readonly selectionDragRangeScratch: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
  private readonly fillDragPreviewScratch: FillDragPreview = {
    axis: "vertical",
    unionRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 }
  };
  private lastDragPickedRow: number | null = null;
  private lastDragPickedCol: number | null = null;
  private lastFillHandlePointerRow: number | null = null;
  private lastFillHandlePointerCol: number | null = null;
  private autoScrollFrame: number | null = null;

  private dragMode: "selection" | "fillHandle" | null = null;
  private fillHandleState: {
    source: CellRange;
    target: CellRange;
    mode: FillMode;
    previewTarget: CellRange;
    endCell: { row: number; col: number };
  } | null = null;

  private resizePointerId: number | null = null;
  private resizeDrag: ResizeDragState | null = null;

  private lastAnnounced: {
    selection: { row: number; col: number } | null;
    range: CellRange | null;
    statusText: string;
    activeCellLabel: string;
  } = {
    selection: null,
    range: null,
    statusText: "",
    activeCellLabel: ""
  };

  private lastEmittedViewport: GridViewportState | null = null;

  private devicePixelRatio = 1;

  private readonly a11yStatusId: string;
  private readonly a11yStatusEl: HTMLDivElement;
  private readonly a11yActiveCellId: string;
  private readonly a11yActiveCellEl: HTMLDivElement;
  private readonly containerRestore: {
    role: string | null;
    ariaRowcount: string | null;
    ariaColcount: string | null;
    ariaMultiselectable: string | null;
    ariaDescribedBy: string | null;
    touchAction: string;
  };

  private disposeFns: Array<() => void> = [];

  constructor(options: {
    container: HTMLElement;
    provider: CellProvider;
    rowCount: number;
    colCount: number;
    canvases: { grid: HTMLCanvasElement; content: HTMLCanvasElement; selection: HTMLCanvasElement };
    scrollbars: { vTrack: HTMLDivElement; vThumb: HTMLDivElement; hTrack: HTMLDivElement; hThumb: HTMLDivElement };
    frozenRows?: number;
    frozenCols?: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
    prefetchOverscanRows?: number;
    prefetchOverscanCols?: number;
    imageResolver?: CanvasGridImageResolver | null;
    enableResize?: boolean;
    enableKeyboard?: boolean;
    enableWheel?: boolean;
    callbacks?: DesktopSharedGridCallbacks;
  }) {
    this.container = options.container;
    this.provider = options.provider;
    this.callbacks = options.callbacks ?? {};

    this.gridCanvas = options.canvases.grid;
    this.contentCanvas = options.canvases.content;
    this.selectionCanvas = options.canvases.selection;

    this.vTrack = options.scrollbars.vTrack;
    this.vThumb = options.scrollbars.vThumb;
    this.hTrack = options.scrollbars.hTrack;
    this.hThumb = options.scrollbars.hThumb;

    this.frozenRows = options.frozenRows ?? 0;
    this.frozenCols = options.frozenCols ?? 0;
    this.headerRows = this.frozenRows > 0 ? 1 : 0;
    this.headerCols = this.frozenCols > 0 ? 1 : 0;

    this.renderer = new CanvasGridRenderer({
      provider: options.provider,
      rowCount: options.rowCount,
      colCount: options.colCount,
      defaultRowHeight: options.defaultRowHeight,
      defaultColWidth: options.defaultColWidth,
      prefetchOverscanRows: options.prefetchOverscanRows,
      prefetchOverscanCols: options.prefetchOverscanCols,
      // Desktop UX: cell content is monospace by default, while headers and other chrome
      // remain in system UI fonts.
      defaultCellFontFamily: resolveCssVar("--font-mono", { fallback: DEFAULT_GRID_MONOSPACE_FONT_FAMILY }),
      defaultHeaderFontFamily: resolveCssVar("--font-sans", { fallback: DEFAULT_GRID_FONT_FAMILY }),
      imageResolver: options.imageResolver ?? null
    });

    // Match the React CanvasGrid behaviour: enable resize unless explicitly disabled.
    const enableResize = options.enableResize ?? true;

    this.a11yStatusId = `desktop-grid-a11y-${Math.random().toString(36).slice(2)}`;
    this.a11yStatusEl = document.createElement("div");
    this.a11yStatusEl.id = this.a11yStatusId;
    this.a11yStatusEl.dataset.testid = "canvas-grid-a11y-status";
    this.a11yStatusEl.setAttribute("role", "status");
    this.a11yStatusEl.setAttribute("aria-live", "polite");
    this.a11yStatusEl.setAttribute("aria-atomic", "true");
    applySrOnlyStyle(this.a11yStatusEl);
    this.a11yStatusEl.textContent = describeCellForA11y({
      selection: null,
      range: null,
      provider: this.provider,
      headerRows: this.headerRows,
      headerCols: this.headerCols,
    });
    this.container.appendChild(this.a11yStatusEl);

    this.a11yActiveCellId = `desktop-grid-active-cell-${this.a11yStatusId}`;
    this.a11yActiveCellEl = document.createElement("div");
    this.a11yActiveCellEl.id = this.a11yActiveCellId;
    this.a11yActiveCellEl.dataset.testid = "canvas-grid-a11y-active-cell";
    this.a11yActiveCellEl.setAttribute("role", "gridcell");
    // Keep the element mounted for the lifetime of the grid so aria-activedescendant can
    // reliably reference it when the active cell changes.
    this.a11yActiveCellEl.setAttribute("aria-hidden", "true");
    applySrOnlyStyle(this.a11yActiveCellEl);
    this.container.appendChild(this.a11yActiveCellEl);

    this.containerRestore = {
      role: this.container.getAttribute("role"),
      ariaRowcount: this.container.getAttribute("aria-rowcount"),
      ariaColcount: this.container.getAttribute("aria-colcount"),
      ariaMultiselectable: this.container.getAttribute("aria-multiselectable"),
      ariaDescribedBy: this.container.getAttribute("aria-describedby"),
      touchAction: this.container.style.touchAction,
    };

    this.container.setAttribute("role", "grid");
    this.container.setAttribute("aria-rowcount", String(options.rowCount));
    this.container.setAttribute("aria-colcount", String(options.colCount));
    this.container.setAttribute("aria-multiselectable", "true");
    // If the container is re-used (e.g. grid mode switches), ensure we start from a clean state.
    this.container.removeAttribute("aria-activedescendant");
    const describedBy = new Set((this.containerRestore.ariaDescribedBy ?? "").split(/\s+/).filter(Boolean));
    describedBy.add(this.a11yStatusId);
    this.container.setAttribute("aria-describedby", Array.from(describedBy).join(" "));
    this.container.style.touchAction = "none";

    this.renderer.attach({
      grid: this.gridCanvas,
      content: this.contentCanvas,
      selection: this.selectionCanvas
    });

    // `CanvasGridRenderer.attach()` assigns inline z-index values (0/1/2) so the grid can be
    // embedded in isolation without depending on any external CSS.
    //
    // The desktop app relies on a single z-index system shared across canvases + DOM overlays
    // (drawings, charts, selection, outline, scrollbars). Clear the inline z-index so
    // `apps/desktop/src/styles/charts-overlay.css` can deterministically control stacking.
    this.gridCanvas.style.zIndex = "";
    this.contentCanvas.style.zIndex = "";
    this.selectionCanvas.style.zIndex = "";

    this.renderer.setFrozen(this.frozenRows, this.frozenCols);
    this.renderer.setFillHandleEnabled(this.interactionMode === "default" && Boolean(this.callbacks.onFillCommit));

    if (this.provider.subscribe) {
      this.disposeFns.push(
        this.provider.subscribe((update: CellProviderUpdate) => {
          const selection = this.renderer.getSelection();
          if (!selection) return;

          if (update.type === "invalidateAll") {
            this.announceSelection(selection, this.renderer.getSelectionRange());
            return;
          }

          const range = update.range;
          if (
            selection.row >= range.startRow &&
            selection.row < range.endRow &&
            selection.col >= range.startCol &&
            selection.col < range.endCol
          ) {
            this.announceSelection(selection, this.renderer.getSelectionRange());
          }
        })
      );
    }

    // Keep desktop shared scrollbars in sync with renderer viewport changes (axis size overrides,
    // frozen pane updates, zoom, resize, etc) even when scroll offsets do not change.
    //
    // Batch notifications to the next animation frame to avoid redundant work during resize drags.
    this.disposeFns.push(
      this.renderer.subscribeViewport(
        () => {
          this.syncScrollbars();
          // Only emit scroll notifications once the grid has a real viewport size. This avoids
          // dispatching an initial `{ width: 0, height: 0 }` viewport from the constructor-time
          // `setFrozen()` call before the host has a chance to run the first `resize()`.
          const viewport = this.renderer.getViewportState();
          if (viewport.width > 0 || viewport.height > 0) {
            this.emitScroll();
          }
        },
        { animationFrame: true }
      )
    );

    // Attempt to resolve theme CSS vars if the host app defines them, and keep the
    // renderer in sync with future theme/media changes.
    this.refreshTheme();
    this.installThemeWatchers();

    // Ensure the selection canvas captures pointer events (background/content are paint-only).
    this.gridCanvas.style.pointerEvents = "none";
    this.contentCanvas.style.pointerEvents = "none";
    this.selectionCanvas.style.pointerEvents = "auto";

    const enableWheel = options.enableWheel ?? true;
    const enableKeyboard = options.enableKeyboard ?? true;

    if (enableWheel) this.installWheelHandler();
    if (enableKeyboard) this.installKeyboardHandler();
    this.installPointerHandlers({ enableResize });
    this.installScrollbarHandlers();
    this.syncScrollbars();
  }

  destroy(): void {
    for (const dispose of this.disposeFns) dispose();
    this.disposeFns = [];
    this.stopAutoScroll();
    this.renderer.destroy();
    this.container.removeAttribute("aria-activedescendant");
    const restoreAttribute = (name: string, value: string | null) => {
      if (value == null) this.container.removeAttribute(name);
      else this.container.setAttribute(name, value);
    };
    restoreAttribute("aria-describedby", this.containerRestore.ariaDescribedBy);
    restoreAttribute("aria-multiselectable", this.containerRestore.ariaMultiselectable);
    restoreAttribute("aria-rowcount", this.containerRestore.ariaRowcount);
    restoreAttribute("aria-colcount", this.containerRestore.ariaColcount);
    restoreAttribute("role", this.containerRestore.role);
    this.container.style.touchAction = this.containerRestore.touchAction;
    this.a11yStatusEl.remove();
    this.a11yActiveCellEl.remove();

    // Release backing store allocations for the canvas layers in case the
    // DesktopSharedGrid instance (or its canvas elements) are kept referenced
    // after teardown (tests/hot reload, split view toggling, etc).
    const resetCanvas = (canvas: HTMLCanvasElement) => {
      try {
        canvas.width = 0;
        canvas.height = 0;
      } catch {
        // Best-effort: ignore failures for mocked canvases.
      }
    };
    resetCanvas(this.gridCanvas);
    resetCanvas(this.contentCanvas);
    resetCanvas(this.selectionCanvas);
  }

  setInteractionMode(mode: DesktopGridInteractionMode): void {
    this.interactionMode = mode;
    if (mode !== "default") {
      this.cancelFillHandleDrag();
    }
    this.renderer.setFillHandleEnabled(mode === "default" && Boolean(this.callbacks.onFillCommit));
  }

  /**
   * Cancel an in-progress fill handle drag (if any).
   *
   * Returns true when a fill drag was active and was canceled.
   */
  cancelFillHandleDrag(): boolean {
    if (this.dragMode !== "fillHandle") return false;

    this.dragMode = null;
    this.fillHandleState = null;
    this.lastPointerViewport = null;
    this.lastDragPickedRow = null;
    this.lastDragPickedCol = null;
    this.lastFillHandlePointerRow = null;
    this.lastFillHandlePointerCol = null;
    this.selectionAnchor = null;
    this.clearViewportOrigin();
    this.stopAutoScroll();

    const pointerId = this.selectionPointerId;
    this.selectionPointerId = null;

    this.renderer.setFillPreviewRange(null);
    this.selectionCanvas.style.cursor = "default";

    if (pointerId !== null) {
      try {
        this.selectionCanvas.releasePointerCapture?.(pointerId);
      } catch {
        // Ignore capture release failures.
      }
    }

    return true;
  }

  /**
   * Cancel an in-progress row/col resize drag (if any).
   *
   * Returns true when a resize drag was active and was canceled.
   */
  cancelResizeDrag(): boolean {
    const pointerId = this.resizePointerId;
    if (pointerId == null) return false;

    this.resizePointerId = null;
    this.resizeDrag = null;
    this.clearViewportOrigin();
    this.selectionCanvas.style.cursor = "default";

    try {
      this.selectionCanvas.releasePointerCapture?.(pointerId);
    } catch {
      // Ignore capture release failures.
    }

    return true;
  }

  /**
   * Cancel an in-progress selection drag (including formula range-selection drags).
   *
   * Returns true when a selection drag was active and was canceled.
   */
  cancelSelectionDrag(): boolean {
    if (this.dragMode !== "selection") return false;
    const pointerId = this.selectionPointerId;
    if (pointerId == null) return false;

    this.dragMode = null;
    this.selectionPointerId = null;
    this.selectionAnchor = null;
    this.lastPointerViewport = null;
    this.lastDragPickedRow = null;
    this.lastDragPickedCol = null;
    this.lastFillHandlePointerRow = null;
    this.lastFillHandlePointerCol = null;
    this.transientRange = null;
    this.clearViewportOrigin();
    this.stopAutoScroll();

    // Clear any in-progress range-selection overlay (used while editing formulas).
    this.renderer.setRangeSelection(null);
    this.selectionCanvas.style.cursor = "default";

    try {
      this.selectionCanvas.releasePointerCapture?.(pointerId);
    } catch {
      // Ignore capture release failures.
    }

    return true;
  }

  getScroll(): { x: number; y: number } {
    return this.renderer.scroll.getScroll();
  }

  private alignScroll(pos: { x: number; y: number }): { x: number; y: number } {
    return alignScrollToDevicePixels(pos, this.renderer.scroll.getMaxScroll(), this.devicePixelRatio);
  }

  scrollTo(x: number, y: number): void {
    // `CanvasGridRenderer.setScroll` invalidates for scroll unconditionally (even if the scroll
    // position doesn't actually change). Skip calling it when we're already at the requested
    // scroll offsets so we don't trigger redundant render work.
    const before = this.renderer.scroll.getScroll();
    const nextX = Number.isFinite(x) ? x : 0;
    const nextY = Number.isFinite(y) ? y : 0;
    const aligned = this.alignScroll({ x: nextX, y: nextY });
    if (before.x !== aligned.x || before.y !== aligned.y) {
      this.renderer.setScroll(aligned.x, aligned.y);
    }
    this.syncScrollbars();
    this.emitScroll();
  }

  scrollBy(dx: number, dy: number): void {
    // Support `scrollBy(0,0)` as a "sync scrollbars + notify" call site (used by some tests and
    // legacy call sites to force a scrollbar refresh after viewport changes).
    if (dx === 0 && dy === 0) {
      this.syncScrollbars();
      this.emitScroll();
      return;
    }

    const before = this.renderer.scroll.getScroll();
    this.renderer.scrollBy(dx, dy);
    const after = this.renderer.scroll.getScroll();
    if (before.x === after.x && before.y === after.y) return;
    this.syncScrollbars();
    this.emitScroll();
  }

  scrollToCell(row: number, col: number, opts?: { align?: ScrollToCellAlign; padding?: number }): void {
    const before = this.renderer.scroll.getScroll();
    this.renderer.scrollToCell(row, col, opts);
    const after = this.renderer.scroll.getScroll();
    if (before.x === after.x && before.y === after.y) return;
    this.syncScrollbars();
    this.emitScroll();
  }

  getZoom(): number {
    return this.renderer.getZoom();
  }

  setZoom(zoom: number): void {
    const before = this.renderer.getZoom();
    this.renderer.setZoom(zoom);
    if (this.renderer.getZoom() === before) return;
    // `renderer.setZoom()` already schedules a repaint via `markAllDirty()`, but
    // request an additional frame in case the grid is already mid-frame.
    this.renderer.requestRender();
  }

  getCellRect(row: number, col: number): { x: number; y: number; width: number; height: number } | null {
    return this.renderer.getCellRect(row, col);
  }

  getPerfStats(): Readonly<GridPerfStats> {
    return this.renderer.getPerfStats();
  }

  setPerfStatsEnabled(enabled: boolean): void {
    this.renderer.setPerfStatsEnabled(enabled);
    // Perf stats are drawn as part of the grid content layer, which only repaints
    // when there are dirty regions. Force a repaint so toggling via the Ribbon
    // updates immediately even when the grid is otherwise idle.
    this.renderer.markAllDirty();
  }

  setSelectionRanges(
    ranges: CellRange[] | null,
    opts?: {
      activeIndex?: number;
      activeCell?: { row: number; col: number } | null;
      /**
       * When true (default), scroll the active cell into view after updating selection.
       * Set to false for programmatic selection syncing (eg split-view) where each pane
       * should preserve its own scroll position.
       */
      scrollIntoView?: boolean;
    }
  ): void {
    this.keyboardAnchor = null;
    this.transientRange = null;
    this.renderer.setRangeSelection(null);

    const prevSelection = this.renderer.getSelection();
    const prevRange = this.renderer.getSelectionRange();

    this.renderer.setSelectionRanges(ranges, { activeIndex: opts?.activeIndex, activeCell: opts?.activeCell ?? undefined });

    const nextSelection = this.renderer.getSelection();
    const nextRange = this.renderer.getSelectionRange();

    this.announceSelection(nextSelection, nextRange);
    this.emitSelectionChange(prevSelection, nextSelection);
    this.emitSelectionRangeChange(prevRange, nextRange);

    if (nextSelection && opts?.scrollIntoView !== false) {
      this.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
    }
  }

  setRangeSelection(range: CellRange | null): void {
    this.transientRange = range;
    this.renderer.setRangeSelection(range);
    this.announceSelection(this.renderer.getSelection(), range);
  }

  clearRangeSelection(): void {
    this.transientRange = null;
    this.renderer.setRangeSelection(null);
    this.announceSelection(this.renderer.getSelection(), this.renderer.getSelectionRange());
  }

  private announceSelection(selection: { row: number; col: number } | null, range: CellRange | null): void {
    const statusText = describeCellForA11y({
      selection,
      range,
      provider: this.provider,
      headerRows: this.headerRows,
      headerCols: this.headerCols,
    });
    const activeCellLabel = selection ? describeActiveCellLabel(selection, this.provider, this.headerRows, this.headerCols) ?? "" : "";

    if (
      (this.lastAnnounced.selection?.row ?? null) === (selection?.row ?? null) &&
      (this.lastAnnounced.selection?.col ?? null) === (selection?.col ?? null) &&
      rangesEqual(this.lastAnnounced.range, range) &&
      this.lastAnnounced.statusText === statusText &&
      this.lastAnnounced.activeCellLabel === activeCellLabel
    ) {
      return;
    }

    this.lastAnnounced = { selection, range, statusText, activeCellLabel };
    this.a11yStatusEl.textContent = statusText;

    if (!selection) {
      this.container.removeAttribute("aria-activedescendant");
      this.a11yActiveCellEl.setAttribute("aria-hidden", "true");
      this.a11yActiveCellEl.textContent = "";
      this.a11yActiveCellEl.removeAttribute("aria-rowindex");
      this.a11yActiveCellEl.removeAttribute("aria-colindex");
      this.a11yActiveCellEl.removeAttribute("aria-selected");
      return;
    }

    // Ensure the active gridcell exists before referencing it with aria-activedescendant.
    if (!this.a11yActiveCellEl.isConnected) {
      this.container.appendChild(this.a11yActiveCellEl);
    }

    this.a11yActiveCellEl.removeAttribute("aria-hidden");
    this.a11yActiveCellEl.setAttribute("aria-rowindex", String(selection.row + 1));
    this.a11yActiveCellEl.setAttribute("aria-colindex", String(selection.col + 1));
    this.a11yActiveCellEl.setAttribute("aria-selected", "true");
    this.container.setAttribute("aria-activedescendant", this.a11yActiveCellId);

    this.a11yActiveCellEl.textContent = activeCellLabel;
  }

  private emitScroll(): void {
    const viewport = this.renderer.getViewportState();

    const prev = this.lastEmittedViewport;
    if (
      prev &&
      prev.scrollX === viewport.scrollX &&
      prev.scrollY === viewport.scrollY &&
      prev.width === viewport.width &&
      prev.height === viewport.height &&
      prev.maxScrollX === viewport.maxScrollX &&
      prev.maxScrollY === viewport.maxScrollY &&
      prev.frozenRows === viewport.frozenRows &&
      prev.frozenCols === viewport.frozenCols &&
      prev.frozenWidth === viewport.frozenWidth &&
      prev.frozenHeight === viewport.frozenHeight &&
      prev.totalWidth === viewport.totalWidth &&
      prev.totalHeight === viewport.totalHeight &&
      prev.main.rows.start === viewport.main.rows.start &&
      prev.main.rows.end === viewport.main.rows.end &&
      prev.main.rows.offset === viewport.main.rows.offset &&
      prev.main.cols.start === viewport.main.cols.start &&
      prev.main.cols.end === viewport.main.cols.end &&
      prev.main.cols.offset === viewport.main.cols.offset
    ) {
      return;
    }

    this.lastEmittedViewport = viewport;
    this.callbacks.onScroll?.({ x: viewport.scrollX, y: viewport.scrollY }, viewport);
  }

  private emitSelectionChange(
    prev: { row: number; col: number } | null,
    next: { row: number; col: number } | null
  ): void {
    if ((prev?.row ?? null) === (next?.row ?? null) && (prev?.col ?? null) === (next?.col ?? null)) return;
    this.callbacks.onSelectionChange?.(next);
  }

  private emitSelectionRangeChange(prev: CellRange | null, next: CellRange | null): void {
    if (rangesEqual(prev, next)) return;
    this.callbacks.onSelectionRangeChange?.(next);
  }

  /**
   * Re-resolve the grid theme from CSS variables and apply it to the renderer.
   *
   * This is needed because the host app can change `<html data-theme=...>` or
   * system preferences (dark mode / forced-colors) at runtime.
   */
  private refreshTheme(): void {
    const before = this.renderer.getTheme();
    this.renderer.setTheme(resolveGridThemeFromCssVars(this.container));
    // `setTheme` schedules a repaint via rAF when the theme actually changes, but
    // we want shared-grid theme updates (via MutationObserver / matchMedia) to be
    // reflected immediately.
    if (this.renderer.getTheme() !== before) {
      // Some cell providers (e.g. DocumentCellProvider) resolve CSS variables into
      // concrete canvas colors for per-cell styling (hyperlinks, default borders,
      // etc.). When the theme changes, those cached styles need to be recomputed.
      (this.provider as any)?.invalidateAll?.();
      this.renderer.renderImmediately();
    }
  }

  private installThemeWatchers(): void {
    const refreshTheme = () => this.refreshTheme();

    const attributeFilter = ["style", "class", "data-theme", "data-reduced-motion"];
    const observers: MutationObserver[] = [];

    if (typeof MutationObserver !== "undefined") {
      const observe = (el: Element | null) => {
        if (!el) return;
        const observer = new MutationObserver(() => refreshTheme());
        observer.observe(el, { attributes: true, attributeFilter });
        observers.push(observer);
      };

      observe(this.container);

      const doc = this.container.ownerDocument;
      const root = doc?.documentElement;
      if (root && root !== this.container) observe(root);

      const body = doc?.body;
      if (body && body !== this.container && body !== root) observe(body);
    }

    const view = this.container.ownerDocument?.defaultView;
    const canMatchMedia = Boolean(view && typeof view.matchMedia === "function");
    const mqlDark = canMatchMedia ? view!.matchMedia("(prefers-color-scheme: dark)") : null;
    const mqlContrast = canMatchMedia ? view!.matchMedia("(prefers-contrast: more)") : null;
    const mqlForcedColors = canMatchMedia ? view!.matchMedia("(forced-colors: active)") : null;
    const mqlReducedMotion = canMatchMedia ? view!.matchMedia("(prefers-reduced-motion: reduce)") : null;

    const onMediaChange = () => refreshTheme();

    const attachMediaListener = (mql: MediaQueryList | null) => {
      if (!mql) return () => {};
      const legacy = mql as unknown as {
        addListener?: (listener: () => void) => void;
        removeListener?: (listener: () => void) => void;
      };

      if (typeof (mql as any).addEventListener === "function") {
        mql.addEventListener("change", onMediaChange);
        return () => mql.removeEventListener("change", onMediaChange);
      }

      legacy.addListener?.(onMediaChange);
      return () => legacy.removeListener?.(onMediaChange);
    };

    const detachDark = attachMediaListener(mqlDark);
    const detachContrast = attachMediaListener(mqlContrast);
    const detachForced = attachMediaListener(mqlForcedColors);
    const detachReducedMotion = attachMediaListener(mqlReducedMotion);

    this.disposeFns.push(() => {
      for (const observer of observers) observer.disconnect();
      detachDark();
      detachContrast();
      detachForced();
      detachReducedMotion();
    });
  }

  private installWheelHandler(): void {
    const onWheel = (event: WheelEvent) => {
      const target = event.target as HTMLElement | null;
      const renderer = this.renderer;
      const isZoomGesture = event.ctrlKey || event.metaKey;

      // Let the comments panel handle regular scrolling without affecting the grid.
      if (target?.closest?.('[data-testid="comments-panel"]') && !isZoomGesture) return;

      if (isZoomGesture) {
        const viewport = renderer.scroll.getViewportState();
        const startZoom = renderer.getZoom();

        const lineHeight = 16 * startZoom;
        const delta = wheelDeltaToPixels(event.deltaY, event.deltaMode, { lineHeight, pageSize: viewport.height });

        if (delta === 0) return;
        event.preventDefault();

        // Avoid layout reads during high-frequency pinch-zoom: when the wheel event targets one of
        // the full-size canvas layers (positioned at 0,0 in the container), `offsetX/offsetY` are
        // already viewport coords.
        const target = event.target;
        const useOffsets =
          target === this.container || target === this.selectionCanvas || target === this.gridCanvas || target === this.contentCanvas;
        const point = this.hoverViewportPointScratch;
        if (useOffsets) {
          point.x = event.offsetX;
          point.y = event.offsetY;
        } else {
          this.getViewportPoint(event, point);
        }
        const zoomFactor = Math.exp(-delta * 0.001);
        const nextZoom = startZoom * zoomFactor;

        renderer.setZoom(nextZoom, { anchorX: point.x, anchorY: point.y });
        if (renderer.getZoom() === startZoom) return;
        return;
      }

      let deltaX = event.deltaX;
      let deltaY = event.deltaY;

      const viewport = renderer.scroll.getViewportState();
      const lineHeight = 16 * renderer.getZoom();
      const pageWidth = Math.max(0, viewport.width - viewport.frozenWidth);
      const pageHeight = Math.max(0, viewport.height - viewport.frozenHeight);

      deltaX = wheelDeltaToPixels(deltaX, event.deltaMode, { lineHeight, pageSize: pageWidth });
      deltaY = wheelDeltaToPixels(deltaY, event.deltaMode, { lineHeight, pageSize: pageHeight });

      // Common spreadsheet UX: shift+wheel scrolls horizontally.
      if (event.shiftKey) {
        deltaX += deltaY;
        deltaY = 0;
      }

      if (deltaX === 0 && deltaY === 0) return;

      event.preventDefault();
      this.scrollBy(deltaX, deltaY);
    };

    this.container.addEventListener("wheel", onWheel, { passive: false });
    this.disposeFns.push(() => this.container.removeEventListener("wheel", onWheel));
  }

  private installKeyboardHandler(): void {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented) return;

      const target = event.target as HTMLElement | null;
      if (target) {
        const tag = target.tagName;
        // Never steal key events from active text editing (cell editor overlays, inputs in panels, etc).
        if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return;
      }

      // Match primary-grid behavior: Escape cancels an in-progress fill handle drag.
      // (In the primary pane this is handled by SpreadsheetApp.onKeyDown; this shared-grid
      // renderer is used directly by split-view secondary panes.)
      if (event.key === "Escape" && this.cancelFillHandleDrag()) {
        event.preventDefault();
        return;
      }

      if (this.interactionMode === "rangeSelection") return;

      const renderer = this.renderer;
      const selection = renderer.getSelection();
      const { rowCount, colCount } = renderer.scroll.getCounts();
      if (rowCount === 0 || colCount === 0) return;

      const active =
        selection ??
        ({
          row: clampIndex(this.headerRows, 0, rowCount - 1),
          col: clampIndex(this.headerCols, 0, colCount - 1)
        } as const);

      const ctrlOrMeta = event.ctrlKey || event.metaKey;
      const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
      const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

      const applySelectionRange = (range: CellRange) => {
        this.keyboardAnchor = null;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        renderer.setSelectionRange(range, { activeCell: prevSelection ?? active });

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        if (nextSelection) {
          this.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
        }

        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
      };

      // Excel semantics:
      // - F2 edits the active cell.
      // - Shift+F2 adds/edits a comment (handled globally via KeybindingService / CommandRegistry).
      if (event.key === "F2") {
        if (event.shiftKey) return;
        event.preventDefault();
        if (!selection) renderer.setSelection(active);
        this.callbacks.onRequestCellEdit?.({ row: active.row, col: active.col });
        return;
      }

      // Shift+Space selects the entire row, like Excel.
      if (!ctrlOrMeta && !event.altKey && event.shiftKey && (event.code === "Space" || event.key === " ")) {
        event.preventDefault();

        const startCol = this.headerCols >= colCount ? 0 : this.headerCols;
        applySelectionRange({
          startRow: active.row,
          endRow: active.row + 1,
          startCol,
          endCol: colCount
        });
        return;
      }

      // Ctrl/Cmd+Space selects the entire column, like Excel.
      if (ctrlOrMeta && !event.altKey && (event.code === "Space" || event.key === " ")) {
        event.preventDefault();

        const startRow = this.headerRows >= rowCount ? 0 : this.headerRows;
        applySelectionRange({
          startRow,
          endRow: rowCount,
          startCol: active.col,
          endCol: active.col + 1
        });
        return;
      }

      // Ctrl/Cmd+A selects all cells.
      if (ctrlOrMeta && !event.altKey && event.key.toLowerCase() === "a") {
        event.preventDefault();

        const startRow = this.headerRows >= rowCount ? 0 : this.headerRows;
        const startCol = this.headerCols >= colCount ? 0 : this.headerCols;
        applySelectionRange({
          startRow,
          endRow: rowCount,
          startCol,
          endCol: colCount
        });
        return;
      }

      const isPrintable = event.key.length === 1 && !ctrlOrMeta && !event.altKey;
      if (isPrintable) {
        event.preventDefault();
        if (!selection) renderer.setSelection(active);
        this.callbacks.onRequestCellEdit?.({ row: active.row, col: active.col, initialKey: event.key });
        return;
      }

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      const rangeArea = (r: CellRange) => Math.max(0, r.endRow - r.startRow) * Math.max(0, r.endCol - r.startCol);

      // Excel-like behavior: Tab/Enter moves the active cell *within* the current selection range
      // (wrapping) instead of collapsing selection.
      if ((event.key === "Tab" || event.key === "Enter") && prevRange && rangeArea(prevRange) > 1) {
        event.preventDefault();
        this.keyboardAnchor = null;

        const current = prevSelection ?? { row: prevRange.startRow, col: prevRange.startCol };
        const activeRow = clamp(current.row, prevRange.startRow, prevRange.endRow - 1);
        const activeCol = clamp(current.col, prevRange.startCol, prevRange.endCol - 1);
        const backward = event.shiftKey;

        let nextRow = activeRow;
        let nextCol = activeCol;

        if (event.key === "Tab") {
          if (!backward) {
            if (activeCol + 1 < prevRange.endCol) {
              nextCol = activeCol + 1;
            } else if (activeRow + 1 < prevRange.endRow) {
              nextRow = activeRow + 1;
              nextCol = prevRange.startCol;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (activeCol - 1 >= prevRange.startCol) {
              nextCol = activeCol - 1;
            } else if (activeRow - 1 >= prevRange.startRow) {
              nextRow = activeRow - 1;
              nextCol = prevRange.endCol - 1;
            } else {
              nextRow = prevRange.endRow - 1;
              nextCol = prevRange.endCol - 1;
            }
          }
        } else {
          if (!backward) {
            if (activeRow + 1 < prevRange.endRow) {
              nextRow = activeRow + 1;
            } else if (activeCol + 1 < prevRange.endCol) {
              nextRow = prevRange.startRow;
              nextCol = activeCol + 1;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (activeRow - 1 >= prevRange.startRow) {
              nextRow = activeRow - 1;
            } else if (activeCol - 1 >= prevRange.startCol) {
              nextRow = prevRange.endRow - 1;
              nextCol = activeCol - 1;
            } else {
              nextRow = prevRange.endRow - 1;
              nextCol = prevRange.endCol - 1;
            }
          }
        }

        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        renderer.setSelectionRanges(ranges, { activeIndex, activeCell: { row: nextRow, col: nextCol } });

        this.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        return;
      }

      let nextRow = active.row;
      let nextCol = active.col;
      let handled = true;

      const viewport = renderer.scroll.getViewportState();
      const pageRows = Math.max(1, viewport.main.rows.end - viewport.main.rows.start);
      const pageCols = Math.max(1, viewport.main.cols.end - viewport.main.cols.start);

      switch (event.key) {
        case "ArrowUp":
          nextRow = ctrlOrMeta ? dataStartRow : active.row - 1;
          break;
        case "ArrowDown":
          nextRow = ctrlOrMeta ? rowCount - 1 : active.row + 1;
          break;
        case "ArrowLeft":
          nextCol = ctrlOrMeta ? dataStartCol : active.col - 1;
          break;
        case "ArrowRight":
          nextCol = ctrlOrMeta ? colCount - 1 : active.col + 1;
          break;
        case "PageUp":
          if (event.altKey) {
            nextCol = active.col - pageCols;
          } else {
            nextRow = active.row - pageRows;
          }
          break;
        case "PageDown":
          if (event.altKey) {
            nextCol = active.col + pageCols;
          } else {
            nextRow = active.row + pageRows;
          }
          break;
        case "Home":
          if (ctrlOrMeta) {
            nextRow = dataStartRow;
            nextCol = dataStartCol;
          } else {
            nextCol = dataStartCol;
          }
          break;
        case "End":
          if (ctrlOrMeta) {
            nextRow = rowCount - 1;
            nextCol = colCount - 1;
          } else {
            nextCol = colCount - 1;
          }
          break;
        case "Enter":
          nextRow = active.row + (event.shiftKey ? -1 : 1);
          break;
        case "Tab":
          // Excel-style:
          // - Tab moves right; at the end of the row, wrap to the first column of the next row.
          // - Shift+Tab moves left; at the start of the row, wrap to the last column of the previous row.
          if (event.shiftKey) {
            if (active.col - 1 >= dataStartCol) {
              nextCol = active.col - 1;
            } else if (active.row - 1 >= dataStartRow) {
              nextRow = active.row - 1;
              nextCol = colCount - 1;
            } else {
              nextRow = active.row;
              nextCol = active.col;
            }
          } else {
            if (active.col + 1 < colCount) {
              nextCol = active.col + 1;
            } else if (active.row + 1 < rowCount) {
              nextRow = active.row + 1;
              nextCol = dataStartCol;
            } else {
              nextRow = active.row;
              nextCol = active.col;
            }
          }
          break;
        default:
          handled = false;
      }

      if (!handled) return;
      event.preventDefault();

      nextRow = Math.max(dataStartRow, Math.min(rowCount - 1, nextRow));
      nextCol = Math.max(dataStartCol, Math.min(colCount - 1, nextCol));

      const extendSelection = event.shiftKey && event.key !== "Tab" && event.key !== "Enter";

      if (extendSelection) {
        const anchor = this.keyboardAnchor ?? prevSelection ?? active;
        if (!this.keyboardAnchor) this.keyboardAnchor = anchor;

        const range: CellRange = {
          startRow: Math.min(anchor.row, nextRow),
          endRow: Math.max(anchor.row, nextRow) + 1,
          startCol: Math.min(anchor.col, nextCol),
          endCol: Math.max(anchor.col, nextCol) + 1
        };

        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        const updatedRanges = ranges.length === 0 ? [range] : ranges;
        updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
        renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell: { row: nextRow, col: nextCol } });
      } else {
        this.keyboardAnchor = null;
        renderer.setSelection({ row: nextRow, col: nextCol });
      }

      this.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      this.announceSelection(nextSelection, nextRange);
      this.emitSelectionChange(prevSelection, nextSelection);
      this.emitSelectionRangeChange(prevRange, nextRange);
    };

    this.container.addEventListener("keydown", onKeyDown);
    this.disposeFns.push(() => this.container.removeEventListener("keydown", onKeyDown));
  }

  private stopAutoScroll(): void {
    if (this.autoScrollFrame == null) return;
    cancelAnimationFrame(this.autoScrollFrame);
    this.autoScrollFrame = null;
  }

  private scheduleAutoScroll(): void {
    if (this.autoScrollFrame != null) return;

    const clamp01 = (value: number) => Math.max(0, Math.min(1, value));
    const edge = 28;
    const maxSpeed = 24;

    const tick = () => {
      this.autoScrollFrame = null;
      if (this.selectionPointerId == null) return;
      const point = this.lastPointerViewport;
      if (!point) return;

      const renderer = this.renderer;
      const viewport = renderer.scroll.getViewportState();
      if (viewport.width <= 0 || viewport.height <= 0) return;

      const leftThreshold = viewport.frozenWidth + edge;
      const topThreshold = viewport.frozenHeight + edge;

      const leftFactor = clamp01((leftThreshold - point.x) / edge);
      const rightFactor = clamp01((point.x - (viewport.width - edge)) / edge);
      const topFactor = clamp01((topThreshold - point.y) / edge);
      const bottomFactor = clamp01((point.y - (viewport.height - edge)) / edge);

      const dx = viewport.maxScrollX > 0 ? (rightFactor - leftFactor) * maxSpeed : 0;
      const dy = viewport.maxScrollY > 0 ? (bottomFactor - topFactor) * maxSpeed : 0;

      if (dx === 0 && dy === 0) return;

      const before = renderer.scroll.getScroll();
      renderer.scrollBy(dx, dy);
      const after = renderer.scroll.getScroll();

      if (before.x === after.x && before.y === after.y) return;

      this.syncScrollbars();
      this.emitScroll();

      const clampedX = Math.max(0, Math.min(viewport.width, point.x));
      const clampedY = Math.max(0, Math.min(viewport.height, point.y));
      const picked = renderer.pickCellAt(clampedX, clampedY, this.pickCellScratch);
      if (picked) {
        if (this.dragMode === "fillHandle") this.applyFillHandleDrag(picked);
        else this.applyDragRange(picked);
      }

      this.autoScrollFrame = requestAnimationFrame(tick);
    };

    this.autoScrollFrame = requestAnimationFrame(tick);
  }

  private applyDragRange(picked: { row: number; col: number }): void {
    const renderer = this.renderer;
    const anchor = this.selectionAnchor;
    if (!anchor) return;
    if (this.lastDragPickedRow === picked.row && this.lastDragPickedCol === picked.col) return;
    this.lastDragPickedRow = picked.row;
    this.lastDragPickedCol = picked.col;

    const startRow = Math.min(anchor.row, picked.row);
    const endRow = Math.max(anchor.row, picked.row) + 1;
    const startCol = Math.min(anchor.col, picked.col);
    const endCol = Math.max(anchor.col, picked.col) + 1;

    if (this.interactionMode === "rangeSelection") {
      const prev = this.transientRange;
      if (
        prev &&
        prev.startRow === startRow &&
        prev.endRow === endRow &&
        prev.startCol === startCol &&
        prev.endCol === endCol
      ) {
        return;
      }

      const range: CellRange = { startRow, endRow, startCol, endCol };
      this.transientRange = range;
      renderer.setRangeSelection(range);
      this.announceSelection(renderer.getSelection(), range);
      this.callbacks.onRangeSelectionChange?.(range);
      return;
    }

    const range = this.selectionDragRangeScratch;
    range.startRow = startRow;
    range.endRow = endRow;
    range.startCol = startCol;
    range.endCol = endCol;
    if (!renderer.setActiveSelectionRange(range)) return;

    const nextSelection = renderer.getSelection();
    const nextRange = renderer.getSelectionRange();
    this.announceSelection(nextSelection, nextRange);
    this.callbacks.onSelectionRangeChange?.(nextRange);
  }

  private applyFillHandleDrag(picked: { row: number; col: number }): void {
    const state = this.fillHandleState;
    if (!state) return;
    const { rowCount, colCount } = this.renderer.scroll.getCounts();
    if (rowCount === 0 || colCount === 0) return;
    const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
    const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

    // Clamp pointerCell into the data region so fill handle drags can't extend selection into headers.
    const pointerRow = clamp(picked.row, dataStartRow, rowCount - 1);
    const pointerCol = clamp(picked.col, dataStartCol, colCount - 1);
    if (pointerRow === this.lastFillHandlePointerRow && pointerCol === this.lastFillHandlePointerCol) {
      return;
    }
    this.lastFillHandlePointerRow = pointerRow;
    this.lastFillHandlePointerCol = pointerCol;

    const pointerCell = this.fillHandlePointerCellScratch;
    pointerCell.row = pointerRow;
    pointerCell.col = pointerCol;
    const preview = computeFillPreview(state.source, pointerCell, this.fillDragPreviewScratch);
    const union = preview ? this.fillDragPreviewScratch.unionRange : state.source;
    const targetRange = preview ? this.fillDragPreviewScratch.targetRange : null;

    const unionStartRow = union.startRow;
    const unionEndRow = union.endRow;
    const unionStartCol = union.startCol;
    const unionEndCol = union.endCol;

    const endRow = clamp(picked.row, unionStartRow, unionEndRow - 1);
    const endCol = clamp(picked.col, unionStartCol, unionEndCol - 1);

    if (
      state.target.startRow === unionStartRow &&
      state.target.endRow === unionEndRow &&
      state.target.startCol === unionStartCol &&
      state.target.endCol === unionEndCol &&
      endRow === state.endCell.row &&
      endCol === state.endCell.col
    ) {
      return;
    }

    state.target.startRow = unionStartRow;
    state.target.endRow = unionEndRow;
    state.target.startCol = unionStartCol;
    state.target.endCol = unionEndCol;
    if (targetRange) {
      state.previewTarget.startRow = targetRange.startRow;
      state.previewTarget.endRow = targetRange.endRow;
      state.previewTarget.startCol = targetRange.startCol;
      state.previewTarget.endCol = targetRange.endCol;
    }
    state.endCell.row = endRow;
    state.endCell.col = endCol;
    this.renderer.setFillPreviewRange(state.target);
  }

  private cacheViewportOrigin(): { left: number; top: number } {
    const rect = this.selectionCanvas.getBoundingClientRect();
    const origin = { left: rect.left, top: rect.top };
    this.selectionCanvasViewportOrigin = origin;
    return origin;
  }

  private clearViewportOrigin(): void {
    this.selectionCanvasViewportOrigin = null;
  }

  private getViewportPoint(
    event: { clientX: number; clientY: number },
    out?: { x: number; y: number }
  ): { x: number; y: number } {
    const origin = this.selectionCanvasViewportOrigin;
    let x: number;
    let y: number;
    if (origin) {
      x = event.clientX - origin.left;
      y = event.clientY - origin.top;
    } else {
      const rect = this.selectionCanvas.getBoundingClientRect();
      x = event.clientX - rect.left;
      y = event.clientY - rect.top;
    }

    if (out) {
      out.x = x;
      out.y = y;
      return out;
    }

    return { x, y };
  }

  private getResizeHit(viewportX: number, viewportY: number): ResizeHit | null {
    const renderer = this.renderer;
    const viewport = renderer.scroll.getViewportState();
    const { rowCount, colCount } = renderer.scroll.getCounts();
    if (rowCount === 0 || colCount === 0) return null;

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

    const absScrollX = viewport.frozenWidth + viewport.scrollX;
    const absScrollY = viewport.frozenHeight + viewport.scrollY;

    const colAxis = renderer.scroll.cols;
    const rowAxis = renderer.scroll.rows;

    const headerRowsFrozen = Math.min(this.headerRows, viewport.frozenRows);
    const headerColsFrozen = Math.min(this.headerCols, viewport.frozenCols);
    const headerHeight = rowAxis.totalSize(headerRowsFrozen);
    const headerWidth = colAxis.totalSize(headerColsFrozen);

    const inHeaderRow = headerRowsFrozen > 0 && viewportY >= 0 && viewportY <= Math.min(headerHeight, viewport.height);
    const inRowHeaderCol = headerColsFrozen > 0 && viewportX >= 0 && viewportX <= Math.min(headerWidth, viewport.width);

    const RESIZE_HIT_RADIUS_PX = 4;

    let best: (ResizeHit & { distance: number }) | null = null;

    if (inHeaderRow) {
      const inFrozenCols = viewportX < frozenWidthClamped;
      const sheetX = inFrozenCols ? viewportX : absScrollX + (viewportX - frozenWidthClamped);
      const minCol = inFrozenCols ? 0 : viewport.frozenCols;
      const maxColInclusive = inFrozenCols ? viewport.frozenCols - 1 : colCount - 1;

      if (maxColInclusive >= minCol) {
        const col = colAxis.indexAt(sheetX, { min: minCol, maxInclusive: maxColInclusive });
        const colStart = colAxis.positionOf(col);
        const colEnd = colStart + colAxis.getSize(col);

        const distToStart = Math.abs(sheetX - colStart);
        const distToEnd = Math.abs(sheetX - colEnd);

        if (distToStart <= RESIZE_HIT_RADIUS_PX && col > 0) {
          best = { kind: "col", index: col - 1, distance: distToStart };
        } else if (distToEnd <= RESIZE_HIT_RADIUS_PX) {
          best = { kind: "col", index: col, distance: distToEnd };
        }
      }
    }

    if (inRowHeaderCol) {
      const inFrozenRows = viewportY < frozenHeightClamped;
      const sheetY = inFrozenRows ? viewportY : absScrollY + (viewportY - frozenHeightClamped);
      const minRow = inFrozenRows ? 0 : viewport.frozenRows;
      const maxRowInclusive = inFrozenRows ? viewport.frozenRows - 1 : rowCount - 1;

      if (maxRowInclusive >= minRow) {
        const row = rowAxis.indexAt(sheetY, { min: minRow, maxInclusive: maxRowInclusive });
        const rowStart = rowAxis.positionOf(row);
        const rowEnd = rowStart + rowAxis.getSize(row);

        const distToStart = Math.abs(sheetY - rowStart);
        const distToEnd = Math.abs(sheetY - rowEnd);

        let candidate: (ResizeHit & { distance: number }) | null = null;
        if (distToStart <= RESIZE_HIT_RADIUS_PX && row > 0) {
          candidate = { kind: "row", index: row - 1, distance: distToStart };
        } else if (distToEnd <= RESIZE_HIT_RADIUS_PX) {
          candidate = { kind: "row", index: row, distance: distToEnd };
        }

        if (candidate && (!best || candidate.distance < best.distance)) {
          best = candidate;
        }
      }
    }

    return best ? { kind: best.kind, index: best.index } : null;
  }

  private installPointerHandlers(options: { enableResize: boolean }): void {
    const selectionCanvas = this.selectionCanvas;
    const isMacPlatform = (() => {
      try {
        const platform = typeof navigator !== "undefined" ? navigator.platform : "";
        return /Mac|iPhone|iPad|iPod/.test(platform);
      } catch {
        return false;
      }
    })();

    const MIN_COL_WIDTH = 24;
    const MIN_ROW_HEIGHT = 16;

    const onPointerDown = (event: PointerEvent) => {
      const renderer = this.renderer;

      // Excel/Sheets behavior: right-clicking inside an existing selection keeps the
      // selection intact; right-clicking outside moves the active cell to the clicked
      // cell. We intentionally only support sheet cells for now (not row/col header
      // context menus).
      //
      // Note: On macOS, Ctrl+click is commonly treated as a right click and fires the
      // `contextmenu` event. Ensure we treat it as a context-click (not additive selection).
      const isContextClick =
        event.pointerType === "mouse" &&
        (event.button === 2 || (isMacPlatform && event.button === 0 && event.ctrlKey && !event.metaKey));
      if (isContextClick) {
        // Drawings/pictures are rendered under the shared-grid selection canvas and
        // handle their own right-click selection logic. When a context-click hit a
        // drawing, DrawingInteractionController tags the pointer event so we can
        // avoid moving the active cell underneath the drawing (Excel-like behavior).
        if ((event as any).__formulaDrawingContextClick) {
          return;
        }

        const point = this.getViewportPoint(event);
        const picked = renderer.pickCellAt(point.x, point.y);
        if (!picked) return;

        const isHeaderCell = (this.headerRows > 0 && picked.row < this.headerRows) || (this.headerCols > 0 && picked.col < this.headerCols);
        if (isHeaderCell) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        const ranges = renderer.getSelectionRanges();
        const inSelection = ranges.some(
          (range) =>
            picked.row >= range.startRow &&
            picked.row < range.endRow &&
            picked.col >= range.startCol &&
            picked.col < range.endCol
        );
        if (!inSelection) {
          renderer.setSelection(picked);
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();
        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);

        // Best-effort: keep focus on the grid so keyboard navigation continues.
        try {
          (this.container as any).focus?.({ preventScroll: true });
        } catch {
          this.container.focus?.();
        }
        return;
      }

      event.preventDefault();
      this.keyboardAnchor = null;
      this.dragMode = null;
      this.fillHandleState = null;
      renderer.setFillPreviewRange(null);
      this.cacheViewportOrigin();
      const point = this.getViewportPoint(event);
      this.lastPointerViewport = point;

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          this.resizePointerId = event.pointerId;
          selectionCanvas.setPointerCapture?.(event.pointerId);

          if (hit.kind === "col") {
            this.resizeDrag = { kind: "col", index: hit.index, startClient: event.clientX, startSize: renderer.getColWidth(hit.index) };
            selectionCanvas.style.cursor = "col-resize";
          } else {
            this.resizeDrag = { kind: "row", index: hit.index, startClient: event.clientY, startSize: renderer.getRowHeight(hit.index) };
            selectionCanvas.style.cursor = "row-resize";
          }
          return;
        }
      }

      if (this.interactionMode === "default" && this.callbacks.onFillCommit) {
        if (hitTestSelectionHandle(renderer, point.x, point.y)) {
          const source = renderer.getSelectionRange();
          if (source) {
            this.selectionPointerId = event.pointerId;
            this.dragMode = "fillHandle";
            this.selectionAnchor = null;
            const mode: FillMode = event.altKey ? "formulas" : event.metaKey || event.ctrlKey ? "copy" : "series";
            const target: CellRange = { ...source };
            const previewTarget: CellRange = { ...source };
            this.fillHandleState = {
              source,
              target,
              mode,
              previewTarget,
              endCell: { row: source.endRow - 1, col: source.endCol - 1 }
            };
            // Seed the fill-handle pointer-cell cache with the selection corner so high-frequency
            // pointermoves within the same cell don't recompute previews.
            const { rowCount, colCount } = renderer.scroll.getCounts();
            const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
            const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;
            this.lastFillHandlePointerRow = clamp(source.endRow - 1, dataStartRow, Math.max(0, rowCount - 1));
            this.lastFillHandlePointerCol = clamp(source.endCol - 1, dataStartCol, Math.max(0, colCount - 1));
            selectionCanvas.setPointerCapture?.(event.pointerId);

            // Focus the container so keyboard shortcuts (e.g. Escape) still work while dragging.
            try {
              (this.container as any).focus?.({ preventScroll: true });
            } catch {
              this.container.focus?.();
            }

            this.transientRange = null;
            renderer.setRangeSelection(null);

            renderer.setFillPreviewRange(source);
            this.scheduleAutoScroll();
            return;
          }
        }
      }

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) {
        this.clearViewportOrigin();
        return;
      }

      if (this.interactionMode === "rangeSelection") {
        this.selectionPointerId = event.pointerId;
        this.dragMode = "selection";
        selectionCanvas.setPointerCapture?.(event.pointerId);

        this.selectionAnchor = picked;
        this.lastDragPickedRow = picked.row;
        this.lastDragPickedCol = picked.col;
        const range: CellRange = { startRow: picked.row, endRow: picked.row + 1, startCol: picked.col, endCol: picked.col + 1 };
        this.transientRange = range;
        renderer.setRangeSelection(range);
        this.announceSelection(renderer.getSelection(), range);
        this.callbacks.onRangeSelectionStart?.(range);
        this.scheduleAutoScroll();
        return;
      }

      // Focus the container so keyboard navigation continues after mouse interaction.
      try {
        (this.container as any).focus?.({ preventScroll: true });
      } catch {
        this.container.focus?.();
      }

      this.transientRange = null;
      renderer.setRangeSelection(null);

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      // Ctrl/Cmd+click on a URL-like cell value should open it externally instead
      // of being treated as an additive selection gesture.
      if (this.interactionMode === "default" && (event.metaKey || event.ctrlKey) && event.button === 0) {
        const cell = this.provider.getCell(picked.row, picked.col);
        const raw = (cell as any)?.value;
        if (typeof raw === "string" && looksLikeExternalHyperlink(raw)) {
          const uri = raw.trim();

          // Match normal click behavior (make the clicked cell active) while still
          // allowing the OS browser open behavior behind Ctrl/Cmd.
          this.selectionAnchor = picked;
          renderer.setSelection(picked);

          const nextSelection = renderer.getSelection();
          const nextRange = renderer.getSelectionRange();
          this.announceSelection(nextSelection, nextRange);
          this.emitSelectionChange(prevSelection, nextSelection);
          this.emitSelectionRangeChange(prevRange, nextRange);

          void openExternalHyperlink(uri, {
            shellOpen,
            confirmUntrustedProtocol: async (message) => {
              return nativeDialogs.confirm(message);
            },
          }).catch(() => {
            // Best-effort: link opening should not crash grid interaction.
          });
          return;
        }
      }

      const isAdditive = event.metaKey || event.ctrlKey;
      const isExtend = event.shiftKey;

      const { rowCount, colCount } = renderer.scroll.getCounts();
      const viewport = renderer.scroll.getViewportState();
      const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
      const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

      const applyHeaderRange = (range: CellRange, activeCell: { row: number; col: number }) => {
        if (isAdditive) {
          const existing = renderer.getSelectionRanges();
          const nextRanges = [...existing, range];
          renderer.setSelectionRanges(nextRanges, { activeIndex: nextRanges.length - 1, activeCell });
          return;
        }

        if (isExtend && prevSelection) {
          const existing = renderer.getSelectionRanges();
          const activeIndex = renderer.getActiveSelectionIndex();
          const updatedRanges = existing.length === 0 ? [range] : existing;
          updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
          renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell });
          return;
        }

        renderer.setSelectionRange(range, { activeCell });
      };

      const isCornerHeader =
        this.headerRows > 0 && this.headerCols > 0 && picked.row < this.headerRows && picked.col < this.headerCols;
      const isColumnHeader = this.headerRows > 0 && picked.row < this.headerRows && picked.col >= this.headerCols;
      const isRowHeader = this.headerCols > 0 && picked.col < this.headerCols && picked.row >= this.headerRows;

      if (isCornerHeader || isColumnHeader || isRowHeader) {
        if (isCornerHeader) {
          const range: CellRange = { startRow: dataStartRow, endRow: rowCount, startCol: dataStartCol, endCol: colCount };
          const activeCell =
            prevSelection ??
            ({
              row: Math.max(dataStartRow, viewport.main.rows.start),
              col: Math.max(dataStartCol, viewport.main.cols.start)
            } as const);
          applyHeaderRange(range, activeCell);
        } else if (isColumnHeader) {
          const anchorCol = prevSelection ? clamp(prevSelection.col, dataStartCol, colCount - 1) : picked.col;
          const startCol = isExtend && prevSelection ? Math.min(anchorCol, picked.col) : picked.col;
          const endCol = (isExtend && prevSelection ? Math.max(anchorCol, picked.col) : picked.col) + 1;

          const range: CellRange = {
            startRow: dataStartRow,
            endRow: rowCount,
            startCol,
            endCol: Math.min(colCount, endCol)
          };

          const baseRow = prevSelection ? prevSelection.row : Math.max(dataStartRow, viewport.main.rows.start);
          applyHeaderRange(range, { row: baseRow, col: picked.col });
        } else {
          const anchorRow = prevSelection ? clamp(prevSelection.row, dataStartRow, rowCount - 1) : picked.row;
          const startRow = isExtend && prevSelection ? Math.min(anchorRow, picked.row) : picked.row;
          const endRow = (isExtend && prevSelection ? Math.max(anchorRow, picked.row) : picked.row) + 1;

          const range: CellRange = {
            startRow,
            endRow: Math.min(rowCount, endRow),
            startCol: dataStartCol,
            endCol: colCount
          };

          const baseCol = prevSelection ? prevSelection.col : Math.max(dataStartCol, viewport.main.cols.start);
          applyHeaderRange(range, { row: picked.row, col: baseCol });
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
        this.scheduleAutoScroll();
        this.clearViewportOrigin();
        return;
      }

      this.selectionPointerId = event.pointerId;
      this.dragMode = "selection";
      selectionCanvas.setPointerCapture?.(event.pointerId);
      this.lastDragPickedRow = picked.row;
      this.lastDragPickedCol = picked.col;

      if (isAdditive) {
        this.selectionAnchor = picked;
        renderer.addSelectionRange({ startRow: picked.row, endRow: picked.row + 1, startCol: picked.col, endCol: picked.col + 1 });
      } else if (isExtend && prevSelection) {
        this.selectionAnchor = prevSelection;
        const range: CellRange = {
          startRow: Math.min(prevSelection.row, picked.row),
          endRow: Math.max(prevSelection.row, picked.row) + 1,
          startCol: Math.min(prevSelection.col, picked.col),
          endCol: Math.max(prevSelection.col, picked.col) + 1
        };
        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        const updatedRanges = ranges.length === 0 ? [range] : ranges;
        updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
        renderer.setSelectionRanges(updatedRanges, { activeIndex });
      } else {
        this.selectionAnchor = picked;
        renderer.setSelection(picked);
      }

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      this.announceSelection(nextSelection, nextRange);
      this.emitSelectionChange(prevSelection, nextSelection);
      this.emitSelectionRangeChange(prevRange, nextRange);

      this.scheduleAutoScroll();
    };

    const onPointerMove = (event: PointerEvent) => {
      const renderer = this.renderer;

      if (this.resizePointerId != null) {
        if (event.pointerId !== this.resizePointerId) return;
        const drag = this.resizeDrag;
        if (!drag) return;
        event.preventDefault();

        if (drag.kind === "col") {
          const delta = event.clientX - drag.startClient;
          renderer.setColWidth(drag.index, Math.max(MIN_COL_WIDTH * renderer.getZoom(), drag.startSize + delta));
        } else {
          const delta = event.clientY - drag.startClient;
          renderer.setRowHeight(drag.index, Math.max(MIN_ROW_HEIGHT * renderer.getZoom(), drag.startSize + delta));
        }

        return;
      }

      if (this.selectionPointerId == null) return;
      if (event.pointerId !== this.selectionPointerId) return;
      event.preventDefault();

      const point = this.getViewportPoint(event, this.dragViewportPointScratch);
      this.lastPointerViewport = point;

      const picked = renderer.pickCellAt(point.x, point.y, this.pickCellScratch);
      if (!picked) return;
      if (this.dragMode === "fillHandle") this.applyFillHandleDrag(picked);
      else this.applyDragRange(picked);
      this.scheduleAutoScroll();
    };

    const endDrag = (event: PointerEvent) => {
      const renderer = this.renderer;
      if (this.resizePointerId != null && event.pointerId === this.resizePointerId) {
        const drag = this.resizeDrag;
        this.resizePointerId = null;
        this.resizeDrag = null;
        this.clearViewportOrigin();
        selectionCanvas.style.cursor = "default";
        try {
          selectionCanvas.releasePointerCapture?.(event.pointerId);
        } catch {
          // Ignore.
        }

        if (drag) {
          const endSize = drag.kind === "col" ? this.renderer.getColWidth(drag.index) : this.renderer.getRowHeight(drag.index);
          const defaultSize = drag.kind === "col" ? this.renderer.scroll.cols.defaultSize : this.renderer.scroll.rows.defaultSize;
          if (endSize !== drag.startSize) {
            this.callbacks.onAxisSizeChange?.({
              kind: drag.kind,
              index: drag.index,
              size: endSize,
              previousSize: drag.startSize,
              defaultSize,
              zoom: this.renderer.getZoom(),
              source: "resize"
            });
          }
        }
      }

      if (this.selectionPointerId == null) return;
      if (event.pointerId !== this.selectionPointerId) return;

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();
      const dragMode = this.dragMode;
      this.dragMode = null;

      this.selectionPointerId = null;
      this.selectionAnchor = null;
      this.lastPointerViewport = null;
      this.lastDragPickedRow = null;
      this.lastDragPickedCol = null;
      this.lastFillHandlePointerRow = null;
      this.lastFillHandlePointerCol = null;
      this.clearViewportOrigin();
      this.stopAutoScroll();

      if (dragMode === "fillHandle") {
        const state = this.fillHandleState;
        this.fillHandleState = null;
        renderer.setFillPreviewRange(null);

        const shouldCommit = event.type === "pointerup";
        if (state && shouldCommit && !rangesEqual(state.source, state.target)) {
          const commitResult = this.callbacks.onFillCommit?.({
            sourceRange: state.source,
            targetRange: state.previewTarget,
            mode: state.mode
          });
          void Promise.resolve(commitResult).catch(() => {
            // Consumers own commit error handling; swallow to avoid unhandled rejections.
          });

          const ranges = renderer.getSelectionRanges();
          const activeIndex = renderer.getActiveSelectionIndex();
          const updatedRanges = ranges.length === 0 ? [state.target] : [...ranges];
          updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = state.target;
          renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell: state.endCell });
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();
        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
      } else {
        this.fillHandleState = null;
        renderer.setFillPreviewRange(null);

        if (this.interactionMode === "rangeSelection") {
          const range = this.transientRange;
          if (range) this.callbacks.onRangeSelectionEnd?.(range);
        }
      }

      selectionCanvas.style.cursor = "default";

      try {
        selectionCanvas.releasePointerCapture?.(event.pointerId);
      } catch {
        // Ignore.
      }
    };

    const onPointerHover = (event: PointerEvent) => {
      if (this.resizePointerId != null || this.selectionPointerId != null) return;
      // Avoid layout reads during high-frequency hover events by using `offsetX/offsetY`
      // when the event targets the selection canvas (viewport coords in this case).
      const useOffsets =
        event.target === selectionCanvas && Number.isFinite(event.offsetX) && Number.isFinite(event.offsetY);
      const point = this.hoverViewportPointScratch;
      if (useOffsets) {
        point.x = event.offsetX;
        point.y = event.offsetY;
      } else {
        this.getViewportPoint(event, point);
      }

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          selectionCanvas.style.cursor = hit.kind === "col" ? "col-resize" : "row-resize";
          return;
        }
      }

      if (this.interactionMode === "default" && this.callbacks.onFillCommit) {
        if (hitTestSelectionHandle(this.renderer, point.x, point.y)) {
          selectionCanvas.style.cursor = "crosshair";
          return;
        }
      }

      selectionCanvas.style.cursor = "default";
    };

    const onPointerLeave = () => {
      if (this.resizePointerId != null || this.selectionPointerId != null) return;
      selectionCanvas.style.cursor = "default";
    };

    const onDoubleClick = (event: MouseEvent) => {
      if (this.interactionMode === "rangeSelection") return;
      const renderer = this.renderer;
      const point = this.getViewportPoint(event);

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          event.preventDefault();
          const prevSize = hit.kind === "col" ? renderer.getColWidth(hit.index) : renderer.getRowHeight(hit.index);
          const defaultSize = hit.kind === "col" ? renderer.scroll.cols.defaultSize : renderer.scroll.rows.defaultSize;
          const nextSize =
            hit.kind === "col" ? renderer.autoFitCol(hit.index, { maxWidth: 500 }) : renderer.autoFitRow(hit.index, { maxHeight: 500 });

          if (nextSize !== prevSize) {
            this.callbacks.onAxisSizeChange?.({
              kind: hit.kind,
              index: hit.index,
              size: nextSize,
              previousSize: prevSize,
              defaultSize,
              zoom: renderer.getZoom(),
              source: "autoFit"
            });
          }
          return;
        }
      }

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;
      this.callbacks.onRequestCellEdit?.({ row: picked.row, col: picked.col });
    };

    selectionCanvas.addEventListener("pointerdown", onPointerDown);
    selectionCanvas.addEventListener("pointermove", onPointerMove);
    selectionCanvas.addEventListener("pointermove", onPointerHover);
    selectionCanvas.addEventListener("pointerleave", onPointerLeave);
    selectionCanvas.addEventListener("pointerup", endDrag);
    selectionCanvas.addEventListener("pointercancel", endDrag);
    selectionCanvas.addEventListener("dblclick", onDoubleClick);

    this.disposeFns.push(() => {
      selectionCanvas.removeEventListener("pointerdown", onPointerDown);
      selectionCanvas.removeEventListener("pointermove", onPointerMove);
      selectionCanvas.removeEventListener("pointermove", onPointerHover);
      selectionCanvas.removeEventListener("pointerleave", onPointerLeave);
      selectionCanvas.removeEventListener("pointerup", endDrag);
      selectionCanvas.removeEventListener("pointercancel", endDrag);
      selectionCanvas.removeEventListener("dblclick", onDoubleClick);
    });
  }

  private installScrollbarHandlers(): void {
    const onVThumbDown = (event: PointerEvent) => {
      event.preventDefault();
      event.stopPropagation();

      const renderer = this.renderer;
      const trackRect = this.vTrack.getBoundingClientRect();
      const thumbRect = this.vThumb.getBoundingClientRect();
      const maxScroll = renderer.scroll.getViewportState().maxScrollY;
      if (maxScroll === 0) return;

      const grabOffset = event.clientY - thumbRect.top;
      const thumbTravel = Math.max(0, trackRect.height - thumbRect.height);
      const pointerId = event.pointerId;

      let cleanedUp = false;
      let cleanup = () => {};

      const onMove = (move: PointerEvent) => {
        if (move.pointerId !== pointerId) return;
        move.preventDefault();
        const pointerPos = move.clientY;
        const thumbOffset = pointerPos - trackRect.top - grabOffset;
        const clamped = clamp(thumbOffset, 0, thumbTravel);
        const nextScroll = thumbTravel === 0 ? 0 : (clamped / thumbTravel) * maxScroll;
        const before = renderer.scroll.getScroll();
        const aligned = this.alignScroll({ x: before.x, y: nextScroll });
        if (aligned.y === before.y) return;
        renderer.setScroll(aligned.x, aligned.y);
        this.syncScrollbars();
        this.emitScroll();
      };

      const onUp = (up: PointerEvent) => {
        if (up.pointerId !== pointerId) return;
        cleanup();
      };

      const onCancel = (cancel: PointerEvent) => {
        if (cancel.pointerId !== pointerId) return;
        cleanup();
      };

      cleanup = () => {
        if (cleanedUp) return;
        cleanedUp = true;
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        window.removeEventListener("pointercancel", onCancel);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
      window.addEventListener("pointercancel", onCancel, { passive: false });
    };

    this.vThumb.addEventListener("pointerdown", onVThumbDown, { passive: false });
    this.disposeFns.push(() => this.vThumb.removeEventListener("pointerdown", onVThumbDown));

    const onHThumbDown = (event: PointerEvent) => {
      event.preventDefault();
      event.stopPropagation();

      const renderer = this.renderer;
      const trackRect = this.hTrack.getBoundingClientRect();
      const thumbRect = this.hThumb.getBoundingClientRect();
      const maxScroll = renderer.scroll.getViewportState().maxScrollX;
      if (maxScroll === 0) return;

      const grabOffset = event.clientX - thumbRect.left;
      const thumbTravel = Math.max(0, trackRect.width - thumbRect.width);
      const pointerId = event.pointerId;

      let cleanedUp = false;
      let cleanup = () => {};

      const onMove = (move: PointerEvent) => {
        if (move.pointerId !== pointerId) return;
        move.preventDefault();
        const pointerPos = move.clientX;
        const thumbOffset = pointerPos - trackRect.left - grabOffset;
        const clamped = clamp(thumbOffset, 0, thumbTravel);
        const nextScroll = thumbTravel === 0 ? 0 : (clamped / thumbTravel) * maxScroll;
        const before = renderer.scroll.getScroll();
        const aligned = this.alignScroll({ x: nextScroll, y: before.y });
        if (aligned.x === before.x) return;
        renderer.setScroll(aligned.x, aligned.y);
        this.syncScrollbars();
        this.emitScroll();
      };

      const onUp = (up: PointerEvent) => {
        if (up.pointerId !== pointerId) return;
        cleanup();
      };

      const onCancel = (cancel: PointerEvent) => {
        if (cancel.pointerId !== pointerId) return;
        cleanup();
      };

      cleanup = () => {
        if (cleanedUp) return;
        cleanedUp = true;
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        window.removeEventListener("pointercancel", onCancel);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
      window.addEventListener("pointercancel", onCancel, { passive: false });
    };

    this.hThumb.addEventListener("pointerdown", onHThumbDown, { passive: false });
    this.disposeFns.push(() => this.hThumb.removeEventListener("pointerdown", onHThumbDown));

    const onVTrackDown = (event: PointerEvent) => {
      if (event.target !== this.vTrack) return;
      const renderer = this.renderer;
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollY = viewport.maxScrollY;
      const trackRect = this.vTrack.getBoundingClientRect();
      const minThumbSize = 24 * renderer.getZoom();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().y,
        viewportSize: Math.max(0, viewport.height - viewport.frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - viewport.frozenHeight),
        trackSize: trackRect.height,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.height - thumb.size);
      if (thumbTravel === 0 || maxScrollY === 0) return;

      const pointerPos = event.clientY - trackRect.top;
      const targetOffset = pointerPos - thumb.size / 2;
      const clampedOffset = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clampedOffset / thumbTravel) * maxScrollY;

      const before = renderer.scroll.getScroll();
      const aligned = this.alignScroll({ x: before.x, y: nextScroll });
      if (aligned.y === before.y) return;
      renderer.setScroll(aligned.x, aligned.y);
      this.syncScrollbars();
      this.emitScroll();
    };

    this.vTrack.addEventListener("pointerdown", onVTrackDown, { passive: false });
    this.disposeFns.push(() => this.vTrack.removeEventListener("pointerdown", onVTrackDown));

    const onHTrackDown = (event: PointerEvent) => {
      if (event.target !== this.hTrack) return;
      const renderer = this.renderer;
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollX = viewport.maxScrollX;
      const trackRect = this.hTrack.getBoundingClientRect();
      const minThumbSize = 24 * renderer.getZoom();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().x,
        viewportSize: Math.max(0, viewport.width - viewport.frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - viewport.frozenWidth),
        trackSize: trackRect.width,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.width - thumb.size);
      if (thumbTravel === 0 || maxScrollX === 0) return;

      const pointerPos = event.clientX - trackRect.left;
      const targetOffset = pointerPos - thumb.size / 2;
      const clampedOffset = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clampedOffset / thumbTravel) * maxScrollX;

      const before = renderer.scroll.getScroll();
      const aligned = this.alignScroll({ x: nextScroll, y: before.y });
      if (aligned.x === before.x) return;
      renderer.setScroll(aligned.x, aligned.y);
      this.syncScrollbars();
      this.emitScroll();
    };

    this.hTrack.addEventListener("pointerdown", onHTrackDown, { passive: false });
    this.disposeFns.push(() => this.hTrack.removeEventListener("pointerdown", onHTrackDown));
  }

  syncScrollbars(): void {
    const renderer = this.renderer;
    const viewport = renderer.scroll.getViewportState();
    const scroll = renderer.scroll.getScroll();
    const minThumbSize = 24 * renderer.getZoom();

    const maxX = viewport.maxScrollX;
    const maxY = viewport.maxScrollY;
    const showH = maxX > 0;
    const showV = maxY > 0;

    const padding = 2;
    const thickness = 10;

    const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);

    const prev = this.lastScrollbarLayout;
    const layoutChanged =
      prev === null ||
      prev.showV !== showV ||
      prev.showH !== showH ||
      prev.frozenWidth !== frozenWidth ||
      prev.frozenHeight !== frozenHeight;

    if (layoutChanged) {
      // Track layout (positioning/visibility) is a function of viewport/frozen sizes and whether
      // each axis is scrollable. Avoid rewriting these styles on every scroll event.
      this.vTrack.style.display = showV ? "block" : "none";
      this.hTrack.style.display = showH ? "block" : "none";

      if (showV) {
        this.vTrack.style.right = `${padding}px`;
        this.vTrack.style.top = `${frozenHeight + padding}px`;
        this.vTrack.style.bottom = `${(showH ? thickness : 0) + padding}px`;
        this.vTrack.style.width = `${thickness}px`;
      }

      if (showH) {
        this.hTrack.style.left = `${frozenWidth + padding}px`;
        this.hTrack.style.right = `${(showV ? thickness : 0) + padding}px`;
        this.hTrack.style.bottom = `${padding}px`;
        this.hTrack.style.height = `${thickness}px`;
      }

      if (prev) {
        prev.showV = showV;
        prev.showH = showH;
        prev.frozenWidth = frozenWidth;
        prev.frozenHeight = frozenHeight;
      } else {
        this.lastScrollbarLayout = { showV, showH, frozenWidth, frozenHeight };
      }
    }

    if (showV) {
      // Avoid layout reads during scroll; the track size is deterministic from the
      // viewport + the same top/bottom offsets applied above.
      const trackSize = Math.max(0, viewport.height - frozenHeight - (showH ? thickness : 0) - 2 * padding);
      const thumb = computeScrollbarThumb({
        scrollPos: scroll.y,
        viewportSize: Math.max(0, viewport.height - frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - frozenHeight),
        trackSize,
        minThumbSize,
        out: this.scrollbarThumbScratch.v
      });

      if (this.lastScrollbarThumb.vSize !== thumb.size) {
        this.vThumb.style.height = `${thumb.size}px`;
        this.lastScrollbarThumb.vSize = thumb.size;
      }
      if (this.lastScrollbarThumb.vOffset !== thumb.offset) {
        this.vThumb.style.transform = `translateY(${thumb.offset}px)`;
        this.lastScrollbarThumb.vOffset = thumb.offset;
      }
    } else {
      this.lastScrollbarThumb.vSize = null;
      this.lastScrollbarThumb.vOffset = null;
    }

    if (showH) {
      // Avoid layout reads during scroll; the track size is deterministic from the
      // viewport + the same left/right offsets applied above.
      const trackSize = Math.max(0, viewport.width - frozenWidth - (showV ? thickness : 0) - 2 * padding);
      const thumb = computeScrollbarThumb({
        scrollPos: scroll.x,
        viewportSize: Math.max(0, viewport.width - frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - frozenWidth),
        trackSize,
        minThumbSize,
        out: this.scrollbarThumbScratch.h
      });

      if (this.lastScrollbarThumb.hSize !== thumb.size) {
        this.hThumb.style.width = `${thumb.size}px`;
        this.lastScrollbarThumb.hSize = thumb.size;
      }
      if (this.lastScrollbarThumb.hOffset !== thumb.offset) {
        this.hThumb.style.transform = `translateX(${thumb.offset}px)`;
        this.lastScrollbarThumb.hOffset = thumb.offset;
      }
    } else {
      this.lastScrollbarThumb.hSize = null;
      this.lastScrollbarThumb.hOffset = null;
    }
  }

  resize(width: number, height: number, devicePixelRatio: number): void {
    this.devicePixelRatio = devicePixelRatio;
    this.renderer.resize(width, height, devicePixelRatio);
    if (this.selectionCanvasViewportOrigin) {
      this.cacheViewportOrigin();
    }
  }
}
