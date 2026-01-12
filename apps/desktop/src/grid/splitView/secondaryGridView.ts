import type { CellRange, GridAxisSizeChange } from "@formula/grid";
import type { DocumentController } from "../../document/documentController.js";
import { DesktopSharedGrid, type DesktopSharedGridCallbacks } from "../shared/desktopSharedGrid.js";
import { DocumentCellProvider } from "../shared/documentCellProvider.js";

type ScrollState = { scrollX: number; scrollY: number };

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, value));
}

/**
 * Mounts a shared CanvasGrid-based renderer into an arbitrary container.
 *
 * This is used by the desktop split-view secondary pane.
 */
export class SecondaryGridView {
  readonly container: HTMLElement;
  readonly provider: DocumentCellProvider;
  readonly grid: DesktopSharedGrid;

  private readonly document: DocumentController;
  private readonly getSheetId: () => string;
  private readonly headerRows = 1;
  private readonly headerCols = 1;

  private sheetId: string;

  private readonly resizeObserver: ResizeObserver;
  private readonly disposeFns: Array<() => void> = [];

  private readonly persistDebounceMs: number;
  private pendingScroll: ScrollState | null = null;
  private scrollPersistTimer: number | null = null;
  private pendingZoom: number | null = null;
  private zoomPersistTimer: number | null = null;
  private lastZoom = 1;

  private readonly persistScroll?: (scroll: ScrollState) => void;
  private readonly persistZoom?: (zoom: number) => void;

  constructor(options: {
    container: HTMLElement;
    /**
     * Optional provider to reuse (e.g. the primary shared-grid provider).
     *
     * When supplied, `document`/`getSheetId`/`showFormulas`/`getComputedValue` are
     * ignored for provider creation, but still accepted for convenience.
     */
    provider?: DocumentCellProvider;
    document: DocumentController;
    getSheetId: () => string;
    rowCount: number;
    colCount: number;
    showFormulas: () => boolean;
    getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
    onSelectionChange?: (selection: { row: number; col: number } | null) => void;
    onSelectionRangeChange?: (range: CellRange | null) => void;
    getCommentMeta?: (row: number, col: number) => { resolved: boolean } | null;
    callbacks?: DesktopSharedGridCallbacks;
    initialScroll?: ScrollState;
    initialZoom?: number;
    persistScroll?: (scroll: ScrollState) => void;
    persistZoom?: (zoom: number) => void;
    persistDebounceMs?: number;
  }) {
    this.container = options.container;
    this.document = options.document;
    this.getSheetId = options.getSheetId;
    this.persistScroll = options.persistScroll;
    this.persistZoom = options.persistZoom;
    this.persistDebounceMs = options.persistDebounceMs ?? 150;
    this.sheetId = options.getSheetId();

    // Clear any placeholder content from the split-view scaffolding.
    this.container.replaceChildren();

    const gridCanvas = document.createElement("canvas");
    gridCanvas.className = "grid-canvas";
    gridCanvas.setAttribute("aria-hidden", "true");

    const contentCanvas = document.createElement("canvas");
    contentCanvas.className = "grid-canvas";
    contentCanvas.setAttribute("aria-hidden", "true");

    const selectionCanvas = document.createElement("canvas");
    selectionCanvas.className = "grid-canvas";
    selectionCanvas.setAttribute("aria-hidden", "true");

    this.container.appendChild(gridCanvas);
    this.container.appendChild(contentCanvas);
    this.container.appendChild(selectionCanvas);

    const vTrack = document.createElement("div");
    vTrack.setAttribute("aria-hidden", "true");
    vTrack.setAttribute("data-testid", "scrollbar-track-y-secondary");
    vTrack.className = "grid-scrollbar-track grid-scrollbar-track--vertical";

    const vThumb = document.createElement("div");
    vThumb.setAttribute("aria-hidden", "true");
    vThumb.setAttribute("data-testid", "scrollbar-thumb-y-secondary");
    vThumb.className = "grid-scrollbar-thumb";
    vTrack.appendChild(vThumb);
    this.container.appendChild(vTrack);

    const hTrack = document.createElement("div");
    hTrack.setAttribute("aria-hidden", "true");
    hTrack.setAttribute("data-testid", "scrollbar-track-x-secondary");
    hTrack.className = "grid-scrollbar-track grid-scrollbar-track--horizontal";

    const hThumb = document.createElement("div");
    hThumb.setAttribute("aria-hidden", "true");
    hThumb.setAttribute("data-testid", "scrollbar-thumb-x-secondary");
    hThumb.className = "grid-scrollbar-thumb";
    hTrack.appendChild(hThumb);
    this.container.appendChild(hTrack);

    this.provider =
      options.provider ??
      new DocumentCellProvider({
        document: options.document,
        getSheetId: options.getSheetId,
        headerRows: this.headerRows,
        headerCols: this.headerCols,
        rowCount: options.rowCount,
        colCount: options.colCount,
        showFormulas: options.showFormulas,
        getComputedValue: options.getComputedValue,
        getCommentMeta: options.getCommentMeta
      });

    const externalCallbacks: DesktopSharedGridCallbacks = options.callbacks ?? {};

    this.grid = new DesktopSharedGrid({
      container: this.container,
      provider: this.provider,
      rowCount: options.rowCount,
      colCount: options.colCount,
      frozenRows: this.headerRows,
      frozenCols: this.headerCols,
      defaultRowHeight: 24,
      defaultColWidth: 100,
      enableResize: true,
      enableKeyboard: false,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      callbacks: {
        ...externalCallbacks,
        onScroll: (scroll, viewport) => {
          this.container.dataset.scrollX = String(scroll.x);
          this.container.dataset.scrollY = String(scroll.y);
          this.schedulePersistScroll({ scrollX: scroll.x, scrollY: scroll.y });

          const zoom = this.grid.renderer.getZoom();
          this.container.dataset.zoom = String(zoom);
          if (Math.abs(zoom - this.lastZoom) > 1e-6) {
            this.lastZoom = zoom;
            this.schedulePersistZoom(zoom);
          }

          externalCallbacks.onScroll?.(scroll, viewport);
        },
        onAxisSizeChange: (change) => {
          this.onAxisSizeChange(change);
          externalCallbacks.onAxisSizeChange?.(change);
        },
        onSelectionChange: (selection) => {
          options.onSelectionChange?.(selection);
          externalCallbacks.onSelectionChange?.(selection);
        },
        onSelectionRangeChange: (range) => {
          options.onSelectionRangeChange?.(range);
          externalCallbacks.onSelectionRangeChange?.(range);
        },
      }
    });

    // Match SpreadsheetApp header sizing so cell hit targets line up.
    this.grid.renderer.setColWidth(0, 48);
    this.grid.renderer.setRowHeight(0, 24);

    // Initial sizing (ResizeObserver will keep it updated).
    this.resizeToContainer();

    const initialZoom = clamp(options.initialZoom ?? 1, 0.25, 4);
    if (initialZoom !== 1) {
      this.grid.renderer.setZoom(initialZoom);
      this.grid.syncScrollbars();
    }
    this.lastZoom = this.grid.renderer.getZoom();
    this.container.dataset.zoom = String(this.lastZoom);

    // Apply frozen panes + axis overrides from the DocumentController sheet view.
    this.syncSheetViewFromDocument();

    const initialScroll = options.initialScroll ?? { scrollX: 0, scrollY: 0 };
    if (initialScroll.scrollX !== 0 || initialScroll.scrollY !== 0) {
      this.grid.scrollTo(initialScroll.scrollX, initialScroll.scrollY);
    } else {
      // Still set the dataset so tests can read a stable baseline.
      const scroll = this.grid.getScroll();
      this.container.dataset.scrollX = String(scroll.x);
      this.container.dataset.scrollY = String(scroll.y);
    }

    // Re-apply view state when the document emits sheet view deltas (freeze panes, row/col sizes).
    const unsubscribeSheetView = this.document.on("change", (payload: any) => {
      const deltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
      if (deltas.length === 0) return;
      const sheetId = this.getSheetId();
      if (!deltas.some((delta: any) => String(delta?.sheetId ?? "") === sheetId)) return;
      this.syncSheetViewFromDocument();
    });
    this.disposeFns.push(() => unsubscribeSheetView());
    this.resizeObserver = new ResizeObserver(() => this.resizeToContainer());
    this.resizeObserver.observe(this.container);
  }

  destroy(): void {
    this.flushPersistence();
    this.resizeObserver.disconnect();
    for (const dispose of this.disposeFns) dispose();
    this.disposeFns.length = 0;
    this.grid.destroy();
    // Remove any DOM we created (the container stays in place).
    this.container.replaceChildren();
  }

  private resizeToContainer(): void {
    const width = this.container.clientWidth;
    const height = this.container.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    this.grid.resize(width, height, dpr);
  }

  /**
   * Sync frozen panes + row/col size overrides from `DocumentController.getSheetView()`
   * into the secondary pane renderer.
   */
  syncSheetViewFromDocument(): void {
    const sheetId = this.getSheetId();
    if (sheetId !== this.sheetId) {
      this.sheetId = sheetId;
    }

    const view = this.document.getSheetView(sheetId) as {
      frozenRows?: number;
      frozenCols?: number;
      colWidths?: Record<string, number>;
      rowHeights?: Record<string, number>;
    } | null;

    const maxRows = Math.max(0, this.grid.renderer.scroll.getCounts().rowCount - this.headerRows);
    const maxCols = Math.max(0, this.grid.renderer.scroll.getCounts().colCount - this.headerCols);

    const normalizeFrozen = (value: unknown, max: number): number => {
      const num = Number(value);
      if (!Number.isFinite(num)) return 0;
      return Math.max(0, Math.min(Math.trunc(num), max));
    };

    const frozenRows = normalizeFrozen(view?.frozenRows, maxRows);
    const frozenCols = normalizeFrozen(view?.frozenCols, maxCols);
    this.grid.renderer.setFrozen(this.headerRows + frozenRows, this.headerCols + frozenCols);

    const zoom = this.grid.renderer.getZoom();

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

    // Batch apply to avoid N-per-index invalidations when many overrides exist.
    const colSizes = new Map<number, number>();
    for (let i = 0; i < this.headerCols; i += 1) {
      colSizes.set(i, this.grid.renderer.getColWidth(i));
    }
    for (const [col, base] of nextCols) {
      if (col >= maxCols) continue;
      colSizes.set(col + this.headerCols, base * zoom);
    }

    const rowSizes = new Map<number, number>();
    for (let i = 0; i < this.headerRows; i += 1) {
      rowSizes.set(i, this.grid.renderer.getRowHeight(i));
    }
    for (const [row, base] of nextRows) {
      if (row >= maxRows) continue;
      rowSizes.set(row + this.headerRows, base * zoom);
    }

    this.grid.renderer.applyAxisSizeOverrides({ rows: rowSizes, cols: colSizes }, { resetUnspecified: true });

    this.grid.syncScrollbars();
    const scroll = this.grid.getScroll();
    this.container.dataset.scrollX = String(scroll.x);
    this.container.dataset.scrollY = String(scroll.y);
  }

  private onAxisSizeChange(change: GridAxisSizeChange): void {
    const sheetId = this.getSheetId();
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;

    if (change.kind === "col") {
      const docCol = change.index - this.headerCols;
      if (docCol < 0) return;
      const label = change.source === "autoFit" ? "Autofit Column Width" : "Resize Column";
      if (isDefault) {
        this.document.resetColWidth(sheetId, docCol, { label });
      } else {
        this.document.setColWidth(sheetId, docCol, baseSize, { label });
      }
      return;
    }

    const docRow = change.index - this.headerRows;
    if (docRow < 0) return;
    const label = change.source === "autoFit" ? "Autofit Row Height" : "Resize Row";
    if (isDefault) {
      this.document.resetRowHeight(sheetId, docRow, { label });
    } else {
      this.document.setRowHeight(sheetId, docRow, baseSize, { label });
    }
  }

  private schedulePersistScroll(scroll: ScrollState): void {
    if (!this.persistScroll) return;
    this.pendingScroll = scroll;

    if (this.scrollPersistTimer != null) {
      window.clearTimeout(this.scrollPersistTimer);
    }
    this.scrollPersistTimer = window.setTimeout(() => {
      this.scrollPersistTimer = null;
      const pending = this.pendingScroll;
      this.pendingScroll = null;
      if (!pending) return;
      this.persistScroll?.(pending);
    }, this.persistDebounceMs);
  }

  private schedulePersistZoom(zoom: number): void {
    if (!this.persistZoom) return;
    this.pendingZoom = zoom;

    if (this.zoomPersistTimer != null) {
      window.clearTimeout(this.zoomPersistTimer);
    }
    this.zoomPersistTimer = window.setTimeout(() => {
      this.zoomPersistTimer = null;
      const pending = this.pendingZoom;
      this.pendingZoom = null;
      if (pending == null) return;
      this.persistZoom?.(pending);
    }, this.persistDebounceMs);
  }

  private flushPersistence(): void {
    if (this.scrollPersistTimer != null) {
      window.clearTimeout(this.scrollPersistTimer);
      this.scrollPersistTimer = null;
    }
    if (this.zoomPersistTimer != null) {
      window.clearTimeout(this.zoomPersistTimer);
      this.zoomPersistTimer = null;
    }

    if (this.pendingScroll) {
      this.persistScroll?.(this.pendingScroll);
      this.pendingScroll = null;
    }
    if (this.pendingZoom != null) {
      this.persistZoom?.(this.pendingZoom);
      this.pendingZoom = null;
    }
  }
}
