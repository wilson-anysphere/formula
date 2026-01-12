import type { DocumentController } from "../../document/documentController.js";
import { DesktopSharedGrid } from "../shared/desktopSharedGrid.js";
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

  private readonly resizeObserver: ResizeObserver;
  private readonly disposeFns: Array<() => void> = [];

  private readonly persistDebounceMs: number;
  private pendingScroll: ScrollState | null = null;
  private scrollPersistTimer: number | null = null;
  private pendingZoom: number | null = null;
  private zoomPersistTimer: number | null = null;

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
    getCommentMeta?: (cellRef: string) => { resolved: boolean } | null;
    initialScroll?: ScrollState;
    initialZoom?: number;
    persistScroll?: (scroll: ScrollState) => void;
    persistZoom?: (zoom: number) => void;
    persistDebounceMs?: number;
  }) {
    this.container = options.container;
    this.persistScroll = options.persistScroll;
    this.persistZoom = options.persistZoom;
    this.persistDebounceMs = options.persistDebounceMs ?? 150;

    // Match SpreadsheetApp behaviour: keep everything clipped to the grid viewport.
    this.container.style.overflow = "hidden";

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
    vTrack.style.position = "absolute";
    vTrack.style.background = "var(--bg-tertiary)";
    vTrack.style.borderRadius = "6px";
    vTrack.style.zIndex = "5";
    vTrack.style.opacity = "0.9";

    const vThumb = document.createElement("div");
    vThumb.setAttribute("aria-hidden", "true");
    vThumb.setAttribute("data-testid", "scrollbar-thumb-y-secondary");
    vThumb.style.position = "absolute";
    vThumb.style.left = "1px";
    vThumb.style.right = "1px";
    vThumb.style.top = "0";
    vThumb.style.height = "40px";
    vThumb.style.background = "var(--text-secondary)";
    vThumb.style.borderRadius = "6px";
    vThumb.style.cursor = "pointer";
    vTrack.appendChild(vThumb);
    this.container.appendChild(vTrack);

    const hTrack = document.createElement("div");
    hTrack.setAttribute("aria-hidden", "true");
    hTrack.setAttribute("data-testid", "scrollbar-track-x-secondary");
    hTrack.style.position = "absolute";
    hTrack.style.background = "var(--bg-tertiary)";
    hTrack.style.borderRadius = "6px";
    hTrack.style.zIndex = "5";
    hTrack.style.opacity = "0.9";

    const hThumb = document.createElement("div");
    hThumb.setAttribute("aria-hidden", "true");
    hThumb.setAttribute("data-testid", "scrollbar-thumb-x-secondary");
    hThumb.style.position = "absolute";
    hThumb.style.top = "1px";
    hThumb.style.bottom = "1px";
    hThumb.style.left = "0";
    hThumb.style.width = "40px";
    hThumb.style.background = "var(--text-secondary)";
    hThumb.style.borderRadius = "6px";
    hThumb.style.cursor = "pointer";
    hTrack.appendChild(hThumb);
    this.container.appendChild(hTrack);

    const headerRows = 1;
    const headerCols = 1;

    this.provider =
      options.provider ??
      new DocumentCellProvider({
        document: options.document,
        getSheetId: options.getSheetId,
        headerRows,
        headerCols,
        rowCount: options.rowCount,
        colCount: options.colCount,
        showFormulas: options.showFormulas,
        getComputedValue: options.getComputedValue,
        getCommentMeta: options.getCommentMeta
      });

    this.grid = new DesktopSharedGrid({
      container: this.container,
      provider: this.provider,
      rowCount: options.rowCount,
      colCount: options.colCount,
      frozenRows: headerRows,
      frozenCols: headerCols,
      defaultRowHeight: 24,
      defaultColWidth: 100,
      enableResize: true,
      enableKeyboard: false,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      callbacks: {
        onScroll: (scroll) => {
          this.container.dataset.scrollX = String(scroll.x);
          this.container.dataset.scrollY = String(scroll.y);
          this.schedulePersistScroll({ scrollX: scroll.x, scrollY: scroll.y });
        }
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
    this.container.dataset.zoom = String(this.grid.renderer.getZoom());

    const initialScroll = options.initialScroll ?? { scrollX: 0, scrollY: 0 };
    if (initialScroll.scrollX !== 0 || initialScroll.scrollY !== 0) {
      this.grid.scrollTo(initialScroll.scrollX, initialScroll.scrollY);
    } else {
      // Still set the dataset so tests can read a stable baseline.
      const scroll = this.grid.getScroll();
      this.container.dataset.scrollX = String(scroll.x);
      this.container.dataset.scrollY = String(scroll.y);
    }

    this.installZoomHandler();

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

  private installZoomHandler(): void {
    const onWheel = (event: WheelEvent) => {
      const primary = event.ctrlKey || event.metaKey;
      if (!primary) return;
      // Prevent browser page zoom.
      event.preventDefault();
      event.stopPropagation();

      let deltaY = event.deltaY;
      if (event.deltaMode === 1) {
        deltaY *= 16;
      } else if (event.deltaMode === 2) {
        // Pages -> approximate with the viewport height.
        deltaY *= Math.max(1, this.container.getBoundingClientRect().height);
      }

      // Smooth exponential scale: ~10% per wheel "notch" (deltaY=100).
      const factor = Math.pow(1.001, -deltaY);
      const current = this.grid.renderer.getZoom();
      const next = clamp(current * factor, 0.25, 4);

      const rect = this.container.getBoundingClientRect();
      const anchorX = event.clientX - rect.left;
      const anchorY = event.clientY - rect.top;

      this.grid.renderer.setZoom(next, { anchorX, anchorY });
      this.grid.syncScrollbars();

      this.container.dataset.zoom = String(this.grid.renderer.getZoom());
      const scroll = this.grid.getScroll();
      this.container.dataset.scrollX = String(scroll.x);
      this.container.dataset.scrollY = String(scroll.y);

      this.schedulePersistZoom(this.grid.renderer.getZoom());
    };

    this.container.addEventListener("wheel", onWheel, { passive: false });
    this.disposeFns.push(() => this.container.removeEventListener("wheel", onWheel));
  }

  private resizeToContainer(): void {
    const width = this.container.clientWidth;
    const height = this.container.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    this.grid.resize(width, height, dpr);
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
