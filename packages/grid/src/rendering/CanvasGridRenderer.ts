import type { CellData, CellProvider, CellProviderUpdate, CellRange } from "../model/CellProvider";
import { DirtyRegionTracker, type Rect } from "./DirtyRegionTracker";
import { setupHiDpiCanvas } from "./HiDpiCanvas";
import { LruCache } from "../utils/LruCache";
import type { GridPresence } from "../presence/types";
import type { GridViewportState } from "../virtualization/VirtualScrollManager";
import { VirtualScrollManager } from "../virtualization/VirtualScrollManager";
import {
  TextLayoutEngine,
  createCanvasTextMeasurer,
  detectBaseDirection,
  drawTextLayout,
  resolveAlign,
  toCanvasFontString
} from "@formula/text-layout";

type Layer = "background" | "content" | "selection";

interface Selection {
  row: number;
  col: number;
}

function isSameCellRange(a: CellRange | null, b: CellRange | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.startRow === b.startRow && a.endRow === b.endRow && a.startCol === b.startCol && a.endCol === b.endCol;
}

function intersectRect(a: Rect, b: Rect): Rect | null {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.width, b.x + b.width);
  const y2 = Math.min(a.y + a.height, b.y + b.height);
  const width = x2 - x1;
  const height = y2 - y1;
  if (width <= 0 || height <= 0) return null;
  return { x: x1, y: y1, width, height };
}

function crispLine(pos: number): number {
  return Math.round(pos) + 0.5;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return clamp(Math.trunc(value), min, max);
}

function padRect(rect: Rect, padding: number): Rect {
  return { x: rect.x - padding, y: rect.y - padding, width: rect.width + padding * 2, height: rect.height + padding * 2 };
}

function parseHexColor(color: string): { r: number; g: number; b: number } | null {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return null;
  const value = Number.parseInt(match[1], 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255
  };
}

function pickTextColor(backgroundColor: string): string {
  const rgb = parseHexColor(backgroundColor);
  if (!rgb) return "#ffffff";
  const luma = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
  return luma > 0.6 ? "#000000" : "#ffffff";
}

const DEFAULT_ERROR_TEXT_COLOR = "#cc0000";

export function formatCellDisplayText(value: CellData["value"]): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function resolveCellTextColor(value: CellData["value"], explicitColor: string | undefined): string {
  if (explicitColor !== undefined) return explicitColor;
  if (typeof value === "string" && value.startsWith("#")) return DEFAULT_ERROR_TEXT_COLOR;
  return "#111111";
}

export class CanvasGridRenderer {
  private readonly provider: CellProvider;
  readonly scroll: VirtualScrollManager;

  private readonly prefetchOverscanRows: number;
  private readonly prefetchOverscanCols: number;
  private lastPrefetchRanges: CellRange[] | null = null;

  private gridCanvas?: HTMLCanvasElement;
  private gridCtx?: CanvasRenderingContext2D;
  private contentCanvas?: HTMLCanvasElement;
  private contentCtx?: CanvasRenderingContext2D;
  private selectionCanvas?: HTMLCanvasElement;
  private selectionCtx?: CanvasRenderingContext2D;

  private blitCanvas?: HTMLCanvasElement;
  private blitCtx?: CanvasRenderingContext2D;

  private unsubscribeProvider?: () => void;

  private devicePixelRatio = 1;

  private readonly dirty = {
    background: new DirtyRegionTracker(),
    content: new DirtyRegionTracker(),
    selection: new DirtyRegionTracker()
  };

  private scheduled = false;
  private forceFullRedraw = true;

  private lastRendered = {
    width: 0,
    height: 0,
    frozenRows: 0,
    frozenCols: 0,
    frozenWidth: 0,
    frozenHeight: 0,
    scrollX: 0,
    scrollY: 0,
    devicePixelRatio: 1
  };

  private selection: Selection | null = null;
  private selectionRange: CellRange | null = null;
  private rangeSelection: CellRange | null = null;

  private remotePresences: GridPresence[] = [];

  private readonly formattedCache = new LruCache<string, string>(50_000);
  private readonly textWidthCache = new LruCache<string, number>(10_000);
  private textLayoutEngine?: TextLayoutEngine;

  private readonly presenceFont = "12px system-ui, sans-serif";

  constructor(options: {
    provider: CellProvider;
    rowCount: number;
    colCount: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
    prefetchOverscanRows?: number;
    prefetchOverscanCols?: number;
  }) {
    this.provider = options.provider;
    this.prefetchOverscanRows = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanRows);
    this.prefetchOverscanCols = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanCols);
    this.scroll = new VirtualScrollManager({
      rowCount: options.rowCount,
      colCount: options.colCount,
      defaultRowHeight: options.defaultRowHeight,
      defaultColWidth: options.defaultColWidth
    });
  }

  attach(canvases: {
    grid: HTMLCanvasElement;
    content: HTMLCanvasElement;
    selection: HTMLCanvasElement;
  }): void {
    this.gridCanvas = canvases.grid;
    this.gridCtx = canvases.grid.getContext("2d", { alpha: false }) ?? undefined;

    this.contentCanvas = canvases.content;
    this.contentCtx = canvases.content.getContext("2d") ?? undefined;

    this.selectionCanvas = canvases.selection;
    this.selectionCtx = canvases.selection.getContext("2d") ?? undefined;

    if (!this.gridCtx || !this.contentCtx || !this.selectionCtx) {
      throw new Error("Failed to acquire canvas 2D contexts.");
    }

    if (!this.textLayoutEngine) {
      // Dedicated measurer canvas avoids mutating the render context state during layout.
      this.textLayoutEngine = new TextLayoutEngine(createCanvasTextMeasurer(), {
        maxMeasureCacheEntries: 50_000,
        maxLayoutCacheEntries: 10_000
      });
    }

    this.gridCanvas.style.position = "absolute";
    this.contentCanvas.style.position = "absolute";
    this.selectionCanvas.style.position = "absolute";

    this.gridCanvas.style.left = "0";
    this.gridCanvas.style.top = "0";
    this.contentCanvas.style.left = "0";
    this.contentCanvas.style.top = "0";
    this.selectionCanvas.style.left = "0";
    this.selectionCanvas.style.top = "0";

    this.gridCanvas.style.zIndex = "0";
    this.contentCanvas.style.zIndex = "1";
    this.selectionCanvas.style.zIndex = "2";

    if (this.unsubscribeProvider) {
      this.unsubscribeProvider();
      this.unsubscribeProvider = undefined;
    }

    if (this.provider.subscribe) {
      this.unsubscribeProvider = this.provider.subscribe((update) => this.onProviderUpdate(update));
    }

    this.markAllDirty();
  }

  destroy(): void {
    if (this.unsubscribeProvider) {
      this.unsubscribeProvider();
      this.unsubscribeProvider = undefined;
    }
  }

  resize(width: number, height: number, devicePixelRatio: number): void {
    if (!this.gridCanvas || !this.contentCanvas || !this.selectionCanvas) return;
    if (!this.gridCtx || !this.contentCtx || !this.selectionCtx) return;

    this.devicePixelRatio = devicePixelRatio;
    this.scroll.setViewportSize(width, height);

    setupHiDpiCanvas(this.gridCanvas, this.gridCtx, width, height, devicePixelRatio);
    setupHiDpiCanvas(this.contentCanvas, this.contentCtx, width, height, devicePixelRatio);
    setupHiDpiCanvas(this.selectionCanvas, this.selectionCtx, width, height, devicePixelRatio);

    this.ensureBlitCanvas();
    this.markAllDirty();
  }

  setFrozen(frozenRows: number, frozenCols: number): void {
    this.scroll.setFrozen(frozenRows, frozenCols);
    this.markAllDirty();
  }

  setScroll(scrollX: number, scrollY: number): void {
    this.scroll.setScroll(scrollX, scrollY);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);
    this.invalidateForScroll();
  }

  scrollBy(deltaX: number, deltaY: number): void {
    const before = this.scroll.getScroll();
    this.scroll.scrollBy(deltaX, deltaY);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);
    const after = this.scroll.getScroll();
    if (before.x !== after.x || before.y !== after.y) this.invalidateForScroll();
  }

  setSelection(selection: Selection | null): void {
    const previousRange = this.selectionRange;
    this.selection = selection;
    this.selectionRange = selection
      ? {
          startRow: selection.row,
          endRow: selection.row + 1,
          startCol: selection.col,
          endCol: selection.col + 1
        }
      : null;

    this.invalidateSelection(previousRange, this.selectionRange);
  }

  setSelectionRange(range: CellRange | null): void {
    const previousRange = this.selectionRange;
    const normalized = range ? this.normalizeSelectionRange(range) : null;

    if (!normalized) {
      this.selection = null;
      this.selectionRange = null;
      this.invalidateSelection(previousRange, null);
      return;
    }

    const active = this.selection ?? { row: normalized.startRow, col: normalized.startCol };
    this.selection = {
      row: clamp(active.row, normalized.startRow, normalized.endRow - 1),
      col: clamp(active.col, normalized.startCol, normalized.endCol - 1)
    };
    this.selectionRange = normalized;

    this.invalidateSelection(previousRange, normalized);
  }

  getSelectionRange(): CellRange | null {
    return this.selectionRange;
  }

  getSelection(): Selection | null {
    return this.selection ? { ...this.selection } : null;
  }

  setRangeSelection(range: CellRange | null): void {
    const previousRange = this.rangeSelection;
    const normalized = range ? this.normalizeSelectionRange(range) : null;
    if (isSameCellRange(previousRange, normalized)) return;
    this.rangeSelection = normalized;
    this.invalidateSelection(previousRange, normalized);
  }

  setRemotePresences(presences: GridPresence[] | null): void {
    if (presences === this.remotePresences) return;
    this.remotePresences = presences ?? [];

    const viewport = this.scroll.getViewportState();
    this.dirty.selection.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    this.requestRender();
  }

  pickCellAt(viewportX: number, viewportY: number): Selection | null {
    const viewport = this.scroll.getViewportState();
    const { frozenWidth, frozenHeight, frozenRows, frozenCols } = viewport;

    const colAxis = this.scroll.cols;
    const rowAxis = this.scroll.rows;

    const absScrollX = frozenWidth + viewport.scrollX;
    const absScrollY = frozenHeight + viewport.scrollY;

    let sheetX: number;
    let sheetY: number;
    let minRow = 0;
    let maxRowInclusive = this.getRowCount() - 1;
    let minCol = 0;
    let maxColInclusive = this.getColCount() - 1;

    if (viewportX < frozenWidth && viewportY < frozenHeight) {
      sheetX = viewportX;
      sheetY = viewportY;
      maxRowInclusive = frozenRows - 1;
      maxColInclusive = frozenCols - 1;
    } else if (viewportY < frozenHeight) {
      sheetX = absScrollX + (viewportX - frozenWidth);
      sheetY = viewportY;
      maxRowInclusive = frozenRows - 1;
      minCol = frozenCols;
    } else if (viewportX < frozenWidth) {
      sheetX = viewportX;
      sheetY = absScrollY + (viewportY - frozenHeight);
      minRow = frozenRows;
      maxColInclusive = frozenCols - 1;
    } else {
      sheetX = absScrollX + (viewportX - frozenWidth);
      sheetY = absScrollY + (viewportY - frozenHeight);
      minRow = frozenRows;
      minCol = frozenCols;
    }

    if (maxRowInclusive < minRow || maxColInclusive < minCol) return null;

    const row = rowAxis.indexAt(sheetY, { min: minRow, maxInclusive: maxRowInclusive });
    const col = colAxis.indexAt(sheetX, { min: minCol, maxInclusive: maxColInclusive });

    return { row, col };
  }

  renderImmediately(): void {
    this.renderFrame();
  }

  requestRender(): void {
    if (this.scheduled) return;
    this.scheduled = true;
    requestAnimationFrame(() => {
      this.scheduled = false;
      this.renderFrame();
    });
  }

  markAllDirty(): void {
    const viewport = this.scroll.getViewportState();
    const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    this.dirty.background.markDirty(full);
    this.dirty.content.markDirty(full);
    this.dirty.selection.markDirty(full);
    this.forceFullRedraw = true;
    this.prefetchVisibleRange(viewport, { force: true });
    this.requestRender();
  }

  private renderFrame(): void {
    const viewport = this.scroll.getViewportState();

    const scrollDeltaX = this.lastRendered.scrollX - viewport.scrollX;
    const scrollDeltaY = this.lastRendered.scrollY - viewport.scrollY;
    const viewportChanged =
      viewport.width !== this.lastRendered.width ||
      viewport.height !== this.lastRendered.height ||
      viewport.frozenRows !== this.lastRendered.frozenRows ||
      viewport.frozenCols !== this.lastRendered.frozenCols ||
      viewport.frozenWidth !== this.lastRendered.frozenWidth ||
      viewport.frozenHeight !== this.lastRendered.frozenHeight ||
      this.devicePixelRatio !== this.lastRendered.devicePixelRatio;

    if (viewportChanged) {
      this.forceFullRedraw = true;
    }

    if (scrollDeltaX !== 0 || scrollDeltaY !== 0) {
      if (!this.forceFullRedraw && this.canBlitScroll(viewport, scrollDeltaX, scrollDeltaY)) {
        this.blitScroll(viewport, scrollDeltaX, scrollDeltaY);
        this.markScrollDirtyRegions(viewport, scrollDeltaX, scrollDeltaY);
      } else {
        this.markFullViewportDirty(viewport);
      }

      // Selection overlay moves relative to scroll.
      this.dirty.selection.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    }

    this.renderLayer("background", viewport, this.dirty.background.drain());
    this.renderLayer("content", viewport, this.dirty.content.drain());
    this.renderLayer("selection", viewport, this.dirty.selection.drain());

    this.lastRendered = {
      width: viewport.width,
      height: viewport.height,
      frozenRows: viewport.frozenRows,
      frozenCols: viewport.frozenCols,
      frozenWidth: viewport.frozenWidth,
      frozenHeight: viewport.frozenHeight,
      scrollX: viewport.scrollX,
      scrollY: viewport.scrollY,
      devicePixelRatio: this.devicePixelRatio
    };
    this.forceFullRedraw = false;
  }

  private invalidateForScroll(): void {
    const viewport = this.scroll.getViewportState();
    this.prefetchVisibleRange(viewport);
    this.requestRender();
  }

  private prefetchVisibleRange(viewport: GridViewportState, options?: { force?: boolean }): void {
    if (!this.provider.prefetch) return;
    if (viewport.width <= 0 || viewport.height <= 0) return;

    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    const frozenHeight = Math.min(viewport.height, viewport.frozenHeight);
    const frozenWidth = Math.min(viewport.width, viewport.frozenWidth);

    const frozenRowsRange =
      viewport.frozenRows === 0 || frozenHeight === 0
        ? { start: 0, end: 0 }
        : this.scroll.rows.visibleRange(0, frozenHeight, { min: 0, maxExclusive: viewport.frozenRows });
    const frozenColsRange =
      viewport.frozenCols === 0 || frozenWidth === 0
        ? { start: 0, end: 0 }
        : this.scroll.cols.visibleRange(0, frozenWidth, { min: 0, maxExclusive: viewport.frozenCols });

    const mainRows = viewport.main.rows;
    const mainCols = viewport.main.cols;

    const nextRanges: CellRange[] = [];
    const pushRange = (range: CellRange) => {
      if (range.endRow <= range.startRow) return;
      if (range.endCol <= range.startCol) return;
      nextRanges.push(range);
    };

    // Frozen (top-left) quadrant.
    pushRange({
      startRow: frozenRowsRange.start,
      endRow: frozenRowsRange.end,
      startCol: frozenColsRange.start,
      endCol: frozenColsRange.end
    });

    // Frozen rows + scrollable columns (top-right) quadrant.
    pushRange({
      startRow: frozenRowsRange.start,
      endRow: frozenRowsRange.end,
      startCol: Math.max(viewport.frozenCols, mainCols.start - this.prefetchOverscanCols),
      endCol: Math.min(colCount, mainCols.end + this.prefetchOverscanCols)
    });

    // Scrollable rows + frozen columns (bottom-left) quadrant.
    pushRange({
      startRow: Math.max(viewport.frozenRows, mainRows.start - this.prefetchOverscanRows),
      endRow: Math.min(rowCount, mainRows.end + this.prefetchOverscanRows),
      startCol: frozenColsRange.start,
      endCol: frozenColsRange.end
    });

    // Scrollable (main) quadrant.
    pushRange({
      startRow: Math.max(viewport.frozenRows, mainRows.start - this.prefetchOverscanRows),
      endRow: Math.min(rowCount, mainRows.end + this.prefetchOverscanRows),
      startCol: Math.max(viewport.frozenCols, mainCols.start - this.prefetchOverscanCols),
      endCol: Math.min(colCount, mainCols.end + this.prefetchOverscanCols)
    });

    if (!options?.force && this.lastPrefetchRanges && CanvasGridRenderer.rangesListEqual(this.lastPrefetchRanges, nextRanges)) {
      return;
    }

    const prevRanges = this.lastPrefetchRanges;
    this.lastPrefetchRanges = nextRanges;

    if (!options?.force && prevRanges && prevRanges.length === nextRanges.length) {
      for (let i = 0; i < nextRanges.length; i++) {
        const next = nextRanges[i];
        const prev = prevRanges[i];
        if (!prev || !CanvasGridRenderer.rangesEqual(prev, next)) {
          this.provider.prefetch(next);
        }
      }
      return;
    }

    for (const range of nextRanges) {
      this.provider.prefetch(range);
    }
  }

  private static sanitizeOverscan(value: number | undefined): number {
    if (typeof value !== "number" || !Number.isFinite(value)) return 0;
    return Math.max(0, Math.floor(value));
  }

  private static rangesEqual(a: CellRange, b: CellRange): boolean {
    return (
      a.startRow === b.startRow &&
      a.endRow === b.endRow &&
      a.startCol === b.startCol &&
      a.endCol === b.endCol
    );
  }

  private static rangesListEqual(a: CellRange[], b: CellRange[]): boolean {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
      if (!CanvasGridRenderer.rangesEqual(a[i], b[i])) return false;
    }
    return true;
  }

  private alignScrollToDevicePixels(pos: { x: number; y: number }): { x: number; y: number } {
    const dpr = Number.isFinite(this.devicePixelRatio) && this.devicePixelRatio > 0 ? this.devicePixelRatio : 1;
    const step = 1 / dpr;
    const { maxScrollX, maxScrollY } = this.scroll.getMaxScroll();

    const maxAlignedX = Math.floor(maxScrollX / step) * step;
    const maxAlignedY = Math.floor(maxScrollY / step) * step;

    const x = Math.min(maxAlignedX, Math.max(0, Math.round(pos.x / step) * step));
    const y = Math.min(maxAlignedY, Math.max(0, Math.round(pos.y / step) * step));
    return { x, y };
  }

  private markFullViewportDirty(viewport: GridViewportState): void {
    const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    this.dirty.background.markDirty(full);
    this.dirty.content.markDirty(full);
  }

  private ensureBlitCanvas(): void {
    const gridCanvas = this.gridCanvas;
    if (!gridCanvas) return;

    if (!this.blitCanvas) {
      this.blitCanvas = document.createElement("canvas");
      this.blitCtx = this.blitCanvas.getContext("2d") ?? undefined;
    }

    if (!this.blitCanvas || !this.blitCtx) return;

    if (this.blitCanvas.width !== gridCanvas.width) this.blitCanvas.width = gridCanvas.width;
    if (this.blitCanvas.height !== gridCanvas.height) this.blitCanvas.height = gridCanvas.height;
  }

  private canBlitScroll(viewport: GridViewportState, deltaX: number, deltaY: number): boolean {
    if (!this.gridCanvas || !this.contentCanvas) return false;
    if (!this.gridCtx || !this.contentCtx) return false;
    if (!this.blitCanvas || !this.blitCtx) return false;
    if (viewport.width === 0 || viewport.height === 0) return false;

    const scrollableWidth = Math.max(0, viewport.width - viewport.frozenWidth);
    const scrollableHeight = Math.max(0, viewport.height - viewport.frozenHeight);

    if (deltaX !== 0 && (scrollableWidth === 0 || Math.abs(deltaX) >= scrollableWidth)) return false;
    if (deltaY !== 0 && (scrollableHeight === 0 || Math.abs(deltaY) >= scrollableHeight)) return false;

    const dpr = Number.isFinite(this.devicePixelRatio) && this.devicePixelRatio > 0 ? this.devicePixelRatio : 1;
    const dxDevice = deltaX * dpr;
    const dyDevice = deltaY * dpr;
    const epsilon = 1e-6;
    if (deltaX !== 0 && Math.abs(dxDevice - Math.round(dxDevice)) > epsilon) return false;
    if (deltaY !== 0 && Math.abs(dyDevice - Math.round(dyDevice)) > epsilon) return false;

    return true;
  }

  private blitScroll(viewport: GridViewportState, deltaX: number, deltaY: number): void {
    this.ensureBlitCanvas();
    if (!this.blitCanvas || !this.blitCtx) return;

    this.blitLayer("background", viewport, deltaX, deltaY);
    this.blitLayer("content", viewport, deltaX, deltaY);
  }

  private blitLayer(layer: "background" | "content", viewport: GridViewportState, deltaX: number, deltaY: number): void {
    const canvas = layer === "background" ? this.gridCanvas : this.contentCanvas;
    const ctx = layer === "background" ? this.gridCtx : this.contentCtx;
    if (!canvas || !ctx) return;
    if (!this.blitCanvas || !this.blitCtx) return;

    const dpr = Number.isFinite(this.devicePixelRatio) && this.devicePixelRatio > 0 ? this.devicePixelRatio : 1;
    const dx = Math.round(deltaX * dpr);
    const dy = Math.round(deltaY * dpr);

    const widthPx = canvas.width;
    const heightPx = canvas.height;

    // Copy current layer into the blit buffer.
    this.blitCtx.setTransform(1, 0, 0, 1, 0, 0);
    this.blitCtx.clearRect(0, 0, widthPx, heightPx);
    this.blitCtx.drawImage(canvas, 0, 0);

    const frozenWidthPx = Math.round(viewport.frozenWidth * dpr);
    const frozenHeightPx = Math.round(viewport.frozenHeight * dpr);

    const quadrants = [
      {
        rect: { x: frozenWidthPx, y: 0, width: widthPx - frozenWidthPx, height: frozenHeightPx },
        shiftX: dx,
        shiftY: 0
      },
      {
        rect: { x: 0, y: frozenHeightPx, width: frozenWidthPx, height: heightPx - frozenHeightPx },
        shiftX: 0,
        shiftY: dy
      },
      {
        rect: {
          x: frozenWidthPx,
          y: frozenHeightPx,
          width: widthPx - frozenWidthPx,
          height: heightPx - frozenHeightPx
        },
        shiftX: dx,
        shiftY: dy
      }
    ];

    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);

    for (const quadrant of quadrants) {
      const { rect, shiftX, shiftY } = quadrant;
      if (rect.width <= 0 || rect.height <= 0) continue;
      if (shiftX === 0 && shiftY === 0) continue;

      if (layer === "background") {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      } else {
        ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(rect.x, rect.y, rect.width, rect.height);
      ctx.clip();
      ctx.drawImage(this.blitCanvas, shiftX, shiftY);
      ctx.restore();
    }

    ctx.restore();
  }

  private markScrollDirtyRegions(viewport: GridViewportState, deltaX: number, deltaY: number): void {
    const frozenWidth = viewport.frozenWidth;
    const frozenHeight = viewport.frozenHeight;

    const topRight = { x: frozenWidth, y: 0, width: viewport.width - frozenWidth, height: frozenHeight };
    const bottomLeft = { x: 0, y: frozenHeight, width: frozenWidth, height: viewport.height - frozenHeight };
    const main = {
      x: frozenWidth,
      y: frozenHeight,
      width: viewport.width - frozenWidth,
      height: viewport.height - frozenHeight
    };

    const candidates: { rect: Rect; shiftX: number; shiftY: number }[] = [
      { rect: topRight, shiftX: deltaX, shiftY: 0 },
      { rect: bottomLeft, shiftX: 0, shiftY: deltaY },
      { rect: main, shiftX: deltaX, shiftY: deltaY }
    ];

    for (const { rect, shiftX, shiftY } of candidates) {
      if (rect.width <= 0 || rect.height <= 0) continue;

      if (shiftX > 0) {
        this.markDirtyBoth({ x: rect.x, y: rect.y, width: shiftX, height: rect.height });
      } else if (shiftX < 0) {
        this.markDirtyBoth({
          x: rect.x + rect.width + shiftX,
          y: rect.y,
          width: -shiftX,
          height: rect.height
        });
      }

      if (shiftY > 0) {
        this.markDirtyBoth({ x: rect.x, y: rect.y, width: rect.width, height: shiftY });
      } else if (shiftY < 0) {
        this.markDirtyBoth({
          x: rect.x,
          y: rect.y + rect.height + shiftY,
          width: rect.width,
          height: -shiftY
        });
      }
    }
  }

  private markDirtyBoth(rect: Rect): void {
    const padded: Rect = { x: rect.x - 1, y: rect.y - 1, width: rect.width + 2, height: rect.height + 2 };
    this.dirty.background.markDirty(padded);
    this.dirty.content.markDirty(padded);
  }

  private onProviderUpdate(update: CellProviderUpdate): void {
    if (update.type === "invalidateAll") {
      this.markAllDirty();
      return;
    }

    const viewport = this.scroll.getViewportState();
    const rects = this.rangeToViewportRects(update.range, viewport);
    for (const rect of rects) {
      this.dirty.background.markDirty(rect);
      this.dirty.content.markDirty(rect);
    }
    this.requestRender();
  }

  private rangeToViewportRects(range: CellRange, viewport: GridViewportState): Rect[] {
    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);

    const frozenRows = viewport.frozenRows;
    const frozenCols = viewport.frozenCols;

    const rowsFrozenStart = range.startRow;
    const rowsFrozenEnd = Math.min(range.endRow, frozenRows);
    const rowsScrollStart = Math.max(range.startRow, frozenRows);
    const rowsScrollEnd = range.endRow;

    const colsFrozenStart = range.startCol;
    const colsFrozenEnd = Math.min(range.endCol, frozenCols);
    const colsScrollStart = Math.max(range.startCol, frozenCols);
    const colsScrollEnd = range.endCol;

    const rects: Rect[] = [];

    const addRect = (rowStart: number, rowEnd: number, colStart: number, colEnd: number, scrollRows: boolean, scrollCols: boolean) => {
      if (rowStart >= rowEnd || colStart >= colEnd) return;

      const x1 = colAxis.positionOf(colStart);
      const x2 = colAxis.positionOf(colEnd);
      const y1 = rowAxis.positionOf(rowStart);
      const y2 = rowAxis.positionOf(rowEnd);

      const x = scrollCols ? x1 - viewport.scrollX : x1;
      const y = scrollRows ? y1 - viewport.scrollY : y1;

      const quadrantRect: Rect = {
        x: scrollCols ? frozenWidth : 0,
        y: scrollRows ? frozenHeight : 0,
        width: scrollCols ? Math.max(0, viewport.width - frozenWidth) : frozenWidth,
        height: scrollRows ? Math.max(0, viewport.height - frozenHeight) : frozenHeight
      };

      const rect = intersectRect({ x, y, width: x2 - x1, height: y2 - y1 }, quadrantRect);
      if (rect) rects.push(rect);
    };

    addRect(rowsFrozenStart, rowsFrozenEnd, colsFrozenStart, colsFrozenEnd, false, false);
    addRect(rowsFrozenStart, rowsFrozenEnd, colsScrollStart, colsScrollEnd, false, true);
    addRect(rowsScrollStart, rowsScrollEnd, colsFrozenStart, colsFrozenEnd, true, false);
    addRect(rowsScrollStart, rowsScrollEnd, colsScrollStart, colsScrollEnd, true, true);

    return rects;
  }

  private renderLayer(layer: Layer, viewport: GridViewportState, regions: Rect[]): void {
    const ctx =
      layer === "background"
        ? this.gridCtx
        : layer === "content"
          ? this.contentCtx
          : this.selectionCtx;

    if (!ctx) return;
    if (viewport.width === 0 || viewport.height === 0) return;
    if (regions.length === 0) return;

    const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    const shouldFullRender =
      regions.length > 8 ||
      regions.some((region) => region.x <= 0 && region.y <= 0 && region.width >= viewport.width && region.height >= viewport.height);

    const toRender = shouldFullRender ? [full] : regions;

    for (const region of toRender) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(region.x, region.y, region.width, region.height);
      ctx.clip();

      if (layer === "background") {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(region.x, region.y, region.width, region.height);
      } else {
        ctx.clearRect(region.x, region.y, region.width, region.height);
      }

      this.renderQuadrants(layer, viewport, region);

      if (layer === "selection") {
        this.renderRemotePresenceOverlays(ctx, viewport);
      }

      ctx.restore();
    }

    if (layer === "selection") {
      this.drawFreezeLines(ctx, viewport);
    }
  }

  private renderQuadrants(layer: Layer, viewport: GridViewportState, region: Rect): void {
    const { frozenCols, frozenRows, frozenWidth, frozenHeight, width, height } = viewport;
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    const absScrollX = frozenWidth + viewport.scrollX;
    const absScrollY = frozenHeight + viewport.scrollY;

    const quadrants = [
      {
        originX: 0,
        originY: 0,
        rect: { x: 0, y: 0, width: frozenWidth, height: frozenHeight },
        minRow: 0,
        maxRowExclusive: frozenRows,
        minCol: 0,
        maxColExclusive: frozenCols,
        scrollBaseX: 0,
        scrollBaseY: 0
      },
      {
        originX: frozenWidth,
        originY: 0,
        rect: { x: frozenWidth, y: 0, width: width - frozenWidth, height: frozenHeight },
        minRow: 0,
        maxRowExclusive: frozenRows,
        minCol: frozenCols,
        maxColExclusive: colCount,
        scrollBaseX: absScrollX,
        scrollBaseY: 0
      },
      {
        originX: 0,
        originY: frozenHeight,
        rect: { x: 0, y: frozenHeight, width: frozenWidth, height: height - frozenHeight },
        minRow: frozenRows,
        maxRowExclusive: rowCount,
        minCol: 0,
        maxColExclusive: frozenCols,
        scrollBaseX: 0,
        scrollBaseY: absScrollY
      },
      {
        originX: frozenWidth,
        originY: frozenHeight,
        rect: { x: frozenWidth, y: frozenHeight, width: width - frozenWidth, height: height - frozenHeight },
        minRow: frozenRows,
        maxRowExclusive: rowCount,
        minCol: frozenCols,
        maxColExclusive: colCount,
        scrollBaseX: absScrollX,
        scrollBaseY: absScrollY
      }
    ];

    for (const quadrant of quadrants) {
      if (quadrant.rect.width <= 0 || quadrant.rect.height <= 0) continue;
      if (quadrant.maxRowExclusive <= quadrant.minRow || quadrant.maxColExclusive <= quadrant.minCol) continue;

      const intersection = intersectRect(region, quadrant.rect);
      if (!intersection) continue;

      const sheetX = quadrant.scrollBaseX + (intersection.x - quadrant.originX);
      const sheetY = quadrant.scrollBaseY + (intersection.y - quadrant.originY);
      const sheetXEnd = sheetX + intersection.width;
      const sheetYEnd = sheetY + intersection.height;

      const startRow = this.scroll.rows.indexAt(sheetY, {
        min: quadrant.minRow,
        maxInclusive: quadrant.maxRowExclusive - 1
      });
      const endRow = Math.min(
        this.scroll.rows.indexAt(sheetYEnd, {
          min: quadrant.minRow,
          maxInclusive: quadrant.maxRowExclusive - 1
        }) + 1,
        quadrant.maxRowExclusive
      );

      const startCol = this.scroll.cols.indexAt(sheetX, {
        min: quadrant.minCol,
        maxInclusive: quadrant.maxColExclusive - 1
      });
      const endCol = Math.min(
        this.scroll.cols.indexAt(sheetXEnd, {
          min: quadrant.minCol,
          maxInclusive: quadrant.maxColExclusive - 1
        }) + 1,
        quadrant.maxColExclusive
      );

      if (endRow <= startRow || endCol <= startCol) continue;

      if (layer === "selection") {
        this.renderSelectionQuadrant(intersection, viewport);
        continue;
      }

      if (layer === "background") {
        this.renderBackgroundQuadrant(intersection, quadrant, startRow, endRow, startCol, endCol);
      } else {
        this.renderContentQuadrant(intersection, quadrant, startRow, endRow, startCol, endCol);
      }
    }
  }

  private renderBackgroundQuadrant(
    intersection: Rect,
    quadrant: {
      originX: number;
      originY: number;
      scrollBaseX: number;
      scrollBaseY: number;
    },
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number
  ): void {
    if (!this.gridCtx) return;
    const ctx = this.gridCtx;

    const fills = new Map<string, Rect[]>();

    let rowYSheet = this.scroll.rows.positionOf(startRow);
    for (let row = startRow; row < endRow; row++) {
      const rowHeight = this.scroll.rows.getSize(row);
      const y = rowYSheet - quadrant.scrollBaseY + quadrant.originY;

      let colXSheet = this.scroll.cols.positionOf(startCol);
      for (let col = startCol; col < endCol; col++) {
        const colWidth = this.scroll.cols.getSize(col);
        const x = colXSheet - quadrant.scrollBaseX + quadrant.originX;

        const cell = this.provider.getCell(row, col);
        const fill = cell?.style?.fill;
        if (fill) {
          const bucket = fills.get(fill) ?? [];
          bucket.push({ x, y, width: colWidth, height: rowHeight });
          fills.set(fill, bucket);
        }

        colXSheet += colWidth;
      }

      rowYSheet += rowHeight;
    }

    for (const [fill, rects] of fills) {
      ctx.fillStyle = fill;
      ctx.beginPath();
      for (const rect of rects) {
        const clipped = intersectRect(rect, intersection);
        if (!clipped) continue;
        ctx.rect(clipped.x, clipped.y, clipped.width, clipped.height);
      }
      ctx.fill();
    }

    ctx.strokeStyle = "#e6e6e6";
    ctx.lineWidth = 1;

    ctx.beginPath();
    let xSheet = this.scroll.cols.positionOf(startCol);
    for (let col = startCol; col <= endCol; col++) {
      const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
      if (x >= intersection.x - 1 && x <= intersection.x + intersection.width + 1) {
        const cx = crispLine(x);
        ctx.moveTo(cx, intersection.y);
        ctx.lineTo(cx, intersection.y + intersection.height);
      }
      if (col < endCol) xSheet += this.scroll.cols.getSize(col);
    }

    let ySheet = this.scroll.rows.positionOf(startRow);
    for (let row = startRow; row <= endRow; row++) {
      const y = ySheet - quadrant.scrollBaseY + quadrant.originY;
      if (y >= intersection.y - 1 && y <= intersection.y + intersection.height + 1) {
        const cy = crispLine(y);
        ctx.moveTo(intersection.x, cy);
        ctx.lineTo(intersection.x + intersection.width, cy);
      }
      if (row < endRow) ySheet += this.scroll.rows.getSize(row);
    }

    ctx.stroke();
  }

  private renderContentQuadrant(
    _intersection: Rect,
    quadrant: {
      originX: number;
      originY: number;
      scrollBaseX: number;
      scrollBaseY: number;
    },
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number
  ): void {
    if (!this.contentCtx) return;
    const ctx = this.contentCtx;

    const layoutEngine = this.textLayoutEngine;
    ctx.textBaseline = "alphabetic";
    ctx.textAlign = "left";

    let currentFont = "";
    let currentFillStyle = "";
    const paddingX = 4;
    const paddingY = 2;

    let rowYSheet = this.scroll.rows.positionOf(startRow);
    for (let row = startRow; row < endRow; row++) {
      const rowHeight = this.scroll.rows.getSize(row);
      const y = rowYSheet - quadrant.scrollBaseY + quadrant.originY;

      let colXSheet = this.scroll.cols.positionOf(startCol);
      for (let col = startCol; col < endCol; col++) {
        const colWidth = this.scroll.cols.getSize(col);
        const x = colXSheet - quadrant.scrollBaseX + quadrant.originX;

        const cell = this.provider.getCell(row, col);
        if (cell && cell.value !== null) {
          const style = cell.style;
          const fontSize = style?.fontSize ?? 12;
          const fontFamily = style?.fontFamily ?? "system-ui";
          const fontWeight = style?.fontWeight ?? "400";
          const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight };
          const font = toCanvasFontString(fontSpec);

          if (font !== currentFont) {
            ctx.font = font;
            currentFont = font;
          }

          const fillStyle = resolveCellTextColor(cell.value, style?.color);
          if (fillStyle !== currentFillStyle) {
            ctx.fillStyle = fillStyle;
            currentFillStyle = fillStyle;
          }

          const formattedKey = formatCellDisplayText(cell.value);
          let text = this.formattedCache.get(formattedKey);
          if (text === undefined) {
            text = formattedKey;
            this.formattedCache.set(formattedKey, text);
          }

          const wrapMode = style?.wrapMode ?? "none";
          const direction = style?.direction ?? "auto";
          const verticalAlign = style?.verticalAlign ?? "middle";
          const rotationDeg = style?.rotationDeg ?? 0;

          const availableWidth = Math.max(0, colWidth - paddingX * 2);
          const availableHeight = Math.max(0, rowHeight - paddingY * 2);
          const lineHeight = Math.ceil(fontSize * 1.2);
          const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

          const align: CanvasTextAlign =
            style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
          const layoutAlign =
            align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
              ? (align as "left" | "right" | "center" | "start" | "end")
              : "start";

          const hasExplicitNewline = /[\r\n]/.test(text);
          const rotationRad = (rotationDeg * Math.PI) / 180;

          if (wrapMode === "none" && !hasExplicitNewline && rotationDeg === 0) {
            // Fast path for the common case: single-line text with clipping, using cached metrics.
            const baseDirection = direction === "auto" ? detectBaseDirection(text) : direction;
            const resolvedAlign =
              layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
                ? layoutAlign
                : resolveAlign(layoutAlign, baseDirection);

            const measurement = layoutEngine?.measure(text, fontSpec);
            const textWidth = measurement?.width ?? ctx.measureText(text).width;
            const ascent = measurement?.ascent ?? fontSize * 0.8;
            const descent = measurement?.descent ?? fontSize * 0.2;

            let textX = x + paddingX;
            if (resolvedAlign === "center") {
              textX = x + paddingX + (availableWidth - textWidth) / 2;
            } else if (resolvedAlign === "right") {
              textX = x + paddingX + (availableWidth - textWidth);
            }

            let baselineY = y + paddingY + ascent;
            if (verticalAlign === "middle") {
              baselineY = y + rowHeight / 2 + (ascent - descent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + rowHeight - paddingY - descent;
            }

            const shouldClip = textWidth > availableWidth;
            if (shouldClip) {
              ctx.save();
              ctx.beginPath();
              ctx.rect(x, y, colWidth, rowHeight);
              ctx.clip();
              ctx.fillText(text, textX, baselineY);
              ctx.restore();
            } else {
              ctx.fillText(text, textX, baselineY);
            }
          } else if (layoutEngine && availableWidth > 0) {
            const layout = layoutEngine.layout({
              text,
              font: fontSpec,
              maxWidth: availableWidth,
              wrapMode,
              align: layoutAlign,
              direction,
              lineHeightPx: lineHeight,
              maxLines
            });

            let originY = y + paddingY;
            if (verticalAlign === "middle") {
              originY = y + paddingY + Math.max(0, (availableHeight - layout.height) / 2);
            } else if (verticalAlign === "bottom") {
              originY = y + rowHeight - paddingY - layout.height;
            }

            const originX = x + paddingX;
            const shouldClip = layout.width > availableWidth || layout.height > availableHeight || rotationDeg !== 0;

            if (shouldClip) {
              ctx.save();
              ctx.beginPath();
              ctx.rect(x, y, colWidth, rowHeight);
              ctx.clip();

              if (rotationRad) {
                const cx = x + colWidth / 2;
                const cy = y + rowHeight / 2;
                ctx.translate(cx, cy);
                ctx.rotate(rotationRad);
                ctx.translate(-cx, -cy);
              }

              drawTextLayout(ctx, layout, originX, originY);
              ctx.restore();
            } else {
              drawTextLayout(ctx, layout, originX, originY);
            }
          } else {
            // Fallback: no layout engine available (shouldn't happen in supported environments).
            ctx.save();
            ctx.beginPath();
            ctx.rect(x, y, colWidth, rowHeight);
            ctx.clip();
            ctx.textBaseline = "middle";
            ctx.fillText(text, x + paddingX, y + rowHeight / 2);
            ctx.textBaseline = "alphabetic";
            ctx.restore();
          }
        }

        if (cell?.comment) {
          const resolved = cell.comment.resolved ?? false;
          const maxSize = Math.min(colWidth, rowHeight);
          const size = Math.min(maxSize, Math.max(6, maxSize * 0.25));
          if (size > 0) {
            ctx.save();
            ctx.beginPath();
            ctx.moveTo(x + colWidth, y);
            ctx.lineTo(x + colWidth - size, y);
            ctx.lineTo(x + colWidth, y + size);
            ctx.closePath();
            ctx.fillStyle = resolved ? "#9ca3af" : "#f59e0b";
            ctx.fill();
            ctx.restore();
          }
        }

        colXSheet += colWidth;
      }

      rowYSheet += rowHeight;
    }
  }

  private renderSelectionQuadrant(
    intersection: Rect,
    viewport: GridViewportState
  ): void {
    const ctx = this.selectionCtx;
    if (!ctx) return;

    const transientRange = this.rangeSelection;
    if (transientRange) {
      const rects = this.rangeToViewportRects(transientRange, viewport);

      ctx.fillStyle = "rgba(14, 101, 235, 0.12)";
      for (const rect of rects) {
        const clipped = intersectRect(rect, intersection);
        if (!clipped) continue;
        ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
      }

      ctx.strokeStyle = "#0e65eb";
      ctx.lineWidth = 2;
      for (const rect of rects) {
        if (!intersectRect(rect, intersection)) continue;
        if (rect.width <= 2 || rect.height <= 2) continue;
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
      }
    }

    const range = this.selectionRange;
    if (!range) return;

    const rects = this.rangeToViewportRects(range, viewport);
    if (rects.length === 0) return;

    ctx.fillStyle = "rgba(14, 101, 235, 0.12)";
    for (const rect of rects) {
      const clipped = intersectRect(rect, intersection);
      if (!clipped) continue;
      ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
    }

    ctx.strokeStyle = "#0e65eb";
    ctx.lineWidth = 2;
    for (const rect of rects) {
      if (!intersectRect(rect, intersection)) continue;
      if (rect.width <= 2 || rect.height <= 2) continue;
      ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
    }

    const handleSize = 8;
    const handleRow = range.endRow - 1;
    const handleCol = range.endCol - 1;
    const handleCellRect = this.cellRectInViewport(handleRow, handleCol, viewport);
    if (!handleCellRect) return;
    if (handleCellRect.width < handleSize || handleCellRect.height < handleSize) return;

    const handleRect: Rect = {
      x: handleCellRect.x + handleCellRect.width - handleSize / 2,
      y: handleCellRect.y + handleCellRect.height - handleSize / 2,
      width: handleSize,
      height: handleSize
    };
    const handleClipped = intersectRect(handleRect, intersection);
    if (!handleClipped) return;

    ctx.fillStyle = "#0e65eb";
    ctx.fillRect(handleClipped.x, handleClipped.y, handleClipped.width, handleClipped.height);
  }

  private renderRemotePresenceOverlays(ctx: CanvasRenderingContext2D, viewport: GridViewportState): void {
    if (this.remotePresences.length === 0) return;

    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    ctx.save();
    ctx.font = this.presenceFont;
    ctx.textBaseline = "top";

    const selectionFillAlpha = 0.12;
    const selectionStrokeAlpha = 0.9;
    const cursorStrokeWidth = 2;
    const badgePaddingX = 6;
    const badgePaddingY = 3;
    const badgeOffsetX = 8;
    const badgeOffsetY = -18;
    const badgeTextHeight = 14;

    for (const presence of this.remotePresences) {
      const color = presence.color ?? "#4c8bf5";

      if (presence.selections.length > 0) {
        ctx.fillStyle = color;
        ctx.strokeStyle = color;
        ctx.lineWidth = cursorStrokeWidth;

        for (const range of presence.selections) {
          const startRow = Math.min(range.startRow, range.endRow);
          const endRow = Math.max(range.startRow, range.endRow);
          const startCol = Math.min(range.startCol, range.endCol);
          const endCol = Math.max(range.startCol, range.endCol);

          if (endRow < 0 || startRow >= rowCount) continue;
          if (endCol < 0 || startCol >= colCount) continue;

          const clampedStartRow = Math.max(0, startRow);
          const clampedEndRow = Math.min(rowCount - 1, endRow);
          const clampedStartCol = Math.max(0, startCol);
          const clampedEndCol = Math.min(colCount - 1, endCol);

          const rects = this.rangeToViewportRects(
            {
              startRow: clampedStartRow,
              endRow: Math.min(rowCount, clampedEndRow + 1),
              startCol: clampedStartCol,
              endCol: Math.min(colCount, clampedEndCol + 1)
            },
            viewport
          );

          if (rects.length === 0) continue;

          ctx.globalAlpha = selectionFillAlpha;
          for (const rect of rects) {
            ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
          }

          ctx.globalAlpha = selectionStrokeAlpha;
          for (const rect of rects) {
            ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
          }

          ctx.globalAlpha = 1;
        }
      }

      if (presence.cursor) {
        const cursorRect = this.cellRectInViewport(presence.cursor.row, presence.cursor.col, viewport);
        if (!cursorRect) continue;

        ctx.globalAlpha = 1;
        ctx.strokeStyle = color;
        ctx.lineWidth = cursorStrokeWidth;
        ctx.strokeRect(cursorRect.x + 1, cursorRect.y + 1, cursorRect.width - 2, cursorRect.height - 2);

        const name = presence.name ?? "Anonymous";
        const metricsKey = `${this.presenceFont}::${name}`;
        let textWidth = this.textWidthCache.get(metricsKey);
        if (textWidth === undefined) {
          textWidth = ctx.measureText(name).width;
          this.textWidthCache.set(metricsKey, textWidth);
        }

        const badgeWidth = textWidth + badgePaddingX * 2;
        const badgeHeight = badgeTextHeight + badgePaddingY * 2;
        const badgeX = cursorRect.x + cursorRect.width + badgeOffsetX;
        const badgeY = cursorRect.y + badgeOffsetY;

        ctx.fillStyle = color;
        ctx.fillRect(badgeX, badgeY, badgeWidth, badgeHeight);
        ctx.fillStyle = pickTextColor(color);
        ctx.fillText(name, badgeX + badgePaddingX, badgeY + badgePaddingY);
      }
    }

    ctx.restore();
  }

  private drawFreezeLines(ctx: CanvasRenderingContext2D, viewport: GridViewportState): void {
    if (viewport.frozenCols === 0 && viewport.frozenRows === 0) return;

    ctx.strokeStyle = "#c0c0c0";
    ctx.lineWidth = 2;
    ctx.beginPath();

    if (viewport.frozenCols > 0) {
      const x = crispLine(viewport.frozenWidth);
      ctx.moveTo(x, 0);
      ctx.lineTo(x, viewport.height);
    }

    if (viewport.frozenRows > 0) {
      const y = crispLine(viewport.frozenHeight);
      ctx.moveTo(0, y);
      ctx.lineTo(viewport.width, y);
    }

    ctx.stroke();
  }

  private cellRectInViewport(row: number, col: number, viewport: GridViewportState): Rect | null {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const frozenWidth = viewport.frozenWidth;
    const frozenHeight = viewport.frozenHeight;
    const absScrollX = frozenWidth + viewport.scrollX;
    const absScrollY = frozenHeight + viewport.scrollY;

    const colX = colAxis.positionOf(col);
    const rowY = rowAxis.positionOf(row);
    const width = colAxis.getSize(col);
    const height = rowAxis.getSize(row);

    let x: number;
    let y: number;

    if (row < viewport.frozenRows && col < viewport.frozenCols) {
      x = colX;
      y = rowY;
    } else if (row < viewport.frozenRows) {
      x = colX - absScrollX + frozenWidth;
      y = rowY;
    } else if (col < viewport.frozenCols) {
      x = colX;
      y = rowY - absScrollY + frozenHeight;
    } else {
      x = colX - absScrollX + frozenWidth;
      y = rowY - absScrollY + frozenHeight;
    }

    const rect = { x, y, width, height };

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const scrollRows = row >= viewport.frozenRows;
    const scrollCols = col >= viewport.frozenCols;

    const quadrantRect: Rect = {
      x: scrollCols ? frozenWidthClamped : 0,
      y: scrollRows ? frozenHeightClamped : 0,
      width: scrollCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped,
      height: scrollRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped
    };

    return intersectRect(rect, quadrantRect);
  }

  private selectionOverlayRects(range: CellRange | null, viewport: GridViewportState): Rect[] {
    if (!range) return [];

    const rects = this.rangeToViewportRects(range, viewport);

    const handleSize = 8;
    const handleRow = range.endRow - 1;
    const handleCol = range.endCol - 1;

    const handleCellRect = this.cellRectInViewport(handleRow, handleCol, viewport);
    if (handleCellRect && handleCellRect.width >= handleSize && handleCellRect.height >= handleSize) {
      rects.push({
        x: handleCellRect.x + handleCellRect.width - handleSize / 2,
        y: handleCellRect.y + handleCellRect.height - handleSize / 2,
        width: handleSize,
        height: handleSize
      });
    }

    return rects;
  }

  private invalidateSelection(previousRange: CellRange | null, nextRange: CellRange | null): void {
    const viewport = this.scroll.getViewportState();
    const padding = 4;

    const dirtyRects = [
      ...this.selectionOverlayRects(previousRange, viewport),
      ...this.selectionOverlayRects(nextRange, viewport)
    ];

    if (dirtyRects.length === 0) return;

    for (const rect of dirtyRects) {
      this.dirty.selection.markDirty(padRect(rect, padding));
    }

    this.requestRender();
  }

  private normalizeSelectionRange(range: CellRange): CellRange | null {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    let startRow = clampIndex(range.startRow, 0, rowCount);
    let endRow = clampIndex(range.endRow, 0, rowCount);
    let startCol = clampIndex(range.startCol, 0, colCount);
    let endCol = clampIndex(range.endCol, 0, colCount);

    if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
    if (startCol > endCol) [startCol, endCol] = [endCol, startCol];

    if (startRow === endRow || startCol === endCol) return null;

    return { startRow, endRow, startCol, endCol };
  }

  private getRowCount(): number {
    return this.scroll.getCounts().rowCount;
  }

  private getColCount(): number {
    return this.scroll.getCounts().colCount;
  }
}
