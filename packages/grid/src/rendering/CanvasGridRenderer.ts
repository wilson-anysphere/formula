import type { CellProvider, CellProviderUpdate, CellRange } from "../model/CellProvider";
import { DirtyRegionTracker, type Rect } from "./DirtyRegionTracker";
import { setupHiDpiCanvas } from "./HiDpiCanvas";
import { LruCache } from "../utils/LruCache";
import type { GridViewportState } from "../virtualization/VirtualScrollManager";
import { VirtualScrollManager } from "../virtualization/VirtualScrollManager";

type Layer = "background" | "content" | "selection";

interface Selection {
  row: number;
  col: number;
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

export class CanvasGridRenderer {
  private readonly provider: CellProvider;
  readonly scroll: VirtualScrollManager;

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
  private previousSelection: Selection | null = null;

  private readonly formattedCache = new LruCache<string, string>(50_000);
  private readonly textWidthCache = new LruCache<string, number>(50_000);

  constructor(options: {
    provider: CellProvider;
    rowCount: number;
    colCount: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
  }) {
    this.provider = options.provider;
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
    this.previousSelection = this.selection;
    this.selection = selection;

    const viewport = this.scroll.getViewportState();
    const dirtyRect = this.selectionRectDelta(viewport);
    if (dirtyRect) {
      this.dirty.selection.markDirty(dirtyRect);
      this.requestRender();
      return;
    }

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
    this.provider.prefetch?.({
      startRow: viewport.main.rows.start,
      endRow: viewport.main.rows.end,
      startCol: viewport.main.cols.start,
      endCol: viewport.main.cols.end
    });
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
    this.provider.prefetch?.({
      startRow: viewport.main.rows.start,
      endRow: viewport.main.rows.end,
      startCol: viewport.main.cols.start,
      endCol: viewport.main.cols.end
    });
    this.requestRender();
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

    const viewportRect: Rect = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    const rects: Rect[] = [];

    const addRect = (rowStart: number, rowEnd: number, colStart: number, colEnd: number, scrollRows: boolean, scrollCols: boolean) => {
      if (rowStart >= rowEnd || colStart >= colEnd) return;

      const x1 = colAxis.positionOf(colStart);
      const x2 = colAxis.positionOf(colEnd);
      const y1 = rowAxis.positionOf(rowStart);
      const y2 = rowAxis.positionOf(rowEnd);

      const x = scrollCols ? x1 - viewport.scrollX : x1;
      const y = scrollRows ? y1 - viewport.scrollY : y1;

      const rect = intersectRect({ x, y, width: x2 - x1, height: y2 - y1 }, viewportRect);
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

      this.provider.prefetch?.({ startRow, endRow, startCol, endCol });

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

    ctx.textBaseline = "middle";

    let currentFont = "";
    let currentFillStyle = "";
    let currentAlign: CanvasTextAlign = "left";

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
          const font = `${fontWeight} ${fontSize}px ${fontFamily}`;

          if (font !== currentFont) {
            ctx.font = font;
            currentFont = font;
          }

          const fillStyle = style?.color ?? "#111111";
          if (fillStyle !== currentFillStyle) {
            ctx.fillStyle = fillStyle;
            currentFillStyle = fillStyle;
          }

          const formattedKey = `${cell.value}`;
          let text = this.formattedCache.get(formattedKey);
          if (text === undefined) {
            text = formattedKey;
            this.formattedCache.set(formattedKey, text);
          }

          const metricsKey = `${font}::${text}`;
          let textWidth = this.textWidthCache.get(metricsKey);
          if (textWidth === undefined) {
            textWidth = ctx.measureText(text).width;
            this.textWidthCache.set(metricsKey, textWidth);
          }

          const padding = 4;
          const align = style?.textAlign ?? "left";
          if (align !== currentAlign) {
            ctx.textAlign = align;
            currentAlign = align;
          }

          const textX =
            align === "left"
              ? x + padding
              : align === "center"
                ? x + colWidth / 2
                : x + colWidth - padding;
          const textY = y + rowHeight / 2;

          if (textWidth > colWidth - padding * 2) {
            ctx.save();
            ctx.beginPath();
            ctx.rect(x, y, colWidth, rowHeight);
            ctx.clip();
            ctx.fillText(text, textX, textY);
            ctx.restore();
          } else {
            ctx.fillText(text, textX, textY);
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
    if (!this.selection || !this.selectionCtx) return;

    const rect = this.cellRectInViewport(this.selection.row, this.selection.col, viewport);
    if (!rect) return;

    const clipped = intersectRect(rect, intersection);
    if (!clipped) return;

    const ctx = this.selectionCtx;
    ctx.fillStyle = "rgba(14, 101, 235, 0.12)";
    ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);

    ctx.strokeStyle = "#0e65eb";
    ctx.lineWidth = 2;
    ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);

    const handleSize = 8;
    if (rect.width >= handleSize && rect.height >= handleSize) {
      ctx.fillStyle = "#0e65eb";
      ctx.fillRect(
        rect.x + rect.width - handleSize / 2,
        rect.y + rect.height - handleSize / 2,
        handleSize,
        handleSize
      );
    }

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
    const viewportRect = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    return intersectRect(rect, viewportRect);
  }

  private selectionRectDelta(viewport: GridViewportState): Rect | null {
    const next = this.selection ? this.cellRectInViewport(this.selection.row, this.selection.col, viewport) : null;
    const prev = this.previousSelection
      ? this.cellRectInViewport(this.previousSelection.row, this.previousSelection.col, viewport)
      : null;

    if (next && prev) {
      return {
        x: Math.min(next.x, prev.x) - 4,
        y: Math.min(next.y, prev.y) - 4,
        width: Math.max(next.x + next.width, prev.x + prev.width) - Math.min(next.x, prev.x) + 8,
        height: Math.max(next.y + next.height, prev.y + prev.height) - Math.min(next.y, prev.y) + 8
      };
    }

    return next ?? prev;
  }

  private getRowCount(): number {
    return this.scroll.getCounts().rowCount;
  }

  private getColCount(): number {
    return this.scroll.getCounts().colCount;
  }
}
