import type { CellData, CellProvider, CellProviderUpdate, CellRange } from "../model/CellProvider";
import { DirtyRegionTracker, type Rect } from "./DirtyRegionTracker";
import { setupHiDpiCanvas } from "./HiDpiCanvas";
import { LruCache } from "../utils/LruCache";
import type { GridPresence } from "../presence/types";
import type { GridTheme } from "../theme/GridTheme";
import { DEFAULT_GRID_THEME, gridThemesEqual, resolveGridTheme } from "../theme/GridTheme";
import type { GridViewportState } from "../virtualization/VirtualScrollManager";
import { VirtualScrollManager } from "../virtualization/VirtualScrollManager";
import { MergedCellIndex, isInteriorHorizontalGridline, isInteriorVerticalGridline, rangesIntersect } from "./mergedCells";
import {
  TextLayoutEngine,
  createCanvasTextMeasurer,
  detectBaseDirection,
  drawTextLayout,
  resolveAlign,
  toCanvasFontString
} from "@formula/text-layout";

type Layer = "background" | "content" | "selection";

export interface GridPerfStats {
  /** Toggle instrumentation updates. When disabled, stats remain frozen at their last values. */
  enabled: boolean;
  /** Duration of the most recently rendered frame (ms). */
  lastFrameMs: number;
  /**
   * Number of logical cells visited during the last frame's paint passes.
   *
   * This is counted per-cell (not per-layer) in the combined grid paint.
   */
  cellsPainted: number;
  /** Number of `provider.getCell()` calls issued during the last frame. */
  cellFetches: number;
  /** Dirty rectangles drained for each layer at the start of the frame. */
  dirtyRects: {
    background: number;
    content: number;
    selection: number;
    total: number;
  };
  /** Whether the renderer used blitting to reuse previous pixels for scroll. */
  blitUsed: boolean;
}

interface Selection {
  row: number;
  col: number;
}

export type ScrollToCellAlign = "auto" | "start" | "center" | "end";
interface GridRenderQuadrant {
  originX: number;
  originY: number;
  rect: Rect;
  minRow: number;
  maxRowExclusive: number;
  minCol: number;
  maxColExclusive: number;
  scrollBaseX: number;
  scrollBaseY: number;
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

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4.0;

function clampZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) return 1;
  return clamp(zoom, MIN_ZOOM, MAX_ZOOM);
}

function padRect(rect: Rect, padding: number): Rect {
  return { x: rect.x - padding, y: rect.y - padding, width: rect.width + padding * 2, height: rect.height + padding * 2 };
}

function parseCssColor(color: string): { r: number; g: number; b: number } | null {
  const trimmed = color.trim();

  const hex6 = /^#?([0-9a-f]{6})$/i.exec(trimmed);
  if (hex6) {
    const value = Number.parseInt(hex6[1], 16);
    return {
      r: (value >> 16) & 255,
      g: (value >> 8) & 255,
      b: value & 255
    };
  }

  const hex3 = /^#?([0-9a-f]{3})$/i.exec(trimmed);
  if (hex3) {
    const [r, g, b] = hex3[1].split("");
    const value = Number.parseInt(`${r}${r}${g}${g}${b}${b}`, 16);
    return {
      r: (value >> 16) & 255,
      g: (value >> 8) & 255,
      b: value & 255
    };
  }

  const rgb = /^rgba?\(\s*(\d{1,3})(?:\s*,\s*|\s+)(\d{1,3})(?:\s*,\s*|\s+)(\d{1,3})/i.exec(trimmed);
  if (rgb) {
    return {
      r: clamp(Number.parseInt(rgb[1], 10), 0, 255),
      g: clamp(Number.parseInt(rgb[2], 10), 0, 255),
      b: clamp(Number.parseInt(rgb[3], 10), 0, 255)
    };
  }

  return null;
}

function pickTextColor(backgroundColor: string): string {
  const rgb = parseCssColor(backgroundColor);
  if (!rgb) return "#ffffff";
  const luma = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
  return luma > 0.6 ? "#000000" : "#ffffff";
}

type FontSpec = { family: string; sizePx: number; weight?: string | number; style?: string };

type RichTextRunStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: string | boolean;
  color?: string;
  font?: string;
  size_100pt?: number;
};

function pointsToPx(points: number): number {
  // Excel point sizes are typically interpreted at 96DPI.
  return (points * 96) / 72;
}

function engineColorToCanvasColor(color: unknown): string | undefined {
  if (typeof color !== "string") return undefined;
  if (!color.startsWith("#")) return color;
  if (color.length !== 9) return color;

  // Engine colors are serialized as `#AARRGGBB`.
  const hex = color.slice(1);
  const a = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const r = Number.parseInt(hex.slice(2, 4), 16);
  const g = Number.parseInt(hex.slice(4, 6), 16);
  const b = Number.parseInt(hex.slice(6, 8), 16);

  if (Number.isNaN(a) || Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return color;
  }

  if (a >= 1) {
    return `rgb(${r}, ${g}, ${b})`;
  }
  return `rgba(${r}, ${g}, ${b}, ${a})`;
}

function buildCodePointIndex(text: string): number[] {
  const offsets: number[] = [0];
  let utf16Offset = 0;
  for (const ch of text) {
    utf16Offset += ch.length;
    offsets.push(utf16Offset);
  }
  return offsets;
}

function sliceByCodePointRange(text: string, offsets: number[], start: number, end: number): string {
  const len = offsets.length - 1;
  const s = clampIndex(start, 0, len);
  const e = clampIndex(end, s, len);
  return text.slice(offsets[s], offsets[e]);
}

function isUnderline(style: RichTextRunStyle | undefined): boolean {
  const value = style?.underline;
  return Boolean(value && value !== "none");
}

function fontSpecForRichTextStyle(
  style: RichTextRunStyle | undefined,
  defaults: Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number; style: string },
  zoom: number
): FontSpec {
  const family = typeof style?.font === "string" ? style.font : defaults.family;
  const sizePx =
    typeof style?.size_100pt === "number" && Number.isFinite(style.size_100pt)
      ? pointsToPx(style.size_100pt / 100) * zoom
      : defaults.sizePx;
  const weight = style?.bold === true ? "bold" : defaults.weight;
  const fontStyle = style?.italic === true ? "italic" : defaults.style;
  return { family, sizePx, weight, style: fontStyle };
}

const EXPLICIT_NEWLINE_RE = /[\r\n]/;
const MAX_TEXT_OVERFLOW_COLUMNS = 128;
const EMPTY_MERGED_INDEX = new MergedCellIndex([]);

export function formatCellDisplayText(value: CellData["value"]): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function resolveCellTextColor(value: CellData["value"], explicitColor: string | undefined): string {
  return resolveCellTextColorWithTheme(value, explicitColor, undefined);
}

export function resolveCellTextColorWithTheme(
  value: CellData["value"],
  explicitColor: string | undefined,
  theme: Pick<GridTheme, "cellText" | "errorText"> | undefined
): string {
  if (explicitColor !== undefined) return explicitColor;
  if (typeof value === "string" && value.startsWith("#")) return theme?.errorText ?? DEFAULT_GRID_THEME.errorText;
  return theme?.cellText ?? DEFAULT_GRID_THEME.cellText;
}

export class CanvasGridRenderer {
  private readonly provider: CellProvider;
  scroll: VirtualScrollManager;

  // Header rows/cols affect *styling* only (header background + header text color).
  // When unset, the renderer falls back to legacy behavior: treat the first frozen row/col
  // (if any) as the header region.
  private headerRowsOverride: number | null = null;
  private headerColsOverride: number | null = null;

  private readonly prefetchOverscanRows: number;
  private readonly prefetchOverscanCols: number;
  private lastPrefetchRanges: CellRange[] | null = null;

  private zoom = 1;
  private readonly baseDefaultRowHeight: number;
  private readonly baseDefaultColWidth: number;
  private readonly rowHeightOverridesBase = new Map<number, number>();
  private readonly colWidthOverridesBase = new Map<number, number>();

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
  private selectionRanges: CellRange[] = [];
  private activeSelectionIndex = 0;
  private rangeSelection: CellRange | null = null;
  private fillPreviewRange: CellRange | null = null;
  private fillHandleEnabled = true;
  private referenceHighlights: Array<{ range: CellRange; color: string; active: boolean }> = [];

  private remotePresences: GridPresence[] = [];
  private remotePresenceDirtyPadding = 1;

  private readonly textWidthCache = new LruCache<string, number>(10_000);
  private textLayoutEngine?: TextLayoutEngine;

  private presenceFont = "12px system-ui, sans-serif";
  private theme: GridTheme;

  private readonly perfStats: GridPerfStats = {
    enabled: false,
    lastFrameMs: 0,
    cellsPainted: 0,
    cellFetches: 0,
    dirtyRects: { background: 0, content: 0, selection: 0, total: 0 },
    blitUsed: false
  };

  private readonly gridQuadrantScratch: [GridRenderQuadrant, GridRenderQuadrant, GridRenderQuadrant, GridRenderQuadrant] = [
    {
      originX: 0,
      originY: 0,
      rect: { x: 0, y: 0, width: 0, height: 0 },
      minRow: 0,
      maxRowExclusive: 0,
      minCol: 0,
      maxColExclusive: 0,
      scrollBaseX: 0,
      scrollBaseY: 0
    },
    {
      originX: 0,
      originY: 0,
      rect: { x: 0, y: 0, width: 0, height: 0 },
      minRow: 0,
      maxRowExclusive: 0,
      minCol: 0,
      maxColExclusive: 0,
      scrollBaseX: 0,
      scrollBaseY: 0
    },
    {
      originX: 0,
      originY: 0,
      rect: { x: 0, y: 0, width: 0, height: 0 },
      minRow: 0,
      maxRowExclusive: 0,
      minCol: 0,
      maxColExclusive: 0,
      scrollBaseX: 0,
      scrollBaseY: 0
    },
    {
      originX: 0,
      originY: 0,
      rect: { x: 0, y: 0, width: 0, height: 0 },
      minRow: 0,
      maxRowExclusive: 0,
      minCol: 0,
      maxColExclusive: 0,
      scrollBaseX: 0,
      scrollBaseY: 0
    }
  ];
  private mergedIndex: MergedCellIndex = EMPTY_MERGED_INDEX;
  private mergedIndexKey: string | null = null;
  private mergedIndexDirty = true;

  constructor(options: {
    provider: CellProvider;
    rowCount: number;
    colCount: number;
    headerRows?: number | null;
    headerCols?: number | null;
    defaultRowHeight?: number;
    defaultColWidth?: number;
    prefetchOverscanRows?: number;
    prefetchOverscanCols?: number;
    theme?: Partial<GridTheme>;
  }) {
    this.provider = options.provider;
    this.prefetchOverscanRows = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanRows);
    this.prefetchOverscanCols = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanCols);
    this.theme = resolveGridTheme(options.theme);
    this.baseDefaultRowHeight = options.defaultRowHeight ?? 21;
    this.baseDefaultColWidth = options.defaultColWidth ?? 100;
    this.scroll = new VirtualScrollManager({
      rowCount: options.rowCount,
      colCount: options.colCount,
      defaultRowHeight: this.baseDefaultRowHeight,
      defaultColWidth: this.baseDefaultColWidth
    });

    if (options.headerRows !== undefined || options.headerCols !== undefined) {
      this.setHeaders(options.headerRows ?? null, options.headerCols ?? null);
    }

    // Enable stats by default in dev builds where `import.meta.env.PROD` (Vite) is false.
    // In production builds this stays disabled to minimize overhead.
    const metaEnv = (import.meta as any)?.env as { PROD?: boolean } | undefined;
    const nodeEnv = (globalThis as any)?.process?.env?.NODE_ENV as string | undefined;
    const isProd = metaEnv?.PROD === true || nodeEnv === "production";
    this.perfStats.enabled = !isProd;
  }

  getPerfStats(): Readonly<GridPerfStats> {
    return this.perfStats;
  }

  setPerfStatsEnabled(enabled: boolean): void {
    this.perfStats.enabled = enabled;
  }

  setTheme(theme: Partial<GridTheme> | null | undefined): void {
    const next = resolveGridTheme(theme);
    if (gridThemesEqual(this.theme, next)) return;
    this.theme = next;
    this.markAllDirtyForThemeChange();
  }

  getTheme(): GridTheme {
    return this.theme;
  }

  /**
   * Configure the number of header rows/cols for styling.
   *
   * - When set (including `0`), header styling is determined solely by these counts.
   * - When `null`, the renderer falls back to legacy behavior: treat the first frozen
   *   row/col (if any) as the header region.
   */
  setHeaders(headerRows: number | null, headerCols: number | null): void {
    const { rowCount, colCount } = this.scroll.getCounts();

    const sanitize = (value: number | null, max: number): number | null => {
      if (value === null) return null;
      if (!Number.isFinite(value)) return 0;
      return clampIndex(value, 0, max);
    };

    const nextHeaderRows = sanitize(headerRows, rowCount);
    const nextHeaderCols = sanitize(headerCols, colCount);

    if (nextHeaderRows === this.headerRowsOverride && nextHeaderCols === this.headerColsOverride) return;

    this.headerRowsOverride = nextHeaderRows;
    this.headerColsOverride = nextHeaderCols;

    // If we're not attached yet, defer the repaint until the initial `attach()`/`resize()`.
    if (this.gridCtx) {
      this.markAllDirty();
    }
  }

  getZoom(): number {
    return this.zoom;
  }

  setZoom(nextZoom: number, options?: { anchorX?: number; anchorY?: number }): void {
    const clamped = clampZoom(nextZoom);
    if (clamped === this.zoom) return;

    const prevZoom = this.zoom;
    const prevScroll = this.scroll.getScroll();
    const prevViewport = this.scroll.getViewportState();
    const { rowCount, colCount } = this.scroll.getCounts();

    const baseScrollX = prevScroll.x / prevZoom;
    const baseScrollY = prevScroll.y / prevZoom;

    const nextScroll = new VirtualScrollManager({
      rowCount,
      colCount,
      defaultRowHeight: this.baseDefaultRowHeight * clamped,
      defaultColWidth: this.baseDefaultColWidth * clamped
    });
    nextScroll.setViewportSize(prevViewport.width, prevViewport.height);
    nextScroll.setFrozen(prevViewport.frozenRows, prevViewport.frozenCols);

    for (const [row, baseHeight] of this.rowHeightOverridesBase) {
      nextScroll.rows.setSize(row, baseHeight * clamped);
    }
    for (const [col, baseWidth] of this.colWidthOverridesBase) {
      nextScroll.cols.setSize(col, baseWidth * clamped);
    }

    const viewportAfter = nextScroll.getViewportState();

    let targetScrollX = baseScrollX * clamped;
    let targetScrollY = baseScrollY * clamped;

    const anchorX = options?.anchorX;
    if (
      anchorX !== undefined &&
      Number.isFinite(anchorX) &&
      // Only anchor when the zoom gesture is inside the scrollable quadrant both before and after
      // the zoom change. Frozen quadrants do not scroll, so attempting to anchor them produces
      // unexpected jumps when the frozen boundary moves due to zoom.
      anchorX >= prevViewport.frozenWidth &&
      anchorX >= viewportAfter.frozenWidth
    ) {
      const beforeSheetX = anchorX + prevScroll.x;
      const baseX = beforeSheetX / prevZoom;
      targetScrollX = baseX * clamped - anchorX;
    }

    const anchorY = options?.anchorY;
    if (
      anchorY !== undefined &&
      Number.isFinite(anchorY) &&
      anchorY >= prevViewport.frozenHeight &&
      anchorY >= viewportAfter.frozenHeight
    ) {
      const beforeSheetY = anchorY + prevScroll.y;
      const baseY = beforeSheetY / prevZoom;
      targetScrollY = baseY * clamped - anchorY;
    }

    this.scroll = nextScroll;
    this.zoom = clamped;
    this.presenceFont = `${12 * this.zoom}px system-ui, sans-serif`;

    this.scroll.setScroll(targetScrollX, targetScrollY);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);

    if (this.selectionCtx) {
      this.remotePresenceDirtyPadding = this.getRemotePresenceDirtyPadding(this.selectionCtx);
    }

    this.markAllDirty();
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

    if (this.selectionCtx) {
      this.remotePresenceDirtyPadding = this.getRemotePresenceDirtyPadding(this.selectionCtx);
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
    if (this.remotePresences.length > 0) {
      this.remotePresenceDirtyPadding = this.getRemotePresenceDirtyPadding(this.selectionCtx);
    }
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
    const nextRanges = selection
      ? [
          {
            startRow: selection.row,
            endRow: selection.row + 1,
            startCol: selection.col,
            endCol: selection.col + 1
          }
        ]
      : [];

    this.setSelectionRanges(nextRanges, { activeIndex: 0, activeCell: selection });
  }

  setSelectionRange(range: CellRange | null, options?: { activeCell?: Selection | null }): void {
    this.setSelectionRanges(range ? [range] : null, { activeIndex: 0, activeCell: options?.activeCell ?? undefined });
  }

  getSelectionRange(): CellRange | null {
    return this.selectionRanges.length === 0 ? null : { ...this.selectionRanges[this.activeSelectionIndex] };
  }

  getSelectionRanges(): CellRange[] {
    return this.selectionRanges.map((range) => ({ ...range }));
  }

  getActiveSelectionIndex(): number {
    return this.activeSelectionIndex;
  }

  setActiveSelectionIndex(index: number): void {
    if (!Number.isFinite(index)) return;
    if (this.selectionRanges.length === 0) return;
    const next = clamp(Math.trunc(index), 0, this.selectionRanges.length - 1);
    if (next === this.activeSelectionIndex) return;
    this.activeSelectionIndex = next;
    this.markSelectionDirty();
  }

  /**
   * Returns the fill-handle rectangle for the current selection range, in viewport
   * coordinates (relative to the grid canvases).
   *
   * The returned rectangle is clipped to the visible viewport and is `null` when
   * the handle is disabled or not visible.
   */
  getFillHandleRect(): Rect | null {
    if (!this.fillHandleEnabled) return null;
    if (this.selectionRanges.length === 0) return null;
    const range = this.selectionRanges[this.activeSelectionIndex];
    const viewport = this.scroll.getViewportState();
    return this.fillHandleRectInViewport(range, viewport);
  }

  setFillHandleEnabled(enabled: boolean): void {
    const next = Boolean(enabled);
    if (next === this.fillHandleEnabled) return;
    this.fillHandleEnabled = next;
    this.markSelectionDirty();
  }

  getSelection(): Selection | null {
    return this.selection ? { ...this.selection } : null;
  }

  setRangeSelection(range: CellRange | null): void {
    const previousRange = this.rangeSelection;
    const normalized = range ? this.normalizeSelectionRange(range) : null;
    const expanded = normalized ? this.expandRangeToMergedCells(normalized) : null;
    if (isSameCellRange(previousRange, expanded)) return;
    this.rangeSelection = expanded;
    this.markSelectionDirty();
  }

  setSelectionRanges(
    ranges: CellRange[] | null,
    options?: { activeIndex?: number; activeCell?: Selection | null }
  ): void {
    const previousRanges = this.selectionRanges;
    const previousActiveIndex = this.activeSelectionIndex;
    const previousCell = this.selection;

    const normalizedRanges = (ranges ?? [])
      .map((range) => this.normalizeSelectionRange(range))
      .filter((range): range is CellRange => range !== null)
      .map((range) => this.expandRangeToMergedCells(range));

    if (normalizedRanges.length === 0) {
      this.selection = null;
      this.selectionRanges = [];
      this.activeSelectionIndex = 0;
      if (previousRanges.length > 0 || previousCell) this.markSelectionDirty();
      return;
    }

    const requestedIndex = options?.activeIndex ?? this.activeSelectionIndex;
    const activeIndex = clampIndex(requestedIndex, 0, normalizedRanges.length - 1);

    const activeRange = normalizedRanges[activeIndex];
    const requestedCell = options?.activeCell ?? undefined;
    const baseCell = requestedCell ?? previousCell ?? { row: activeRange.startRow, col: activeRange.startCol };

    let nextSelection: Selection = {
      row: clamp(baseCell.row, activeRange.startRow, activeRange.endRow - 1),
      col: clamp(baseCell.col, activeRange.startCol, activeRange.endCol - 1)
    };

    const resolvedActive = this.getMergedAnchorForCell(nextSelection.row, nextSelection.col);
    if (resolvedActive) nextSelection = resolvedActive;

    this.selection = nextSelection;
    this.selectionRanges = normalizedRanges;
    this.activeSelectionIndex = activeIndex;

    const rangesChanged =
      previousRanges.length !== normalizedRanges.length ||
      previousRanges.some((range, idx) => !CanvasGridRenderer.rangesEqual(range, normalizedRanges[idx]!));

    const cellChanged = previousCell?.row !== nextSelection.row || previousCell?.col !== nextSelection.col;

    if (rangesChanged || previousActiveIndex !== activeIndex || cellChanged) {
      this.markSelectionDirty();
    }
  }

  addSelectionRange(range: CellRange): void {
    const normalized = this.normalizeSelectionRange(range);
    if (!normalized) return;

    const nextRanges = [...this.selectionRanges.map((r) => ({ ...r })), normalized];
    const nextIndex = nextRanges.length - 1;
    this.setSelectionRanges(nextRanges, { activeIndex: nextIndex, activeCell: { row: normalized.startRow, col: normalized.startCol } });
  }

  scrollToCell(row: number, col: number, opts?: { align?: ScrollToCellAlign; padding?: number }): void {
    const viewport = this.scroll.getViewportState();
    const { rowCount, colCount } = this.scroll.getCounts();

    if (!Number.isFinite(row) || !Number.isFinite(col)) return;
    if (rowCount === 0 || colCount === 0) return;

    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return;

    const align = opts?.align ?? "auto";
    const padding = Math.max(0, opts?.padding ?? 0);

    const current = this.scroll.getScroll();
    let targetX = current.x;
    let targetY = current.y;

    const colAxis = this.scroll.cols;
    const rowAxis = this.scroll.rows;

    const merged = this.getMergedRangeForCell(row, col);
    const crossesFrozenCols = merged ? merged.startCol < viewport.frozenCols && merged.endCol > viewport.frozenCols : false;
    const crossesFrozenRows = merged ? merged.startRow < viewport.frozenRows && merged.endRow > viewport.frozenRows : false;
    const useMerged = merged != null && !crossesFrozenCols && !crossesFrozenRows;

    // When a merged range crosses frozen boundaries, we can't represent its full bounds as a single
    // viewport-space rect (frozen and scrollable quadrants use different coordinate spaces), but we
    // *can* still scroll to ensure that the non-frozen portion of the merge becomes visible. Treat the
    // scrollable portion of the merge as the target in those cases.
    const rowStart = useMerged ? merged.startRow : crossesFrozenRows ? viewport.frozenRows : row;
    const rowEndExclusive = useMerged ? merged.endRow : crossesFrozenRows ? merged!.endRow : row + 1;
    const colStart = useMerged ? merged.startCol : crossesFrozenCols ? viewport.frozenCols : col;
    const colEndExclusive = useMerged ? merged.endCol : crossesFrozenCols ? merged!.endCol : col + 1;

    const viewStartX = viewport.frozenWidth + padding;
    const viewEndX = Math.max(viewStartX, viewport.width - padding);
    const viewCenterX = viewport.frozenWidth + Math.max(0, viewport.width - viewport.frozenWidth) / 2;
    const viewWidth = viewEndX - viewStartX;

    if (colStart >= viewport.frozenCols) {
      const colX = colAxis.positionOf(colStart);
      const width = colAxis.positionOf(colEndExclusive) - colX;
      const cellX = colX - current.x;
      const cellEnd = cellX + width;

      if (align === "start") {
        targetX = colX - viewStartX;
      } else if (align === "end") {
        targetX = colX + width - viewEndX;
      } else if (align === "center") {
        targetX = colX + width / 2 - viewCenterX;
      } else {
        if (cellX < viewStartX) {
          targetX = colX - viewStartX;
        } else if (cellEnd > viewEndX) {
          // If the cell is wider than the viewport, prefer aligning the left edge so the anchor
          // remains visible (Excel-style behavior).
          targetX = width > viewWidth ? colX - viewStartX : colX + width - viewEndX;
        }
      }
    }

    const viewStartY = viewport.frozenHeight + padding;
    const viewEndY = Math.max(viewStartY, viewport.height - padding);
    const viewCenterY = viewport.frozenHeight + Math.max(0, viewport.height - viewport.frozenHeight) / 2;
    const viewHeight = viewEndY - viewStartY;

    if (rowStart >= viewport.frozenRows) {
      const rowY = rowAxis.positionOf(rowStart);
      const height = rowAxis.positionOf(rowEndExclusive) - rowY;
      const cellY = rowY - current.y;
      const cellEnd = cellY + height;

      if (align === "start") {
        targetY = rowY - viewStartY;
      } else if (align === "end") {
        targetY = rowY + height - viewEndY;
      } else if (align === "center") {
        targetY = rowY + height / 2 - viewCenterY;
      } else {
        if (cellY < viewStartY) {
          targetY = rowY - viewStartY;
        } else if (cellEnd > viewEndY) {
          targetY = height > viewHeight ? rowY - viewStartY : rowY + height - viewEndY;
        }
      }
    }

    const { maxScrollX, maxScrollY } = this.scroll.getMaxScroll();
    targetX = clamp(targetX, 0, maxScrollX);
    targetY = clamp(targetY, 0, maxScrollY);

    if (targetX !== current.x || targetY !== current.y) {
      this.setScroll(targetX, targetY);
    }
  }

  getCellRect(row: number, col: number): Rect | null {
    const viewport = this.scroll.getViewportState();
    const merged = this.getMergedRangeForCell(row, col);
    if (!merged) return this.cellRectInViewport(row, col, viewport, { clampToViewport: false });

    // Merged regions that cross frozen boundaries cannot be represented as a single viewport-space
    // rect because frozen and scrollable quadrants use different coordinate spaces. Fall back to the
    // anchor cell rect in those cases.
    const crossesFrozenCols = merged.startCol < viewport.frozenCols && merged.endCol > viewport.frozenCols;
    const crossesFrozenRows = merged.startRow < viewport.frozenRows && merged.endRow > viewport.frozenRows;
    if (crossesFrozenCols || crossesFrozenRows) {
      return this.cellRectInViewport(merged.startRow, merged.startCol, viewport, { clampToViewport: false });
    }

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const colX = colAxis.positionOf(merged.startCol);
    const rowY = rowAxis.positionOf(merged.startRow);
    const width = colAxis.positionOf(merged.endCol) - colX;
    const height = rowAxis.positionOf(merged.endRow) - rowY;

    const scrollCols = merged.startCol >= viewport.frozenCols;
    const scrollRows = merged.startRow >= viewport.frozenRows;

    const x = scrollCols ? colX - viewport.scrollX : colX;
    const y = scrollRows ? rowY - viewport.scrollY : rowY;

    return { x, y, width, height };
  }

  getViewportState(): GridViewportState {
    return this.scroll.getViewportState();
  }

  setFillPreviewRange(range: CellRange | null): void {
    const previousRange = this.fillPreviewRange;
    const normalized = range ? this.normalizeSelectionRange(range) : null;
    if (isSameCellRange(previousRange, normalized)) return;
    this.fillPreviewRange = normalized;

    const viewport = this.scroll.getViewportState();
    const padding = 4;
    const dirtyRects = [
      ...this.fillPreviewOverlayRects(previousRange, viewport),
      ...this.fillPreviewOverlayRects(normalized, viewport)
    ];
    if (dirtyRects.length === 0) return;

    for (const rect of dirtyRects) {
      this.dirty.selection.markDirty(padRect(rect, padding));
    }

    this.requestRender();
  }

  setReferenceHighlights(highlights: Array<{ range: CellRange; color: string; active?: boolean }> | null): void {
    const normalized =
      highlights
        ?.map((h) => {
          const range = this.normalizeSelectionRange(h.range);
          if (!range) return null;
          return { range, color: h.color, active: Boolean(h.active) };
        })
        .filter((h): h is { range: CellRange; color: string; active: boolean } => h !== null) ?? [];

    if (
      normalized.length === this.referenceHighlights.length &&
        normalized.every((h, i) => {
          const prev = this.referenceHighlights[i];
          return prev && prev.color === h.color && prev.active === h.active && isSameCellRange(prev.range, h.range);
        })
    ) {
      return;
    }

    this.referenceHighlights = normalized;
    const viewport = this.scroll.getViewportState();
    this.dirty.selection.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    this.requestRender();
  }

  setRemotePresences(presences: GridPresence[] | null): void {
    if (presences === this.remotePresences) return;
    this.remotePresences = presences ?? [];
    if (this.selectionCtx) {
      this.remotePresenceDirtyPadding = this.getRemotePresenceDirtyPadding(this.selectionCtx);
    } else {
      this.remotePresenceDirtyPadding = 1;
    }

    const viewport = this.scroll.getViewportState();
    this.dirty.selection.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    this.requestRender();
  }

  getRowHeight(row: number): number {
    this.assertRowIndex(row);
    return this.scroll.rows.getSize(row);
  }

  getColWidth(col: number): number {
    this.assertColIndex(col);
    return this.scroll.cols.getSize(col);
  }

  setRowHeight(row: number, height: number): void {
    this.assertRowIndex(row);
    if (Math.abs(height - this.scroll.rows.defaultSize) < 1e-6) {
      this.rowHeightOverridesBase.delete(row);
    } else {
      this.rowHeightOverridesBase.set(row, height / this.zoom);
    }
    this.scroll.rows.setSize(row, height);
    this.onAxisSizeChanged();
  }

  setColWidth(col: number, width: number): void {
    this.assertColIndex(col);
    if (Math.abs(width - this.scroll.cols.defaultSize) < 1e-6) {
      this.colWidthOverridesBase.delete(col);
    } else {
      this.colWidthOverridesBase.set(col, width / this.zoom);
    }
    this.scroll.cols.setSize(col, width);
    this.onAxisSizeChanged();
  }

  resetRowHeight(row: number): void {
    this.assertRowIndex(row);
    this.rowHeightOverridesBase.delete(row);
    this.scroll.rows.deleteSize(row);
    this.onAxisSizeChanged();
  }

  resetColWidth(col: number): void {
    this.assertColIndex(col);
    this.colWidthOverridesBase.delete(col);
    this.scroll.cols.deleteSize(col);
    this.onAxisSizeChanged();
  }

  /**
   * Auto-fit a column based on the widest visible cell text in the current viewport.
   *
   * Note: This intentionally only scans a bounded viewport window (visible range + small overscan)
   * for ergonomics/performance.
   */
  autoFitCol(col: number, options?: { maxWidth?: number }): number {
    this.assertColIndex(col);
    const viewport = this.scroll.getViewportState();

    const zoom = this.zoom;
    const maxWidth = (options?.maxWidth ?? 500) * zoom;
    const minWidth = 24 * zoom;
    const paddingX = 4 * zoom;
    const extraPadding = 8 * zoom;
    const overscanRows = 2;

    const rowCount = this.getRowCount();
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const frozenRowsRange =
      viewport.frozenRows === 0 || frozenHeightClamped === 0
        ? { start: 0, end: 0 }
        : this.scroll.rows.visibleRange(0, frozenHeightClamped, { min: 0, maxExclusive: viewport.frozenRows });

    const viewportScrollableHeight = Math.max(0, viewport.height - frozenHeightClamped);
    const mainRowsRange =
      viewportScrollableHeight === 0 || rowCount === viewport.frozenRows
        ? { start: 0, end: 0 }
        : {
            start: Math.max(viewport.frozenRows, viewport.main.rows.start - overscanRows),
            end: Math.min(rowCount, viewport.main.rows.end + overscanRows)
          };

    const layoutEngine = this.textLayoutEngine;

    let maxMeasured = 0;
    const measureRow = (row: number) => {
      const cell = this.provider.getCell(row, col);
      if (!cell) return;
      const richText = cell.richText;
      const richTextText = richText?.text ?? "";
      const hasRichText = Boolean(richText && richTextText);
      if (cell.value === null && !hasRichText) return;

      const style = cell.style;
      const fontSize = (style?.fontSize ?? 12) * zoom;
      const fontFamily = style?.fontFamily ?? "system-ui";
      const fontWeight = style?.fontWeight ?? "400";
      const fontStyle = style?.fontStyle ?? "normal";
      const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle } satisfies FontSpec;

      if (hasRichText) {
        const offsets = buildCodePointIndex(richTextText);
        const textLen = offsets.length - 1;
        const rawRuns =
          Array.isArray(richText?.runs) && richText.runs.length > 0
            ? richText.runs
            : [{ start: 0, end: textLen, style: undefined }];

        const defaults = {
          family: fontFamily,
          sizePx: fontSize,
          weight: fontWeight,
          style: fontStyle
        } satisfies Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number; style: string };

        const layoutRuns = rawRuns.map((run) => {
          const runStyle =
            run.style && typeof run.style === "object" ? (run.style as RichTextRunStyle) : (undefined as RichTextRunStyle | undefined);
          return {
            text: sliceByCodePointRange(richTextText, offsets, run.start, run.end),
            font: fontSpecForRichTextStyle(runStyle, defaults, zoom)
          };
        });

        let maxLineWidth = 0;
        let currentLineWidth = 0;
        const newlineRe = /\r\n|\r|\n/g;

        for (const run of layoutRuns) {
          if (!run.text) continue;

          let lastIndex = 0;
          newlineRe.lastIndex = 0;
          let match: RegExpExecArray | null;

          while ((match = newlineRe.exec(run.text))) {
            const segment = run.text.slice(lastIndex, match.index);
            if (segment.length > 0) {
              const measurement = layoutEngine?.measure(segment, run.font);
              const width = measurement?.width ?? segment.length * run.font.sizePx * 0.6;
              if (Number.isFinite(width)) currentLineWidth += width;
            }

            maxLineWidth = Math.max(maxLineWidth, currentLineWidth);
            currentLineWidth = 0;
            lastIndex = match.index + match[0].length;
          }

          const tail = run.text.slice(lastIndex);
          if (tail.length > 0) {
            const measurement = layoutEngine?.measure(tail, run.font);
            const width = measurement?.width ?? tail.length * run.font.sizePx * 0.6;
            if (Number.isFinite(width)) currentLineWidth += width;
          }
        }

        maxLineWidth = Math.max(maxLineWidth, currentLineWidth);
        if (Number.isFinite(maxLineWidth)) maxMeasured = Math.max(maxMeasured, maxLineWidth);
        return;
      }

      const text = formatCellDisplayText(cell.value);
      if (text === "") return;

      const measurement = layoutEngine?.measure(text, fontSpec);
      const width = measurement?.width ?? text.length * fontSize * 0.6;
      if (Number.isFinite(width)) maxMeasured = Math.max(maxMeasured, width);
    };

    for (let row = frozenRowsRange.start; row < frozenRowsRange.end; row++) measureRow(row);
    for (let row = mainRowsRange.start; row < mainRowsRange.end; row++) measureRow(row);

    const next = clamp(Math.ceil(maxMeasured + paddingX * 2 + extraPadding), minWidth, maxWidth);
    this.setColWidth(col, next);
    return next;
  }

  /**
   * Auto-fit a row based on text height for visible cells in the current viewport.
   *
   * If wrapping is enabled on any visible cell, wrapped layout height is used.
   * Otherwise font metrics are used so large fonts can still auto-fit.
   */
  autoFitRow(row: number, options?: { maxHeight?: number }): number {
    this.assertRowIndex(row);
    const viewport = this.scroll.getViewportState();

    const zoom = this.zoom;
    const maxHeight = (options?.maxHeight ?? 500) * zoom;
    const paddingX = 4 * zoom;
    const paddingY = 2 * zoom;
    const overscanCols = 2;

    const colCount = this.getColCount();
    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenColsRange =
      viewport.frozenCols === 0 || frozenWidthClamped === 0
        ? { start: 0, end: 0 }
        : this.scroll.cols.visibleRange(0, frozenWidthClamped, { min: 0, maxExclusive: viewport.frozenCols });

    const viewportScrollableWidth = Math.max(0, viewport.width - frozenWidthClamped);
    const mainColsRange =
      viewportScrollableWidth === 0 || colCount === viewport.frozenCols
        ? { start: 0, end: 0 }
        : {
            start: Math.max(viewport.frozenCols, viewport.main.cols.start - overscanCols),
            end: Math.min(colCount, viewport.main.cols.end + overscanCols)
          };

    const layoutEngine = this.textLayoutEngine;
    const defaultHeight = this.scroll.rows.defaultSize;
    if (!layoutEngine) return defaultHeight;

    let hasContent = false;
    let maxMeasuredHeight = defaultHeight;

    const measureCol = (col: number) => {
      const cell = this.provider.getCell(row, col);
      if (!cell) return;
      const richText = cell.richText;
      const richTextText = richText?.text ?? "";
      const hasRichText = Boolean(richText && richTextText);
      if (cell.value === null && !hasRichText) return;

      const style = cell.style;
      const wrapMode = style?.wrapMode ?? "none";

      const text = hasRichText ? richTextText : formatCellDisplayText(cell.value);
      if (text === "") return;

      hasContent = true;

      const fontSize = (style?.fontSize ?? 12) * zoom;
      const fontFamily = style?.fontFamily ?? "system-ui";
      const fontWeight = style?.fontWeight ?? "400";
      const fontStyle = style?.fontStyle ?? "normal";
      const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle } satisfies FontSpec;

      const defaults = {
        family: fontFamily,
        sizePx: fontSize,
        weight: fontWeight,
        style: fontStyle
      } satisfies Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number; style: string };

      const layoutRuns = (() => {
        if (!hasRichText) return null;

        const offsets = buildCodePointIndex(richTextText);
        const textLen = offsets.length - 1;
        const rawRuns =
          Array.isArray(richText?.runs) && richText.runs.length > 0
            ? richText.runs
            : [{ start: 0, end: textLen, style: undefined }];

        return rawRuns.map((run) => {
          const runStyle =
            run.style && typeof run.style === "object" ? (run.style as RichTextRunStyle) : (undefined as RichTextRunStyle | undefined);
          return {
            text: sliceByCodePointRange(richTextText, offsets, run.start, run.end),
            font: fontSpecForRichTextStyle(runStyle, defaults, zoom)
          };
        });
      })();

      const maxFontSizePx = layoutRuns ? layoutRuns.reduce((acc, run) => Math.max(acc, run.font.sizePx), defaults.sizePx) : fontSize;
      const lineHeight = Math.ceil(maxFontSizePx * 1.2);
      // Even without wrapping, large fonts should auto-fit the row height.
      maxMeasuredHeight = Math.max(maxMeasuredHeight, lineHeight + paddingY * 2);

      if (wrapMode === "none") return;

      const colWidth = this.scroll.cols.getSize(col);
      const availableWidth = Math.max(0, colWidth - paddingX * 2);
      if (availableWidth === 0) return;

      const align: CanvasTextAlign =
        style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
      const layoutAlign =
        align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
          ? (align as "left" | "right" | "center" | "start" | "end")
          : "start";
      const direction = style?.direction ?? "auto";

      const layout = layoutRuns
        ? layoutEngine.layout({
            runs: layoutRuns.map((r) => ({ text: r.text, font: r.font })),
            text: undefined,
            font: defaults,
            maxWidth: availableWidth,
            wrapMode,
            align: layoutAlign,
            direction,
            lineHeightPx: lineHeight,
            maxLines: 1000
          })
        : layoutEngine.layout({
            text,
            font: fontSpec,
            maxWidth: availableWidth,
            wrapMode,
            align: layoutAlign,
            direction,
            lineHeightPx: lineHeight,
            maxLines: 1000
          });

      const height = layout.height + paddingY * 2;
      if (Number.isFinite(height)) maxMeasuredHeight = Math.max(maxMeasuredHeight, height);
    };

    for (let col = frozenColsRange.start; col < frozenColsRange.end; col++) measureCol(col);
    for (let col = mainColsRange.start; col < mainColsRange.end; col++) measureCol(col);

    if (!hasContent) {
      this.resetRowHeight(row);
      return this.scroll.rows.defaultSize;
    }

    const next = clamp(Math.ceil(maxMeasuredHeight), defaultHeight, maxHeight);
    this.setRowHeight(row, next);
    return next;
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

    const resolved = this.getMergedIndex(viewport).resolveCell({ row, col });
    return { row: resolved.row, col: resolved.col };
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
    const perf = this.perfStats;
    const perfEnabled = perf.enabled;
    const frameStart = perfEnabled ? performance.now() : 0;

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
      const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
      if (!this.forceFullRedraw && this.canBlitScroll(viewport, scrollDeltaX, scrollDeltaY)) {
        this.blitScroll(viewport, scrollDeltaX, scrollDeltaY);
        if (perfEnabled) perf.blitUsed = true;
        this.markScrollDirtyRegions(viewport, scrollDeltaX, scrollDeltaY);
      } else {
        if (perfEnabled) perf.blitUsed = false;
        this.markFullViewportDirty(viewport);
        this.dirty.selection.markDirty(full);
      }
    } else if (perfEnabled) {
      perf.blitUsed = false;
    }

    const backgroundRegions = this.dirty.background.drain();
    const contentRegions = this.dirty.content.drain();
    const selectionRegions = this.dirty.selection.drain();

    if (perfEnabled) {
      perf.dirtyRects.background = backgroundRegions.length;
      perf.dirtyRects.content = contentRegions.length;
      perf.dirtyRects.selection = selectionRegions.length;
      perf.dirtyRects.total = backgroundRegions.length + contentRegions.length + selectionRegions.length;
      perf.cellsPainted = 0;
      perf.cellFetches = 0;
    }

    const mergedIndex = this.getMergedIndex(viewport);

    this.renderGridLayers(viewport, mergedIndex, backgroundRegions, contentRegions, perfEnabled ? perf : null);
    this.renderLayer("selection", viewport, mergedIndex, selectionRegions);

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

    if (perfEnabled) {
      perf.lastFrameMs = performance.now() - frameStart;
    }
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
    if (!this.gridCanvas || !this.contentCanvas || !this.selectionCanvas) return false;
    if (!this.gridCtx || !this.contentCtx || !this.selectionCtx) return false;
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
    this.blitLayer("selection", viewport, deltaX, deltaY);
  }

  private blitLayer(layer: "background" | "content" | "selection", viewport: GridViewportState, deltaX: number, deltaY: number): void {
    const canvas = layer === "background" ? this.gridCanvas : layer === "content" ? this.contentCanvas : this.selectionCanvas;
    const ctx = layer === "background" ? this.gridCtx : layer === "content" ? this.contentCtx : this.selectionCtx;
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

    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);

    // Frozen rows + scrollable columns quadrant (top-right): horizontal-only shift.
    {
      const rectX = frozenWidthPx;
      const rectY = 0;
      const rectW = widthPx - frozenWidthPx;
      const rectH = frozenHeightPx;
      if (rectW > 0 && rectH > 0 && dx !== 0) {
        if (layer === "background") {
          ctx.fillStyle = this.theme.gridBg;
          ctx.fillRect(rectX, rectY, rectW, rectH);
        } else {
          ctx.clearRect(rectX, rectY, rectW, rectH);
        }

        ctx.save();
        ctx.beginPath();
        ctx.rect(rectX, rectY, rectW, rectH);
        ctx.clip();
        ctx.drawImage(this.blitCanvas, dx, 0);
        ctx.restore();
      }
    }

    // Scrollable rows + frozen columns quadrant (bottom-left): vertical-only shift.
    {
      const rectX = 0;
      const rectY = frozenHeightPx;
      const rectW = frozenWidthPx;
      const rectH = heightPx - frozenHeightPx;
      if (rectW > 0 && rectH > 0 && dy !== 0) {
        if (layer === "background") {
          ctx.fillStyle = this.theme.gridBg;
          ctx.fillRect(rectX, rectY, rectW, rectH);
        } else {
          ctx.clearRect(rectX, rectY, rectW, rectH);
        }

        ctx.save();
        ctx.beginPath();
        ctx.rect(rectX, rectY, rectW, rectH);
        ctx.clip();
        ctx.drawImage(this.blitCanvas, 0, dy);
        ctx.restore();
      }
    }

    // Scrollable quadrant (main): both-axis shift.
    {
      const rectX = frozenWidthPx;
      const rectY = frozenHeightPx;
      const rectW = widthPx - frozenWidthPx;
      const rectH = heightPx - frozenHeightPx;
      if (rectW > 0 && rectH > 0 && (dx !== 0 || dy !== 0)) {
        if (layer === "background") {
          ctx.fillStyle = this.theme.gridBg;
          ctx.fillRect(rectX, rectY, rectW, rectH);
        } else {
          ctx.clearRect(rectX, rectY, rectW, rectH);
        }

        ctx.save();
        ctx.beginPath();
        ctx.rect(rectX, rectY, rectW, rectH);
        ctx.clip();
        ctx.drawImage(this.blitCanvas, dx, dy);
        ctx.restore();
      }
    }

    ctx.restore();
  }

  private markScrollDirtyRegions(viewport: GridViewportState, deltaX: number, deltaY: number): void {
    const frozenWidth = viewport.frozenWidth;
    const frozenHeight = viewport.frozenHeight;

    const selectionPadding = 1;

    // Frozen rows + scrollable columns (top-right): horizontal-only scroll.
    {
      const rectX = frozenWidth;
      const rectY = 0;
      const rectW = viewport.width - frozenWidth;
      const rectH = frozenHeight;
      if (rectW > 0 && rectH > 0) {
        const shiftX = deltaX;
        if (shiftX > 0) {
          const stripe = { x: rectX, y: rectY, width: shiftX, height: rectH };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        } else if (shiftX < 0) {
          const stripe = { x: rectX + rectW + shiftX, y: rectY, width: -shiftX, height: rectH };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        }
      }
    }

    // Scrollable rows + frozen columns (bottom-left): vertical-only scroll.
    {
      const rectX = 0;
      const rectY = frozenHeight;
      const rectW = frozenWidth;
      const rectH = viewport.height - frozenHeight;
      if (rectW > 0 && rectH > 0) {
        const shiftY = deltaY;
        if (shiftY > 0) {
          const stripe = { x: rectX, y: rectY, width: rectW, height: shiftY };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        } else if (shiftY < 0) {
          const stripe = { x: rectX, y: rectY + rectH + shiftY, width: rectW, height: -shiftY };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        }
      }
    }

    // Main scrollable quadrant: both-axis scroll.
    {
      const rectX = frozenWidth;
      const rectY = frozenHeight;
      const rectW = viewport.width - frozenWidth;
      const rectH = viewport.height - frozenHeight;
      if (rectW > 0 && rectH > 0) {
        const shiftX = deltaX;
        const shiftY = deltaY;
        if (shiftX > 0) {
          const stripe = { x: rectX, y: rectY, width: shiftX, height: rectH };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        } else if (shiftX < 0) {
          const stripe = { x: rectX + rectW + shiftX, y: rectY, width: -shiftX, height: rectH };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        }

        if (shiftY > 0) {
          const stripe = { x: rectX, y: rectY, width: rectW, height: shiftY };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        } else if (shiftY < 0) {
          const stripe = { x: rectX, y: rectY + rectH + shiftY, width: rectW, height: -shiftY };
          this.markDirtyBoth(stripe);
          this.markDirtySelection(stripe, selectionPadding);
        }
      }
    }

    if (this.remotePresences.length > 0) {
      // Cursor name badges can overlap into frozen quadrants, which aren't shifted during blit.
      // Mark the (padded) cursor rects dirty in both the previous and next viewport so badges
      // are cleared/redrawn correctly.
      const previousViewport = {
        ...viewport,
        scrollX: viewport.scrollX + deltaX,
        scrollY: viewport.scrollY + deltaY
      } as GridViewportState;

      for (const presence of this.remotePresences) {
        const cursor = presence.cursor;
        if (!cursor) continue;

        const previousRect = this.cellRectInViewport(cursor.row, cursor.col, previousViewport);
        const nextRect = this.cellRectInViewport(cursor.row, cursor.col, viewport);
        if (previousRect) this.markDirtySelection(previousRect, this.remotePresenceDirtyPadding);
        if (nextRect) this.markDirtySelection(nextRect, this.remotePresenceDirtyPadding);
      }
    }

    // Freeze lines are drawn on the selection layer but should not move with scroll. When we blit
    // the selection layer, the previous freeze line pixels get shifted into the scrollable
    // quadrants, leaving "ghost" lines behind. Mark those shifted lines as dirty so they get
    // cleared and redrawn in the correct location.
    const ghostWidth = 6;
    if (viewport.frozenCols > 0 && deltaX !== 0) {
      const ghostX = crispLine(viewport.frozenWidth) + deltaX;
      this.dirty.selection.markDirty({ x: ghostX - ghostWidth, y: 0, width: ghostWidth * 2, height: viewport.height });
    }
    if (viewport.frozenRows > 0 && deltaY !== 0) {
      const ghostY = crispLine(viewport.frozenHeight) + deltaY;
      this.dirty.selection.markDirty({ x: 0, y: ghostY - ghostWidth, width: viewport.width, height: ghostWidth * 2 });
    }
  }

  private markDirtyBoth(rect: Rect): void {
    const padded: Rect = { x: rect.x - 1, y: rect.y - 1, width: rect.width + 2, height: rect.height + 2 };
    this.dirty.background.markDirty(padded);
    this.dirty.content.markDirty(padded);
  }

  private markDirtySelection(rect: Rect, padding: number): void {
    const p = Number.isFinite(padding) ? Math.max(1, Math.floor(padding)) : 1;
    const padded: Rect = { x: rect.x - p, y: rect.y - p, width: rect.width + p * 2, height: rect.height + p * 2 };
    this.dirty.selection.markDirty(padded);
  }

  private getRemotePresenceDirtyPadding(ctx: CanvasRenderingContext2D): number {
    if (this.remotePresences.length === 0) return 1;

    // Keep in sync with `renderRemotePresenceOverlays`.
    const zoom = this.zoom;
    const badgePaddingX = 6 * zoom;
    const badgePaddingY = 3 * zoom;
    const badgeOffsetX = 8 * zoom;
    const badgeOffsetY = -18 * zoom;
    const badgeTextHeight = 14 * zoom;
    const cursorStrokeWidth = 2 * zoom;

    const previousFont = ctx.font;
    ctx.font = this.presenceFont;

    let padding = cursorStrokeWidth + 4 * zoom;
    for (const presence of this.remotePresences) {
      const name = presence.name ?? "Anonymous";
      const metricsKey = `${this.presenceFont}::${name}`;
      let textWidth = this.textWidthCache.get(metricsKey);
      if (textWidth === undefined) {
        textWidth = ctx.measureText(name).width;
        this.textWidthCache.set(metricsKey, textWidth);
      }

      const badgeWidth = textWidth + badgePaddingX * 2;
      const badgeHeight = badgeTextHeight + badgePaddingY * 2;
      const padX = badgeOffsetX + badgeWidth;
      const padY = Math.max(0, -badgeOffsetY) + badgeHeight;
      padding = Math.max(padding, padX, padY);
    }

    ctx.font = previousFont;
    return Math.ceil(padding);
  }

  private onProviderUpdate(update: CellProviderUpdate): void {
    this.mergedIndexDirty = true;
    if (update.type === "invalidateAll") {
      this.markAllDirty();
      return;
    }

    const viewport = this.scroll.getViewportState();
    const normalized = this.normalizeSelectionRange(update.range);
    if (!normalized) {
      this.requestRender();
      return;
    }

    const expanded = this.expandRangeToMergedCells(normalized);
    const colCount = this.getColCount();
    const overflowExpanded: CellRange = {
      startRow: expanded.startRow,
      endRow: expanded.endRow,
      startCol: Math.max(0, expanded.startCol - MAX_TEXT_OVERFLOW_COLUMNS),
      endCol: Math.min(colCount, expanded.endCol + MAX_TEXT_OVERFLOW_COLUMNS)
    };

    const rects = this.rangeToViewportRects(overflowExpanded, viewport);
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

  private renderGridLayers(
    viewport: GridViewportState,
    mergedIndex: MergedCellIndex,
    backgroundRegions: Rect[],
    contentRegions: Rect[],
    perf: GridPerfStats | null
  ): void {
    if (!this.gridCtx || !this.contentCtx) return;
    if (viewport.width === 0 || viewport.height === 0) return;

    const regions = CanvasGridRenderer.mergeDirtyRegions(backgroundRegions, contentRegions);
    if (regions.length === 0) return;

    const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    const shouldFullRender =
      regions.length > 8 ||
      regions.some((region) => region.x <= 0 && region.y <= 0 && region.width >= viewport.width && region.height >= viewport.height);
    const toRender = shouldFullRender ? [full] : regions;

    const gridCtx = this.gridCtx;
    const contentCtx = this.contentCtx;

    for (const region of toRender) {
      gridCtx.fillStyle = this.theme.gridBg;
      gridCtx.fillRect(region.x, region.y, region.width, region.height);

      contentCtx.clearRect(region.x, region.y, region.width, region.height);

      this.renderGridQuadrants(viewport, mergedIndex, region, perf);
    }
  }

  private static rectsEqual(a: Rect, b: Rect): boolean {
    return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
  }

  private static mergeDirtyRegions(primary: Rect[], secondary: Rect[]): Rect[] {
    if (primary.length === 0) return secondary;
    if (secondary.length === 0) return primary;
    if (primary.length === secondary.length) {
      let equal = true;
      for (let i = 0; i < primary.length; i++) {
        if (!CanvasGridRenderer.rectsEqual(primary[i], secondary[i])) {
          equal = false;
          break;
        }
      }
      if (equal) return primary;
    }

    // Merge secondary rects into primary, roughly matching DirtyRegionTracker's overlap merging.
    for (const rect of secondary) {
      let merged = rect;
      for (let i = 0; i < primary.length; ) {
        const existing = primary[i];
        const overlaps =
          existing.x < merged.x + merged.width &&
          existing.x + existing.width > merged.x &&
          existing.y < merged.y + merged.height &&
          existing.y + existing.height > merged.y;
        if (overlaps) {
          const x1 = Math.min(existing.x, merged.x);
          const y1 = Math.min(existing.y, merged.y);
          const x2 = Math.max(existing.x + existing.width, merged.x + merged.width);
          const y2 = Math.max(existing.y + existing.height, merged.y + merged.height);
          merged = { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
          primary.splice(i, 1);
          continue;
        }
        i++;
      }
      primary.push(merged);
    }

    return primary;
  }

  private renderLayer(layer: Layer, viewport: GridViewportState, mergedIndex: MergedCellIndex, regions: Rect[]): void {
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

    if (layer === "selection") {
      // Selection primitives (selection fill/stroke, remote presence overlays) are already expressed in
      // viewport coordinates, so we can render them once per frame clipped to the union of dirty rects.
      // This avoids re-walking quadrants and repeatedly recomputing selection rects.
      for (const region of toRender) {
        ctx.clearRect(region.x, region.y, region.width, region.height);
      }

      ctx.save();
      ctx.beginPath();
      for (const region of toRender) {
        ctx.rect(region.x, region.y, region.width, region.height);
      }
      ctx.clip();

      this.renderSelectionQuadrant(full, viewport, mergedIndex);
      if (this.remotePresences.length > 0) {
        this.renderRemotePresenceOverlays(ctx, viewport, mergedIndex);
      }

      ctx.restore();
      this.drawFreezeLines(ctx, viewport);
      return;
    }

    for (const region of toRender) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(region.x, region.y, region.width, region.height);
      ctx.clip();

      if (layer === "background") {
        ctx.fillStyle = this.theme.gridBg;
        ctx.fillRect(region.x, region.y, region.width, region.height);
      } else {
        ctx.clearRect(region.x, region.y, region.width, region.height);
      }

      this.renderQuadrants(layer, viewport, region);

      ctx.restore();
    }
  }

  private renderGridQuadrants(viewport: GridViewportState, mergedIndex: MergedCellIndex, region: Rect, perf: GridPerfStats | null): void {
    if (!this.gridCtx || !this.contentCtx) return;

    const { frozenCols, frozenRows, frozenWidth, frozenHeight, width, height } = viewport;
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    const absScrollX = frozenWidth + viewport.scrollX;
    const absScrollY = frozenHeight + viewport.scrollY;

    const quadrants = this.gridQuadrantScratch;

    const q0 = quadrants[0];
    q0.originX = 0;
    q0.originY = 0;
    q0.rect.x = 0;
    q0.rect.y = 0;
    q0.rect.width = frozenWidth;
    q0.rect.height = frozenHeight;
    q0.minRow = 0;
    q0.maxRowExclusive = frozenRows;
    q0.minCol = 0;
    q0.maxColExclusive = frozenCols;
    q0.scrollBaseX = 0;
    q0.scrollBaseY = 0;

    const q1 = quadrants[1];
    q1.originX = frozenWidth;
    q1.originY = 0;
    q1.rect.x = frozenWidth;
    q1.rect.y = 0;
    q1.rect.width = width - frozenWidth;
    q1.rect.height = frozenHeight;
    q1.minRow = 0;
    q1.maxRowExclusive = frozenRows;
    q1.minCol = frozenCols;
    q1.maxColExclusive = colCount;
    q1.scrollBaseX = absScrollX;
    q1.scrollBaseY = 0;

    const q2 = quadrants[2];
    q2.originX = 0;
    q2.originY = frozenHeight;
    q2.rect.x = 0;
    q2.rect.y = frozenHeight;
    q2.rect.width = frozenWidth;
    q2.rect.height = height - frozenHeight;
    q2.minRow = frozenRows;
    q2.maxRowExclusive = rowCount;
    q2.minCol = 0;
    q2.maxColExclusive = frozenCols;
    q2.scrollBaseX = 0;
    q2.scrollBaseY = absScrollY;

    const q3 = quadrants[3];
    q3.originX = frozenWidth;
    q3.originY = frozenHeight;
    q3.rect.x = frozenWidth;
    q3.rect.y = frozenHeight;
    q3.rect.width = width - frozenWidth;
    q3.rect.height = height - frozenHeight;
    q3.minRow = frozenRows;
    q3.maxRowExclusive = rowCount;
    q3.minCol = frozenCols;
    q3.maxColExclusive = colCount;
    q3.scrollBaseX = absScrollX;
    q3.scrollBaseY = absScrollY;

    const gridCtx = this.gridCtx;
    const contentCtx = this.contentCtx;

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

      if (perf?.enabled) {
        perf.cellsPainted += (endRow - startRow) * (endCol - startCol);
      }

      // Clip to the quadrant intersection so partially-visible edge cells don't bleed into other quadrants.
      gridCtx.save();
      gridCtx.beginPath();
      gridCtx.rect(intersection.x, intersection.y, intersection.width, intersection.height);
      gridCtx.clip();

      contentCtx.save();
      contentCtx.beginPath();
      contentCtx.rect(intersection.x, intersection.y, intersection.width, intersection.height);
      contentCtx.clip();

      const headerRows = this.headerRowsOverride ?? (viewport.frozenRows > 0 ? 1 : 0);
      const headerCols = this.headerColsOverride ?? (viewport.frozenCols > 0 ? 1 : 0);
      this.renderGridQuadrant(quadrant, mergedIndex, startRow, endRow, startCol, endCol, headerRows, headerCols, perf);

      contentCtx.restore();
      gridCtx.restore();
    }
  }

  private renderGridQuadrant(
    quadrant: {
      originX: number;
      originY: number;
      scrollBaseX: number;
      scrollBaseY: number;
    },
    mergedIndex: MergedCellIndex,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
    headerRows: number,
    headerCols: number,
    perf: GridPerfStats | null
  ): void {
    if (!this.gridCtx || !this.contentCtx) return;
    const gridCtx = this.gridCtx;
    const contentCtx = this.contentCtx;
    const theme = this.theme;
    const gridBg = theme.gridBg;
    const textColor = theme.cellText;
    const headerBg = theme.headerBg;
    const headerTextColor = theme.headerText;
    const errorTextColor = theme.errorText;
    const commentIndicator = theme.commentIndicator;
    const commentIndicatorResolved = theme.commentIndicatorResolved;
    const trackCellFetches = perf?.enabled === true;
    let cellFetches = 0;

    // Content layer state.
    const layoutEngine = this.textLayoutEngine;
    contentCtx.textBaseline = "alphabetic";
    contentCtx.textAlign = "left";

    let currentTextFill = "";
    let currentGridFill = "";

    const zoom = this.zoom;
    const paddingX = 4 * zoom;
    const paddingY = 2 * zoom;

    // Font specs are part of the text-layout cache key and are returned in layout runs.
    // Avoid mutating a shared object after passing it to the layout engine.
    let fontSpec = { family: "system-ui", sizePx: 12 * zoom, weight: "400", style: "normal" };
    let currentFontFamily = "";
    let currentFontSize = -1;
    let currentFontWeight = "";
    let currentFontStyle = "";

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    const quadrantRange: CellRange = { startRow, endRow, startCol, endCol };
    const mergedRanges = mergedIndex.getRanges();
    const quadrantMergedRanges =
      mergedRanges.length === 0 ? [] : mergedRanges.filter((range) => rangesIntersect(range, quadrantRange));

    const hasMerges = quadrantMergedRanges.length > 0;

    const cellCache = new Map<number, Map<number, CellData | null>>();
    const getCellCached = (row: number, col: number): CellData | null => {
      let rowCache = cellCache.get(row);
      if (!rowCache) {
        rowCache = new Map();
        cellCache.set(row, rowCache);
      }
      if (rowCache.has(col)) return rowCache.get(col) ?? null;
      const cell = this.provider.getCell(row, col);
      if (trackCellFetches) cellFetches += 1;
      rowCache.set(col, cell);
      return cell;
    };

    const blockedCache = new Map<number, Map<number, boolean>>();
    const isBlockedForOverflow = (row: number, col: number): boolean => {
      let rowCache = blockedCache.get(row);
      if (!rowCache) {
        rowCache = new Map();
        blockedCache.set(row, rowCache);
      }
      if (rowCache.has(col)) return rowCache.get(col) ?? false;

      if (mergedRanges.length > 0 && mergedIndex.rangeAt({ row, col })) {
        rowCache.set(col, true);
        return true;
      }

      const cell = getCellCached(row, col);
      const value = cell?.value ?? null;
      const blocked = value !== null && value !== "";
      rowCache.set(col, blocked);
      return blocked;
    };

    const drawCellContent = (options: {
      cell: CellData;
      x: number;
      y: number;
      width: number;
      height: number;
      spanStartRow: number;
      spanEndRow: number;
      spanStartCol: number;
      spanEndCol: number;
      isHeader: boolean;
    }): void => {
      const { cell, x, y, width, height, spanStartRow, spanEndRow, spanStartCol, spanEndCol, isHeader } = options;
      const style = cell.style;

      const richText = cell.richText;
      const richTextText = richText?.text ?? "";
      const hasRichText = Boolean(richText && richTextText);
      const hasValue = cell.value !== null;

      if (hasValue || hasRichText) {
        const fontSize = (style?.fontSize ?? 12) * zoom;
        const fontFamily = style?.fontFamily ?? "system-ui";
        const fontWeight = style?.fontWeight ?? "400";
        const fontStyle = style?.fontStyle ?? "normal";

        if (
          currentFontSize !== fontSize ||
          currentFontFamily !== fontFamily ||
          currentFontWeight !== fontWeight ||
          currentFontStyle !== fontStyle
        ) {
          currentFontSize = fontSize;
          currentFontFamily = fontFamily;
          currentFontWeight = fontWeight;
          currentFontStyle = fontStyle;
          fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle };
          contentCtx.font = toCanvasFontString(fontSpec);
        }

        const explicitColor = style?.color;
        const valueForColor = cell.value !== null ? cell.value : richTextText;
        const fillStyle =
          explicitColor !== undefined
            ? explicitColor
            : typeof valueForColor === "string" && valueForColor.startsWith("#")
              ? errorTextColor
              : isHeader
                ? headerTextColor
                : textColor;
        if (fillStyle !== currentTextFill) {
          contentCtx.fillStyle = fillStyle;
          currentTextFill = fillStyle;
        }

        if (hasRichText) {
          const text = richTextText;
          const wrapMode = style?.wrapMode ?? "none";
          const direction = style?.direction ?? "auto";
          const verticalAlign = style?.verticalAlign ?? "middle";
          const rotationDeg = style?.rotationDeg ?? 0;

          const availableWidth = Math.max(0, width - paddingX * 2);
          const availableHeight = Math.max(0, height - paddingY * 2);

          const align: CanvasTextAlign = style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
          const layoutAlign =
            align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
              ? (align as "left" | "right" | "center" | "start" | "end")
              : "start";

          const hasExplicitNewline = EXPLICIT_NEWLINE_RE.test(text);
          const rotationRad = (rotationDeg * Math.PI) / 180;

          const offsets = buildCodePointIndex(text);
          const textLen = offsets.length - 1;
          const rawRuns =
            Array.isArray(richText?.runs) && richText.runs.length > 0
              ? richText.runs
              : [{ start: 0, end: textLen, style: undefined }];

          const defaults = {
            family: fontFamily,
            sizePx: fontSize,
            weight: fontWeight,
            style: fontStyle
          } satisfies Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number; style: string };

          const layoutRuns = rawRuns.map((run) => {
            const runStyle =
              run.style && typeof run.style === "object" ? (run.style as RichTextRunStyle) : (undefined as RichTextRunStyle | undefined);
            return {
              text: sliceByCodePointRange(text, offsets, run.start, run.end),
              font: fontSpecForRichTextStyle(runStyle, defaults, zoom),
              color: engineColorToCanvasColor(runStyle?.color),
              underline: isUnderline(runStyle)
            };
          });

          const maxFontSizePx = layoutRuns.reduce((acc, run) => Math.max(acc, run.font.sizePx), defaults.sizePx);
          const lineHeight = Math.ceil(maxFontSizePx * 1.2);
          const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

          contentCtx.save();

          if (wrapMode === "none" && !hasExplicitNewline && rotationDeg === 0) {
            // Fast path: single-line rich text (no wrapping). Uses cached measurements.
            const baseDirection = direction === "auto" ? detectBaseDirection(text) : direction;
            const resolvedAlign =
              layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
                ? layoutAlign
                : resolveAlign(layoutAlign, baseDirection);

            const fragments = layoutRuns
              .map((fragment) => {
                if (fragment.text.length === 0) return null;
                const measurement = layoutEngine?.measure(fragment.text, fragment.font);
                const width = measurement?.width ?? contentCtx.measureText(fragment.text).width;
                const ascent = measurement?.ascent ?? fragment.font.sizePx * 0.8;
                const descent = measurement?.descent ?? fragment.font.sizePx * 0.2;
                return {
                  ...fragment,
                  width,
                  ascent,
                  descent
                };
              })
              .filter((fragment): fragment is NonNullable<typeof fragment> => Boolean(fragment));

            const totalWidth = fragments.reduce((acc, fragment) => acc + fragment.width, 0);
            const lineAscent = fragments.reduce((acc, fragment) => Math.max(acc, fragment.ascent), 0);
            const lineDescent = fragments.reduce((acc, fragment) => Math.max(acc, fragment.descent), 0);

            let cursorX = x + paddingX;
            if (resolvedAlign === "center") {
              cursorX = x + paddingX + (availableWidth - totalWidth) / 2;
            } else if (resolvedAlign === "right") {
              cursorX = x + paddingX + (availableWidth - totalWidth);
            }

            let baselineY = y + paddingY + lineAscent;
            if (verticalAlign === "middle") {
              baselineY = y + height / 2 + (lineAscent - lineDescent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + height - paddingY - lineDescent;
            }

            const underlineSegments: Array<{ x1: number; x2: number; y: number; color: string; lineWidth: number }> = [];

            const drawFragments = () => {
              let xCursor = cursorX;
              for (const fragment of fragments) {
                contentCtx.font = toCanvasFontString(fragment.font);
                contentCtx.fillStyle = fragment.color ?? fillStyle;
                contentCtx.fillText(fragment.text, xCursor, baselineY);

                if (fragment.underline) {
                  const underlineOffset = Math.max(1, Math.round(fragment.font.sizePx * 0.08));
                  const underlineY = baselineY + underlineOffset;
                  underlineSegments.push({
                    x1: xCursor,
                    x2: xCursor + fragment.width,
                    y: underlineY,
                    color: contentCtx.fillStyle as string,
                    lineWidth: Math.max(1, Math.round(fragment.font.sizePx / 16))
                  });
                }

                xCursor += fragment.width;
              }
            };

            const shouldClip = totalWidth > availableWidth;
            if (shouldClip) {
              let clipX = x;
              let clipWidth = width;

              if ((resolvedAlign === "left" || resolvedAlign === "right") && totalWidth > width - paddingX) {
                const requiredExtra = paddingX + totalWidth - width;
                if (requiredExtra > 0) {
                  if (resolvedAlign === "left") {
                    let extra = 0;
                    for (
                      let probeCol = spanEndCol, steps = 0;
                      probeCol < colCount && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                      probeCol++, steps++
                    ) {
                      let blocked = false;
                      for (let r = spanStartRow; r < spanEndRow; r++) {
                        if (r < 0 || r >= rowCount) {
                          blocked = true;
                          break;
                        }
                        if (isBlockedForOverflow(r, probeCol)) {
                          blocked = true;
                          break;
                        }
                      }
                      if (blocked) break;
                      extra += colAxis.getSize(probeCol);
                    }
                    clipWidth += extra;
                  } else {
                    let extra = 0;
                    for (
                      let probeCol = spanStartCol - 1, steps = 0;
                      probeCol >= 0 && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                      probeCol--, steps++
                    ) {
                      let blocked = false;
                      for (let r = spanStartRow; r < spanEndRow; r++) {
                        if (r < 0 || r >= rowCount) {
                          blocked = true;
                          break;
                        }
                        if (isBlockedForOverflow(r, probeCol)) {
                          blocked = true;
                          break;
                        }
                      }
                      if (blocked) break;
                      extra += colAxis.getSize(probeCol);
                    }
                    clipX -= extra;
                    clipWidth += extra;
                  }
                }
              }

              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(clipX, y, clipWidth, height);
              contentCtx.clip();
              drawFragments();
              contentCtx.restore();
            } else {
              drawFragments();
            }

            if (underlineSegments.length > 0) {
              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();
              for (const segment of underlineSegments) {
                contentCtx.beginPath();
                contentCtx.strokeStyle = segment.color;
                contentCtx.lineWidth = segment.lineWidth;
                contentCtx.moveTo(segment.x1, segment.y);
                contentCtx.lineTo(segment.x2, segment.y);
                contentCtx.stroke();
              }
              contentCtx.restore();
            }
          } else if (layoutEngine && availableWidth > 0) {
            const layout = layoutEngine.layout({
              runs: layoutRuns.map((r) => ({ text: r.text, font: r.font, color: r.color, underline: r.underline })),
              text: undefined,
              font: defaults,
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
              originY = y + height - paddingY - layout.height;
            }

            const originX = x + paddingX;
            const shouldClip = layout.width > availableWidth || layout.height > availableHeight || rotationDeg !== 0;

            const drawLayout = (collectUnderlines: boolean) => {
              const underlineSegments: Array<{ x1: number; x2: number; y: number; color: string; lineWidth: number }> = [];

              for (let i = 0; i < layout.lines.length; i++) {
                const line = layout.lines[i];
                let xCursor = originX + line.x;
                const baselineY = originY + i * layout.lineHeight + line.ascent;

                for (const run of line.runs as Array<{ text: string; font: FontSpec; color?: string; underline?: boolean }>) {
                  const measurement = layoutEngine.measure(run.text, run.font);
                  contentCtx.font = toCanvasFontString(run.font);
                  contentCtx.fillStyle = run.color ?? fillStyle;
                  contentCtx.fillText(run.text, xCursor, baselineY);

                  if (run.underline) {
                    const underlineOffset = Math.max(1, Math.round(run.font.sizePx * 0.08));
                    const underlineY = baselineY + underlineOffset;
                    const lineWidth = Math.max(1, Math.round(run.font.sizePx / 16));
                    if (collectUnderlines) {
                      underlineSegments.push({
                        x1: xCursor,
                        x2: xCursor + measurement.width,
                        y: underlineY,
                        color: contentCtx.fillStyle as string,
                        lineWidth
                      });
                    } else {
                      contentCtx.beginPath();
                      contentCtx.strokeStyle = contentCtx.fillStyle;
                      contentCtx.lineWidth = lineWidth;
                      contentCtx.moveTo(xCursor, underlineY);
                      contentCtx.lineTo(xCursor + measurement.width, underlineY);
                      contentCtx.stroke();
                    }
                  }

                  xCursor += measurement.width;
                }
              }

              return underlineSegments;
            };

            if (shouldClip) {
              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();

              if (rotationRad) {
                const cx = x + width / 2;
                const cy = y + height / 2;
                contentCtx.translate(cx, cy);
                contentCtx.rotate(rotationRad);
                contentCtx.translate(-cx, -cy);
              }

              drawLayout(false);
              contentCtx.restore();
            } else {
              const underlineSegments = drawLayout(true);
              if (underlineSegments.length > 0) {
                contentCtx.save();
                contentCtx.beginPath();
                contentCtx.rect(x, y, width, height);
                contentCtx.clip();
                for (const segment of underlineSegments) {
                  contentCtx.beginPath();
                  contentCtx.strokeStyle = segment.color;
                  contentCtx.lineWidth = segment.lineWidth;
                  contentCtx.moveTo(segment.x1, segment.y);
                  contentCtx.lineTo(segment.x2, segment.y);
                  contentCtx.stroke();
                }
                contentCtx.restore();
              }
            }
          } else {
            // Fallback: no layout engine available (shouldn't happen in supported environments).
            contentCtx.save();
            contentCtx.beginPath();
            contentCtx.rect(x, y, width, height);
            contentCtx.clip();
            contentCtx.textBaseline = "middle";
            contentCtx.fillStyle = fillStyle;
            contentCtx.font = toCanvasFontString(fontSpec);
            contentCtx.fillText(text, x + paddingX, y + height / 2);
            contentCtx.textBaseline = "alphabetic";
            contentCtx.restore();
          }

          contentCtx.restore();
        } else if (hasValue) {
          const text = formatCellDisplayText(cell.value);
          const wrapMode = style?.wrapMode ?? "none";
          const direction = style?.direction ?? "auto";
          const verticalAlign = style?.verticalAlign ?? "middle";
          const rotationDeg = style?.rotationDeg ?? 0;

          const availableWidth = Math.max(0, width - paddingX * 2);
          const availableHeight = Math.max(0, height - paddingY * 2);
          const lineHeight = Math.ceil(fontSize * 1.2);
          const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

          const align: CanvasTextAlign = style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
          const layoutAlign =
            align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
              ? (align as "left" | "right" | "center" | "start" | "end")
              : "start";

          const hasExplicitNewline = EXPLICIT_NEWLINE_RE.test(text);
          const rotationRad = (rotationDeg * Math.PI) / 180;

          if (wrapMode === "none" && !hasExplicitNewline && rotationDeg === 0) {
            const resolvedAlign =
              layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
                ? layoutAlign
                : resolveAlign(
                    layoutAlign,
                    direction === "auto" ? (typeof cell.value === "string" ? detectBaseDirection(text) : "ltr") : direction
                  );

            const measurement = layoutEngine?.measure(text, fontSpec);
            const textWidth = measurement?.width ?? contentCtx.measureText(text).width;
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
              baselineY = y + height / 2 + (ascent - descent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + height - paddingY - descent;
            }

            const shouldClip = textWidth > availableWidth;
            if (shouldClip) {
              let clipX = x;
              let clipWidth = width;

              if ((resolvedAlign === "left" || resolvedAlign === "right") && textWidth > width - paddingX) {
                const requiredExtra = paddingX + textWidth - width;
                if (requiredExtra > 0) {
                  if (resolvedAlign === "left") {
                    let extra = 0;
                    for (
                      let probeCol = spanEndCol, steps = 0;
                      probeCol < colCount && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                      probeCol++, steps++
                    ) {
                      let blocked = false;
                      for (let r = spanStartRow; r < spanEndRow; r++) {
                        if (r < 0 || r >= rowCount) {
                          blocked = true;
                          break;
                        }
                        if (isBlockedForOverflow(r, probeCol)) {
                          blocked = true;
                          break;
                        }
                      }
                      if (blocked) break;
                      extra += colAxis.getSize(probeCol);
                    }
                    clipWidth += extra;
                  } else {
                    let extra = 0;
                    for (
                      let probeCol = spanStartCol - 1, steps = 0;
                      probeCol >= 0 && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                      probeCol--, steps++
                    ) {
                      let blocked = false;
                      for (let r = spanStartRow; r < spanEndRow; r++) {
                        if (r < 0 || r >= rowCount) {
                          blocked = true;
                          break;
                        }
                        if (isBlockedForOverflow(r, probeCol)) {
                          blocked = true;
                          break;
                        }
                      }
                      if (blocked) break;
                      extra += colAxis.getSize(probeCol);
                    }
                    clipX -= extra;
                    clipWidth += extra;
                  }
                }
              }

              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(clipX, y, clipWidth, height);
              contentCtx.clip();
              contentCtx.fillText(text, textX, baselineY);
              contentCtx.restore();
            } else {
              contentCtx.fillText(text, textX, baselineY);
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
              originY = y + height - paddingY - layout.height;
            }

            const originX = x + paddingX;
            const shouldClip = layout.width > availableWidth || layout.height > availableHeight || rotationDeg !== 0;

            if (shouldClip) {
              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();

              if (rotationRad) {
                const cx = x + width / 2;
                const cy = y + height / 2;
                contentCtx.translate(cx, cy);
                contentCtx.rotate(rotationRad);
                contentCtx.translate(-cx, -cy);
              }

              drawTextLayout(contentCtx, layout, originX, originY);
              contentCtx.restore();
            } else {
              drawTextLayout(contentCtx, layout, originX, originY);
            }
          } else {
            contentCtx.save();
            contentCtx.beginPath();
            contentCtx.rect(x, y, width, height);
            contentCtx.clip();
            contentCtx.textBaseline = "middle";
            contentCtx.fillText(text, x + paddingX, y + height / 2);
            contentCtx.textBaseline = "alphabetic";
            contentCtx.restore();
          }
        }
      }

      if (cell.comment) {
        const resolved = cell.comment.resolved ?? false;
        const maxSize = Math.min(width, height);
        const size = Math.min(maxSize, Math.max(6, maxSize * 0.25));
        if (size > 0) {
          contentCtx.save();
          contentCtx.beginPath();
          contentCtx.moveTo(x + width, y);
          contentCtx.lineTo(x + width - size, y);
          contentCtx.lineTo(x + width, y + size);
          contentCtx.closePath();
          contentCtx.fillStyle = resolved ? commentIndicatorResolved : commentIndicator;
          contentCtx.fill();
          contentCtx.restore();
        }
      }
    };

    // Render merged regions (fill + text) first so we can skip their constituent cells below.
    for (const range of quadrantMergedRanges) {
      const anchorRow = range.startRow;
      const anchorCol = range.startCol;
      const anchorCell = getCellCached(anchorRow, anchorCol);

      const x1Sheet = colAxis.positionOf(range.startCol);
      const x2Sheet = colAxis.positionOf(range.endCol);
      const y1Sheet = rowAxis.positionOf(range.startRow);
      const y2Sheet = rowAxis.positionOf(range.endRow);

      const x = x1Sheet - quadrant.scrollBaseX + quadrant.originX;
      const y = y1Sheet - quadrant.scrollBaseY + quadrant.originY;
      const width = x2Sheet - x1Sheet;
      const height = y2Sheet - y1Sheet;
      if (width <= 0 || height <= 0) continue;

      const anchorStyle = anchorCell?.style;
      const isHeader = anchorRow < headerRows || anchorCol < headerCols;

      const fill = anchorStyle?.fill ?? (isHeader ? headerBg : undefined);
      const fillToDraw = fill && fill !== gridBg ? fill : null;
      if (fillToDraw) {
        if (fillToDraw !== currentGridFill) {
          gridCtx.fillStyle = fillToDraw;
          currentGridFill = fillToDraw;
        }
        gridCtx.fillRect(x, y, width, height);
      }

      if (anchorCell) {
        drawCellContent({
          cell: anchorCell,
          x,
          y,
          width,
          height,
          spanStartRow: range.startRow,
          spanEndRow: range.endRow,
          spanStartCol: range.startCol,
          spanEndCol: range.endCol,
          isHeader
        });
      }
    }

    const startColXSheet = colAxis.positionOf(startCol);
    const startRowYSheet = rowAxis.positionOf(startRow);
    let rowYSheet = startRowYSheet;
    for (let row = startRow; row < endRow; row++) {
      const rowHeight = rowAxis.getSize(row);
      const y = rowYSheet - quadrant.scrollBaseY + quadrant.originY;

      // Batch contiguous fills (per row) to cut down on `fillRect` calls for the
      // common case where large runs share the same background color.
      let fillRunColor: string | null = null;
      let fillRunX = 0;
      let fillRunWidth = 0;

      let colXSheet = startColXSheet;
      for (let col = startCol; col < endCol; col++) {
        const colWidth = colAxis.getSize(col);
        const x = colXSheet - quadrant.scrollBaseX + quadrant.originX;

        if (hasMerges && mergedIndex.rangeAt({ row, col })) {
          if (fillRunColor && fillRunWidth > 0) {
            if (fillRunColor !== currentGridFill) {
              gridCtx.fillStyle = fillRunColor;
              currentGridFill = fillRunColor;
            }
            gridCtx.fillRect(fillRunX, y, fillRunWidth, rowHeight);
          }
          fillRunColor = null;
          fillRunWidth = 0;
          colXSheet += colWidth;
          continue;
        }

        const cell = getCellCached(row, col);
        const style = cell?.style;
        const isHeader = row < headerRows || col < headerCols;

        // Background fill (grid layer).
        const fill = style?.fill ?? (isHeader ? headerBg : undefined);
        const fillToDraw = fill && fill !== gridBg ? fill : null;
        if (fillToDraw) {
          if (fillToDraw !== fillRunColor) {
            if (fillRunColor && fillRunWidth > 0) {
              if (fillRunColor !== currentGridFill) {
                gridCtx.fillStyle = fillRunColor;
                currentGridFill = fillRunColor;
              }
              gridCtx.fillRect(fillRunX, y, fillRunWidth, rowHeight);
            }
            fillRunColor = fillToDraw;
            fillRunX = x;
            fillRunWidth = colWidth;
          } else {
            fillRunWidth += colWidth;
          }
        } else if (fillRunColor) {
          if (fillRunWidth > 0) {
            if (fillRunColor !== currentGridFill) {
              gridCtx.fillStyle = fillRunColor;
              currentGridFill = fillRunColor;
            }
            gridCtx.fillRect(fillRunX, y, fillRunWidth, rowHeight);
          }
          fillRunColor = null;
          fillRunWidth = 0;
        }

        if (cell) {
          drawCellContent({
            cell,
            x,
            y,
            width: colWidth,
            height: rowHeight,
            spanStartRow: row,
            spanEndRow: row + 1,
            spanStartCol: col,
            spanEndCol: col + 1,
            isHeader
          });
        }

        colXSheet += colWidth;
      }

      if (fillRunColor && fillRunWidth > 0) {
        if (fillRunColor !== currentGridFill) {
          gridCtx.fillStyle = fillRunColor;
          currentGridFill = fillRunColor;
        }
        gridCtx.fillRect(fillRunX, y, fillRunWidth, rowHeight);
      }
      rowYSheet += rowHeight;
    }

    if (trackCellFetches) {
      perf!.cellFetches += cellFetches;
    }

    // Gridlines (grid layer), drawn after fills.
    gridCtx.strokeStyle = theme.gridLine;
    gridCtx.lineWidth = 1;

    if (!hasMerges) {
      const xStart = startColXSheet - quadrant.scrollBaseX + quadrant.originX;
      const yStart = startRowYSheet - quadrant.scrollBaseY + quadrant.originY;
      const yEnd = rowYSheet - quadrant.scrollBaseY + quadrant.originY;

      gridCtx.beginPath();
      let xSheet = startColXSheet;
      for (let col = startCol; col <= endCol; col++) {
        const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
        const cx = crispLine(x);
        gridCtx.moveTo(cx, yStart);
        gridCtx.lineTo(cx, yEnd);
        if (col < endCol) xSheet += colAxis.getSize(col);
      }

      const xEnd = xSheet - quadrant.scrollBaseX + quadrant.originX;

      let ySheet = startRowYSheet;
      for (let row = startRow; row <= endRow; row++) {
        const yy = ySheet - quadrant.scrollBaseY + quadrant.originY;
        const cy = crispLine(yy);
        gridCtx.moveTo(xStart, cy);
        gridCtx.lineTo(xEnd, cy);
        if (row < endRow) ySheet += rowAxis.getSize(row);
      }

      gridCtx.stroke();
      return;
    }

    gridCtx.beginPath();
    rowYSheet = startRowYSheet;
    for (let row = startRow; row < endRow; row++) {
      const rowHeight = rowAxis.getSize(row);
      const y = rowYSheet - quadrant.scrollBaseY + quadrant.originY;
      const yNext = y + rowHeight;
      const cyTop = row === startRow ? crispLine(y) : 0;
      const cyBottom = crispLine(yNext);

      let colXSheet = startColXSheet;
      const xLeft = colXSheet - quadrant.scrollBaseX + quadrant.originX;
      const cxLeft = crispLine(xLeft);
      if (!isInteriorVerticalGridline(mergedIndex, row, startCol - 1)) {
        gridCtx.moveTo(cxLeft, y);
        gridCtx.lineTo(cxLeft, yNext);
      }

      for (let col = startCol; col < endCol; col++) {
        const colWidth = colAxis.getSize(col);
        const x = colXSheet - quadrant.scrollBaseX + quadrant.originX;
        const xNext = x + colWidth;

        if (row === startRow && !isInteriorHorizontalGridline(mergedIndex, startRow - 1, col)) {
          gridCtx.moveTo(x, cyTop);
          gridCtx.lineTo(xNext, cyTop);
        }

        if (!isInteriorHorizontalGridline(mergedIndex, row, col)) {
          gridCtx.moveTo(x, cyBottom);
          gridCtx.lineTo(xNext, cyBottom);
        }

        if (!isInteriorVerticalGridline(mergedIndex, row, col)) {
          const cx = crispLine(xNext);
          gridCtx.moveTo(cx, y);
          gridCtx.lineTo(cx, yNext);
        }

        colXSheet += colWidth;
      }

      rowYSheet += rowHeight;
    }

    gridCtx.stroke();
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
        this.renderSelectionQuadrant(intersection, viewport, this.getMergedIndex(viewport));
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

    ctx.strokeStyle = this.theme.gridLine;
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
    const zoom = this.zoom;
    const paddingX = 4 * zoom;
    const paddingY = 2 * zoom;

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
          const fontSize = (style?.fontSize ?? 12) * zoom;
          const fontFamily = style?.fontFamily ?? "system-ui";
          const fontWeight = style?.fontWeight ?? "400";
          const fontStyle = style?.fontStyle ?? "normal";
          const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle };
          const font = toCanvasFontString(fontSpec);

          if (font !== currentFont) {
            ctx.font = font;
            currentFont = font;
          }

          const fillStyle = resolveCellTextColorWithTheme(cell.value, style?.color, this.theme);
          if (fillStyle !== currentFillStyle) {
            ctx.fillStyle = fillStyle;
            currentFillStyle = fillStyle;
          }

          const text = formatCellDisplayText(cell.value);

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
            ctx.fillStyle = resolved ? this.theme.commentIndicatorResolved : this.theme.commentIndicator;
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
    viewport: GridViewportState,
    mergedIndex: MergedCellIndex
  ): void {
    const ctx = this.selectionCtx;
    if (!ctx) return;

    const transientRange = this.rangeSelection;
    if (transientRange) {
      const rects = this.rangeToViewportRects(transientRange, viewport);

      ctx.fillStyle = this.theme.selectionFill;
      for (const rect of rects) {
        const clipped = intersectRect(rect, intersection);
        if (!clipped) continue;
        ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
      }

      ctx.strokeStyle = this.theme.selectionBorder;
      ctx.lineWidth = 2;
      for (const rect of rects) {
        if (!intersectRect(rect, intersection)) continue;
        if (rect.width <= 2 || rect.height <= 2) continue;
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
      }
    }

    const fillPreview = this.fillPreviewRange;
    if (fillPreview) {
      const rects = this.rangeToViewportRects(fillPreview, viewport);

      if (rects.length > 0) {
        ctx.save();

        ctx.fillStyle = this.theme.selectionFill;
        ctx.globalAlpha = 0.6;
        for (const rect of rects) {
          const clipped = intersectRect(rect, intersection);
          if (!clipped) continue;
          ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
        }

        ctx.strokeStyle = this.theme.selectionBorder;
        ctx.globalAlpha = 1;
        ctx.lineWidth = 2;
        ctx.setLineDash([5, 4]);
        for (const rect of rects) {
          if (!intersectRect(rect, intersection)) continue;
          if (rect.width <= 2 || rect.height <= 2) continue;
          ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
        }

        ctx.restore();
      }
    }

    // Primary selection overlays (normal grid selection). These intentionally
    // render *below* formula reference highlights so references remain visible
    // during formula range selection UX.
    const selectionRanges = this.selectionRanges;
    if (selectionRanges.length > 0) {
      const activeIndex = this.activeSelectionIndex;
      const activeRange = selectionRanges[activeIndex];

      const drawRange = (range: CellRange, options: { fillAlpha: number; strokeAlpha: number; strokeWidth: number }) => {
        const rects = this.rangeToViewportRects(range, viewport);
        if (rects.length === 0) return;

        ctx.save();

        ctx.fillStyle = this.theme.selectionFill;
        ctx.globalAlpha = options.fillAlpha;
        for (const rect of rects) {
          const clipped = intersectRect(rect, intersection);
          if (!clipped) continue;
          ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
        }

        ctx.strokeStyle = this.theme.selectionBorder;
        ctx.globalAlpha = options.strokeAlpha;
        ctx.lineWidth = options.strokeWidth;
        const inset = options.strokeWidth / 2;
        for (const rect of rects) {
          if (!intersectRect(rect, intersection)) continue;
          if (rect.width <= options.strokeWidth || rect.height <= options.strokeWidth) continue;
          ctx.strokeRect(rect.x + inset, rect.y + inset, rect.width - options.strokeWidth, rect.height - options.strokeWidth);
        }

        ctx.restore();
      };

      for (let i = 0; i < selectionRanges.length; i++) {
        if (i === activeIndex) continue;
        drawRange(selectionRanges[i], { fillAlpha: 2 / 3, strokeAlpha: 0.8, strokeWidth: 1 });
      }

      drawRange(activeRange, { fillAlpha: 1, strokeAlpha: 1, strokeWidth: 2 });

      if (this.fillHandleEnabled) {
        const handleSize = 8 * this.zoom;
        const handleRow = activeRange.endRow - 1;
        const handleCol = activeRange.endCol - 1;
        const handleCellRect = this.cellRectInViewport(handleRow, handleCol, viewport, { clampToViewport: false });
        if (handleCellRect && handleCellRect.width >= handleSize && handleCellRect.height >= handleSize) {
          const handleRect: Rect = {
            x: handleCellRect.x + handleCellRect.width - handleSize / 2,
            y: handleCellRect.y + handleCellRect.height - handleSize / 2,
            width: handleSize,
            height: handleSize
          };
          const handleClipped = intersectRect(handleRect, intersection);
          if (handleClipped) {
            ctx.fillStyle = this.theme.selectionHandle;
            ctx.fillRect(handleClipped.x, handleClipped.y, handleClipped.width, handleClipped.height);
          }
        }
      }

      const activeCell = this.selection;
      const isSingleCell = activeRange.endRow - activeRange.startRow <= 1 && activeRange.endCol - activeRange.startCol <= 1;
      if (activeCell && !isSingleCell) {
        const merged = mergedIndex.rangeAt(activeCell);
        const activeCellRange = merged ?? {
          startRow: activeCell.row,
          endRow: activeCell.row + 1,
          startCol: activeCell.col,
          endCol: activeCell.col + 1
        };

        const activeRects = this.rangeToViewportRects(activeCellRange, viewport);
        if (activeRects.length > 0) {
          ctx.strokeStyle = this.theme.selectionBorder;
          ctx.lineWidth = 2;
          for (const rect of activeRects) {
            if (!intersectRect(rect, intersection)) continue;
            if (rect.width <= 2 || rect.height <= 2) continue;
            ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
          }
        }
      }
    }

    // Formula reference highlights: render on top of all selection overlays so
    // colored outlines remain visible while editing formulas.
    if (this.referenceHighlights.length > 0) {
      const strokeRects = (rects: Rect[], lineWidth: number) => {
        const inset = lineWidth / 2;
        for (const rect of rects) {
          if (!intersectRect(rect, intersection)) continue;
          if (rect.width <= lineWidth || rect.height <= lineWidth) continue;
          ctx.strokeRect(rect.x + inset, rect.y + inset, rect.width - lineWidth, rect.height - lineWidth);
        }
      };

      ctx.save();

      ctx.lineWidth = 2;
      ctx.setLineDash([4, 3]);

      for (const highlight of this.referenceHighlights) {
        if (highlight.active) continue;
        const rects = this.rangeToViewportRects(highlight.range, viewport);
        if (rects.length === 0) continue;
        ctx.strokeStyle = highlight.color;
        strokeRects(rects, ctx.lineWidth);
      }

      const hasActive = this.referenceHighlights.some((h) => h.active);
      if (hasActive) {
        ctx.lineWidth = 3;
        ctx.setLineDash([]);
        for (const highlight of this.referenceHighlights) {
          if (!highlight.active) continue;
          const rects = this.rangeToViewportRects(highlight.range, viewport);
          if (rects.length === 0) continue;
          ctx.strokeStyle = highlight.color;
          strokeRects(rects, ctx.lineWidth);
        }
      }

      ctx.restore();
    }
  }

  private renderRemotePresenceOverlays(
    ctx: CanvasRenderingContext2D,
    viewport: GridViewportState,
    mergedIndex: MergedCellIndex
  ): void {
    if (this.remotePresences.length === 0) return;

    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    ctx.save();
    ctx.font = this.presenceFont;
    ctx.textBaseline = "top";

    const selectionFillAlpha = 0.12;
    const selectionStrokeAlpha = 0.9;
    const zoom = this.zoom;
    const cursorStrokeWidth = 2 * zoom;
    const badgePaddingX = 6 * zoom;
    const badgePaddingY = 3 * zoom;
    const badgeOffsetX = 8 * zoom;
    const badgeOffsetY = -18 * zoom;
    const badgeTextHeight = 14 * zoom;
    const cursorInset = cursorStrokeWidth / 2;

    for (const presence of this.remotePresences) {
      const color = presence.color ?? this.theme.remotePresenceDefault;

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
            if (rect.width <= cursorStrokeWidth || rect.height <= cursorStrokeWidth) continue;
            ctx.strokeRect(
              rect.x + cursorInset,
              rect.y + cursorInset,
              rect.width - cursorStrokeWidth,
              rect.height - cursorStrokeWidth
            );
          }

          ctx.globalAlpha = 1;
        }
      }

      if (presence.cursor) {
        const cursorCellRect = this.cellRectInViewport(presence.cursor.row, presence.cursor.col, viewport);
        if (!cursorCellRect) continue;

        const cursorRange = mergedIndex.rangeAt(presence.cursor) ?? {
          startRow: presence.cursor.row,
          endRow: presence.cursor.row + 1,
          startCol: presence.cursor.col,
          endCol: presence.cursor.col + 1
        };
        const cursorRects = this.rangeToViewportRects(cursorRange, viewport);
        if (cursorRects.length === 0) continue;

        ctx.globalAlpha = 1;
        ctx.strokeStyle = color;
        ctx.lineWidth = cursorStrokeWidth;
        for (const rect of cursorRects) {
          if (rect.width <= cursorStrokeWidth || rect.height <= cursorStrokeWidth) continue;
          ctx.strokeRect(
            rect.x + cursorInset,
            rect.y + cursorInset,
            rect.width - cursorStrokeWidth,
            rect.height - cursorStrokeWidth
          );
        }

        const name = presence.name ?? "Anonymous";
        const metricsKey = `${this.presenceFont}::${name}`;
        let textWidth = this.textWidthCache.get(metricsKey);
        if (textWidth === undefined) {
          textWidth = ctx.measureText(name).width;
          this.textWidthCache.set(metricsKey, textWidth);
        }

        const badgeWidth = textWidth + badgePaddingX * 2;
        const badgeHeight = badgeTextHeight + badgePaddingY * 2;
        const badgeX = cursorCellRect.x + cursorCellRect.width + badgeOffsetX;
        const badgeY = cursorCellRect.y + badgeOffsetY;

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

    ctx.strokeStyle = this.theme.freezeLine;
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

  private cellRectInViewport(
    row: number,
    col: number,
    viewport: GridViewportState,
    options?: { clampToViewport?: boolean }
  ): Rect | null {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const colX = colAxis.positionOf(col);
    const rowY = rowAxis.positionOf(row);
    const width = colAxis.getSize(col);
    const height = rowAxis.getSize(row);

    let x: number;
    let y: number;

    const scrollCols = col >= viewport.frozenCols;
    const scrollRows = row >= viewport.frozenRows;
    x = scrollCols ? colX - viewport.scrollX : colX;
    y = scrollRows ? rowY - viewport.scrollY : rowY;

    const rect = { x, y, width, height };

    if (options?.clampToViewport === false) return rect;

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const quadrantRect: Rect = {
      x: scrollCols ? frozenWidthClamped : 0,
      y: scrollRows ? frozenHeightClamped : 0,
      width: scrollCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped,
      height: scrollRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped
    };

    return intersectRect(rect, quadrantRect);
  }

  private markSelectionDirty(): void {
    const viewport = this.scroll.getViewportState();
    this.dirty.selection.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    this.requestRender();
  }

  private fillPreviewOverlayRects(range: CellRange | null, viewport: GridViewportState): Rect[] {
    if (!range) return [];
    return this.rangeToViewportRects(range, viewport);
  }

  private fillHandleRectInViewport(range: CellRange, viewport: GridViewportState): Rect | null {
    const handleSize = 8 * this.zoom;
    const handleRow = range.endRow - 1;
    const handleCol = range.endCol - 1;
    const handleCellRect = this.cellRectInViewport(handleRow, handleCol, viewport, { clampToViewport: false });
    if (!handleCellRect) return null;
    if (handleCellRect.width < handleSize || handleCellRect.height < handleSize) return null;

    const handleRect: Rect = {
      x: handleCellRect.x + handleCellRect.width - handleSize / 2,
      y: handleCellRect.y + handleCellRect.height - handleSize / 2,
      width: handleSize,
      height: handleSize
    };

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const quadrantRect: Rect = {
      x: handleCol >= viewport.frozenCols ? frozenWidthClamped : 0,
      y: handleRow >= viewport.frozenRows ? frozenHeightClamped : 0,
      width:
        handleCol >= viewport.frozenCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped,
      height:
        handleRow >= viewport.frozenRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped
    };

    return intersectRect(handleRect, quadrantRect);
  }

  private markAllDirtyForThemeChange(): void {
    const viewport = this.scroll.getViewportState();
    const full = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    this.dirty.background.markDirty(full);
    this.dirty.content.markDirty(full);
    this.dirty.selection.markDirty(full);
    this.forceFullRedraw = true;
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

  private getMergedRangeForCell(row: number, col: number): CellRange | null {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const provider = this.provider;

    if (provider.getMergedRangeAt) {
      const merged = provider.getMergedRangeAt(row, col);
      const normalized = merged ? this.normalizeSelectionRange(merged) : null;
      if (!normalized) return null;
      if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) return null;
      return normalized;
    }

    if (provider.getMergedRangesInRange) {
      const candidates = provider.getMergedRangesInRange({ startRow: row, endRow: row + 1, startCol: col, endCol: col + 1 });
      for (const candidate of candidates) {
        const normalized = this.normalizeSelectionRange(candidate);
        if (!normalized) continue;
        if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) continue;
        if (row < normalized.startRow || row >= normalized.endRow) continue;
        if (col < normalized.startCol || col >= normalized.endCol) continue;
        return normalized;
      }
    }

    return null;
  }

  private getMergedAnchorForCell(row: number, col: number): Selection | null {
    const merged = this.getMergedRangeForCell(row, col);
    if (!merged) return null;
    return { row: merged.startRow, col: merged.startCol };
  }

  private expandRangeToMergedCells(range: CellRange): CellRange {
    const provider = this.provider;
    if (!provider.getMergedRangeAt && !provider.getMergedRangesInRange) return range;

    let expanded: CellRange = range;
    for (let iter = 0; iter < 100; iter++) {
      let changed = false;

      const merges = new Map<string, CellRange>();
      const addMerge = (candidate: CellRange | null) => {
        const normalized = candidate ? this.normalizeSelectionRange(candidate) : null;
        if (!normalized) return;
        if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) return;
        const key = `${normalized.startRow},${normalized.endRow},${normalized.startCol},${normalized.endCol}`;
        merges.set(key, normalized);
      };

      if (provider.getMergedRangesInRange) {
        for (const candidate of provider.getMergedRangesInRange(expanded)) {
          addMerge(candidate);
        }
      } else if (provider.getMergedRangeAt) {
        const startRow = expanded.startRow;
        const endRow = expanded.endRow;
        const startCol = expanded.startCol;
        const endCol = expanded.endCol;

        const lastRow = endRow - 1;
        const lastCol = endCol - 1;

        for (let row = startRow; row < endRow; row++) {
          addMerge(provider.getMergedRangeAt(row, startCol));
          addMerge(provider.getMergedRangeAt(row, lastCol));
        }
        for (let col = startCol; col < endCol; col++) {
          addMerge(provider.getMergedRangeAt(startRow, col));
          addMerge(provider.getMergedRangeAt(lastRow, col));
        }
      }

      for (const merge of merges.values()) {
        if (!rangesIntersect(merge, expanded)) continue;
        const next: CellRange = {
          startRow: Math.min(expanded.startRow, merge.startRow),
          endRow: Math.max(expanded.endRow, merge.endRow),
          startCol: Math.min(expanded.startCol, merge.startCol),
          endCol: Math.max(expanded.endCol, merge.endCol)
        };
        if (!isSameCellRange(next, expanded)) {
          expanded = next;
          changed = true;
        }
      }

      if (!changed) break;
    }

    return expanded;
  }

  private getMergedQueryRanges(viewport: GridViewportState): CellRange[] {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (viewport.width <= 0 || viewport.height <= 0) return [];

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

    const ranges: CellRange[] = [];
    const pushRange = (candidate: CellRange) => {
      const normalized = this.normalizeSelectionRange(candidate);
      if (!normalized) return;
      if (normalized.startRow >= rowCount || normalized.startCol >= colCount) return;
      ranges.push(normalized);
    };

    pushRange({
      startRow: frozenRowsRange.start,
      endRow: frozenRowsRange.end,
      startCol: frozenColsRange.start,
      endCol: frozenColsRange.end
    });

    pushRange({
      startRow: frozenRowsRange.start,
      endRow: frozenRowsRange.end,
      startCol: mainCols.start,
      endCol: mainCols.end
    });

    pushRange({
      startRow: mainRows.start,
      endRow: mainRows.end,
      startCol: frozenColsRange.start,
      endCol: frozenColsRange.end
    });

    pushRange({
      startRow: mainRows.start,
      endRow: mainRows.end,
      startCol: mainCols.start,
      endCol: mainCols.end
    });

    return ranges;
  }

  private getMergedIndex(viewport: GridViewportState): MergedCellIndex {
    const provider = this.provider;
    if (!provider.getMergedRangeAt && !provider.getMergedRangesInRange) {
      this.mergedIndex = EMPTY_MERGED_INDEX;
      this.mergedIndexKey = null;
      this.mergedIndexDirty = false;
      return this.mergedIndex;
    }

    const queryRanges = this.getMergedQueryRanges(viewport);
    const key = queryRanges.map((range) => `${range.startRow},${range.endRow},${range.startCol},${range.endCol}`).join("|");

    if (!this.mergedIndexDirty && this.mergedIndexKey === key) {
      return this.mergedIndex;
    }

    const merges = new Map<string, CellRange>();
    const addMerge = (candidate: CellRange | null) => {
      const normalized = candidate ? this.normalizeSelectionRange(candidate) : null;
      if (!normalized) return;
      if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) return;
      const mergeKey = `${normalized.startRow},${normalized.endRow},${normalized.startCol},${normalized.endCol}`;
      merges.set(mergeKey, normalized);
    };

    if (provider.getMergedRangesInRange) {
      for (const range of queryRanges) {
        for (const merged of provider.getMergedRangesInRange(range)) {
          addMerge(merged);
        }
      }
    } else if (provider.getMergedRangeAt) {
      for (const range of queryRanges) {
        for (let row = range.startRow; row < range.endRow; row++) {
          for (let col = range.startCol; col < range.endCol; col++) {
            const merged = provider.getMergedRangeAt(row, col);
            const normalized = merged ? this.normalizeSelectionRange(merged) : null;
            if (!normalized) continue;
            if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) continue;
            addMerge(normalized);
            if (col + 1 < normalized.endCol) {
              col = normalized.endCol - 1;
            }
          }
        }
      }
    }

    this.mergedIndex = new MergedCellIndex([...merges.values()]);
    this.mergedIndexKey = key;
    this.mergedIndexDirty = false;
    return this.mergedIndex;
  }

  private getRowCount(): number {
    return this.scroll.getCounts().rowCount;
  }

  private getColCount(): number {
    return this.scroll.getCounts().colCount;
  }

  private assertRowIndex(row: number): void {
    const rowCount = this.getRowCount();
    if (!Number.isSafeInteger(row) || row < 0 || row >= rowCount) {
      throw new Error(`row must be a safe integer in [0, ${Math.max(0, rowCount - 1)}], got ${row}`);
    }
  }

  private assertColIndex(col: number): void {
    const colCount = this.getColCount();
    if (!Number.isSafeInteger(col) || col < 0 || col >= colCount) {
      throw new Error(`col must be a safe integer in [0, ${Math.max(0, colCount - 1)}], got ${col}`);
    }
  }

  private onAxisSizeChanged(): void {
    const before = this.scroll.getScroll();
    this.scroll.setScroll(before.x, before.y);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);
    this.markAllDirty();
  }
}
