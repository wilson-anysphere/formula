import type { CellBorderSpec, CellData, CellProvider, CellProviderUpdate, CellRange } from "../model/CellProvider.ts";
import { DirtyRegionTracker, type Rect } from "./DirtyRegionTracker.ts";
import { setupHiDpiCanvas } from "./HiDpiCanvas.ts";
import { LruCache } from "../utils/LruCache.ts";
import { DEFAULT_FILL_HANDLE_SIZE_PX } from "../interaction/fillHandle.ts";
import { clampZoom as clampGridZoom } from "../utils/zoomMath.ts";
import type { GridPresence } from "../presence/types.ts";
import type { GridTheme } from "../theme/GridTheme.ts";
import { DEFAULT_GRID_THEME, gridThemesEqual, resolveGridTheme } from "../theme/GridTheme.ts";
import { DEFAULT_GRID_FONT_FAMILY } from "./defaultFontFamilies.ts";
import { MAX_PNG_DIMENSION, MAX_PNG_PIXELS, readImageDimensions } from "./pngDimensions.ts";
import type { GridViewportState } from "../virtualization/VirtualScrollManager.ts";
import { VirtualScrollManager } from "../virtualization/VirtualScrollManager.ts";
import { alignScrollToDevicePixels as alignScrollToDevicePixelsUtil } from "../virtualization/alignScrollToDevicePixels.ts";
import {
  MergedCellIndex,
  isInteriorHorizontalGridline,
  isInteriorVerticalGridline,
  rangesIntersect,
  type IndexedRowRange
} from "./mergedCells.ts";
import {
  TextLayoutEngine,
  createCanvasTextMeasurer,
  detectBaseDirection,
  drawTextLayout,
  resolveAlign,
  toCanvasFontString
} from "@formula/text-layout";

type Layer = "background" | "content" | "selection";

export type CanvasGridImageSource =
  | ImageBitmap
  | Blob
  | ArrayBuffer
  | Uint8Array
  | ImageData
  | HTMLImageElement
  | HTMLCanvasElement
  | OffscreenCanvas;

/**
 * Resolves an `imageId` (from {@link CellData.image}) into bytes or a decoded image object.
 *
 * Returning `null` indicates the image is missing/unavailable and causes the renderer to draw
 * a placeholder.
 */
export type CanvasGridImageResolver = (imageId: string) => Promise<CanvasGridImageSource | null>;

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

export interface CanvasGridRendererOptions {
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
  /**
   * Optional resolver for images referenced by {@link CellData.image}.
   *
   * When provided, the renderer will cache decoded bitmaps keyed by image id.
   */
  imageResolver?: CanvasGridImageResolver | null;
  /**
   * Maximum number of decoded image entries to retain.
   *
   * This bounds memory usage when scrolling through sheets with many distinct images.
   */
  imageBitmapCacheMaxEntries?: number;
  /**
   * Default font family used when rendering/measuring cell text if `CellStyle.fontFamily` is unset.
   *
   * Defaults to the renderer's system UI font stack.
   */
  defaultCellFontFamily?: string;
  /**
   * Default font family used when rendering/measuring header cell text if `CellStyle.fontFamily` is unset.
   *
   * Defaults to `defaultCellFontFamily` when unset.
   */
  defaultHeaderFontFamily?: string;
  /**
   * Optional tiled background image rendered behind cell fills/gridlines.
   *
   * Callers can update this later via `setBackgroundPatternImage()`.
   */
  backgroundPatternImage?: CanvasImageSource | null;
}

export type GridViewportChangeReason = "axisSize" | "frozen" | "resize" | "zoom";

export interface GridViewportChangeEvent {
  viewport: GridViewportState;
  reason: GridViewportChangeReason;
}

export type GridViewportChangeListener = (event: GridViewportChangeEvent) => void;

export interface GridViewportSubscriptionOptions {
  /**
   * Batch viewport change notifications to the next animation frame.
   *
   * This is useful to avoid redundant work during resize drags (window resize, axis resize, etc).
   */
  animationFrame?: boolean;
  /**
   * Debounce viewport change notifications by the given delay (in ms).
   *
   * When set, this overrides `animationFrame`.
   */
  debounceMs?: number;
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

function rectsOverlap(a: Rect, b: Rect): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
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

function crispStrokePosition(pos: number, lineWidth: number): number {
  const roundedPos = Math.round(pos);
  const roundedWidth = Math.round(lineWidth);
  // For odd integer line widths, offset by 0.5 to land on pixel boundaries.
  if (roundedWidth % 2 === 1) return roundedPos + 0.5;
  return roundedPos;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return clamp(Math.trunc(value), min, max);
}

function ensureSansSerifFallback(fontFamily: string): string {
  // Many of our default font stacks already include a `sans-serif` fallback. Avoid generating
  // duplicated values like `..., sans-serif, sans-serif` which cause snapshot churn and make
  // debugging font issues harder.
  if (fontFamily.toLowerCase().includes("sans-serif")) return fontFamily;
  return `${fontFamily}, sans-serif`;
}

function resolveTextIndentPx(textIndentPx: number | undefined, zoom: number): number {
  if (typeof textIndentPx !== "number" || !Number.isFinite(textIndentPx) || textIndentPx <= 0) return 0;
  return textIndentPx * zoom;
}

function getCanvasImageSourceDimensions(source: CanvasImageSource): { width: number; height: number } | null {
  const anySource = source as any;
  const width =
    typeof anySource.width === "number"
      ? anySource.width
      : typeof anySource.naturalWidth === "number"
        ? anySource.naturalWidth
        : typeof anySource.videoWidth === "number"
          ? anySource.videoWidth
          : null;
  const height =
    typeof anySource.height === "number"
      ? anySource.height
      : typeof anySource.naturalHeight === "number"
        ? anySource.naturalHeight
        : typeof anySource.videoHeight === "number"
          ? anySource.videoHeight
          : null;

  if (typeof width !== "number" || typeof height !== "number") return null;
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
  if (width <= 0 || height <= 0) return null;
  return { width, height };
}

const MIN_PNG_BYTES = 33;

function assertPngNotTooLarge(dims: { width: number; height: number }): void {
  const pixels = dims.width * dims.height;
  if (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || pixels > MAX_PNG_PIXELS) {
    throw new Error(`Image dimensions too large (${dims.width}x${dims.height})`);
  }
}

function guardPngBytes(bytes: Uint8Array): void {
  const dims = readImageDimensions(bytes);
  if (!dims) return;
  assertPngNotTooLarge(dims);
  if (dims.format === "png" && bytes.byteLength < MIN_PNG_BYTES) {
    throw new Error("Invalid PNG: truncated IHDR");
  }
}

async function guardPngBlob(blob: Blob): Promise<void> {
  const readSlice = async (slice: Blob): Promise<Uint8Array | null> => {
    const anySlice = slice as any;
    if (typeof anySlice?.arrayBuffer === "function") {
      try {
        return new Uint8Array(await anySlice.arrayBuffer());
      } catch {
        return null;
      }
    }

    if (typeof (globalThis as any).FileReader === "function") {
      // Fallback for older browsers / jsdom where Blob#arrayBuffer is not implemented.
      try {
        const reader = new (globalThis as any).FileReader() as FileReader;
        const header = await new Promise<ArrayBuffer>((resolve, reject) => {
          reader.onerror = () => reject(reader.error ?? new Error("FileReader failed"));
          reader.onload = () => resolve(reader.result as ArrayBuffer);
          reader.readAsArrayBuffer(slice as Blob);
        });
        return new Uint8Array(header);
      } catch {
        return null;
      }
    }

    return null;
  };

  const TYPE_SNIFF_BYTES = 32;
  const firstReadSize = Math.min(blob.size, TYPE_SNIFF_BYTES);
  // We need at least a minimal signature to do anything useful.
  if (firstReadSize < 10) return;

  const initial = await readSlice(blob.slice(0, firstReadSize));
  if (!initial) return;

  const isJpeg =
    initial.byteLength >= 3 && initial[0] === 0xff && initial[1] === 0xd8 && initial[2] === 0xff;
  const isSvgLike = (() => {
    // SVG is text-based. It typically starts with `<` (possibly after a UTF-8 BOM/whitespace) or a UTF-16 BOM.
    const len = initial.byteLength;
    if (len >= 2 && ((initial[0] === 0xfe && initial[1] === 0xff) || (initial[0] === 0xff && initial[1] === 0xfe))) {
      return true;
    }
    let idx = 0;
    if (len >= 3 && initial[0] === 0xef && initial[1] === 0xbb && initial[2] === 0xbf) idx = 3; // UTF-8 BOM
    while (idx < len) {
      const b = initial[idx]!;
      if (b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d) {
        idx += 1;
        continue;
      }
      break;
    }
    // If the initial sniff slice is all whitespace, treat it as SVG-like so we can read more bytes.
    // Valid SVGs can include leading whitespace/newlines before the first tag.
    if (idx >= len) return true;
    return initial[idx] === 0x3c; // '<'
  })();
  let headerBytes = initial;

  if (isJpeg && blob.size > initial.byteLength) {
    // JPEG dimensions can occur after variable-length metadata segments, so allow a larger sniff.
    const MAX_JPEG_SNIFF_BYTES = 1024 * 1024;
    const toRead = Math.min(blob.size, MAX_JPEG_SNIFF_BYTES);
    if (toRead > initial.byteLength) {
      const larger = await readSlice(blob.slice(0, toRead));
      if (larger) headerBytes = larger;
    }
  } else if (isSvgLike && blob.size > initial.byteLength) {
    // SVG needs more than the initial signature bytes because the `<svg ...>` tag (and width/height)
    // can appear after an XML declaration and comments.
    const MAX_SVG_SNIFF_BYTES = 1024 * 1024;
    const toRead = Math.min(blob.size, MAX_SVG_SNIFF_BYTES);
    if (toRead > initial.byteLength) {
      const larger = await readSlice(blob.slice(0, toRead));
      if (larger) headerBytes = larger;
    }
  }

  const dims = readImageDimensions(headerBytes);
  if (!dims) return;
  assertPngNotTooLarge(dims);
  if (dims.format === "png" && blob.size < MIN_PNG_BYTES) {
    throw new Error("Invalid PNG: truncated IHDR");
  }
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
  strike?: boolean;
  strikethrough?: boolean;
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

function normalizeRichTextRuns(
  textLen: number,
  runs: Array<{ start: number; end: number; style?: Record<string, unknown> }> | undefined
): Array<{ start: number; end: number; style?: Record<string, unknown> }> {
  const len = clampIndex(textLen, 0, Number.isFinite(textLen) ? textLen : 0);
  if (len === 0) return [];

  if (!Array.isArray(runs) || runs.length === 0) {
    return [{ start: 0, end: len, style: undefined }];
  }

  const normalized = runs
    .map((run) => {
      const start = clampIndex(run?.start, 0, len);
      const end = clampIndex(run?.end, start, len);
      const style = run?.style && typeof run.style === "object" && !Array.isArray(run.style) ? run.style : undefined;
      return { start, end, style };
    })
    .filter((run) => run.end > run.start)
    .sort((a, b) => (a.start !== b.start ? a.start - b.start : a.end - b.end));

  if (normalized.length === 0) return [{ start: 0, end: len, style: undefined }];

  const out: Array<{ start: number; end: number; style?: Record<string, unknown> }> = [];
  let cursor = 0;
  for (const run of normalized) {
    if (cursor >= len) break;

    if (run.start > cursor) {
      out.push({ start: cursor, end: run.start, style: undefined });
      cursor = run.start;
    }

    const start = Math.max(cursor, run.start);
    const end = clampIndex(run.end, start, len);
    if (end > start) {
      out.push({ start, end, style: run.style });
      cursor = end;
    }
  }

  if (cursor < len) {
    out.push({ start: cursor, end: len, style: undefined });
  }

  return out;
}

type UnderlineSpec = { underline: boolean; underlineStyle?: "double" };

function resolveUnderline(style: RichTextRunStyle | undefined, defaultValue: UnderlineSpec): UnderlineSpec {
  const value = style?.underline;
  if (value === undefined) return defaultValue;

  if (value === false) return { underline: false, underlineStyle: undefined };
  if (value === true) return { underline: true, underlineStyle: undefined };

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "" || normalized === "none") return { underline: false, underlineStyle: undefined };
    // Excel commonly serializes underline variants as strings (e.g. `"double"`).
    if (normalized === "double" || normalized === "doubleaccounting") {
      return { underline: true, underlineStyle: "double" };
    }
    return { underline: true, underlineStyle: undefined };
  }

  return { underline: Boolean(value), underlineStyle: undefined };
}

function resolveStrike(style: RichTextRunStyle | undefined, defaultValue: boolean): boolean {
  const value = style?.strike ?? style?.strikethrough;
  if (value === undefined) return defaultValue;
  return value === true;
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
  const weight = style?.bold === true ? "bold" : style?.bold === false ? "normal" : defaults.weight;
  const fontStyle = style?.italic === true ? "italic" : style?.italic === false ? "normal" : defaults.style;
  return { family, sizePx, weight, style: fontStyle };
}

const EXPLICIT_NEWLINE_RE = /[\r\n]/;
const MAX_TEXT_OVERFLOW_COLUMNS = 128;
const EMPTY_MERGED_INDEX = new MergedCellIndex([]);
const DEFAULT_CELL_FONT_FAMILY = DEFAULT_GRID_FONT_FAMILY;

interface ViewportListenerEntry {
  listener: GridViewportChangeListener;
  options: GridViewportSubscriptionOptions;
  rafId: number | null;
  timeoutId: ReturnType<typeof setTimeout> | null;
  pendingReason: GridViewportChangeReason | null;
}

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
  private readonly dirtyRegionsScratch = {
    background: [] as Rect[],
    content: [] as Rect[],
    selection: [] as Rect[]
  };
  private readonly selectionDirtyRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly fullViewportRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly fullViewportRectListScratch: Rect[] = [this.fullViewportRectScratch];
  private readonly intersectionRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly scrollDirtyRectPool: Rect[] = [
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 }
  ];
  private scrollDirtyRectPoolIndex = 0;

  private scheduled = false;
  private renderRafId: number | null = null;
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
  private readonly activeSelectionRangeScratch: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
  private rangeSelection: CellRange | null = null;
  private fillPreviewRange: CellRange | null = null;
  private readonly fillPreviewRangeScratch: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
  private fillHandleEnabled = true;
  private fillHandleRectMemoRange: CellRange | null = null;
  private fillHandleRectMemoViewport: GridViewportState | null = null;
  private fillHandleRectMemoZoom = -1;
  private fillHandleRectMemoVisible = false;
  private readonly fillHandleRectMemoRect: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private referenceHighlights: Array<{ range: CellRange; color: string; active: boolean }> = [];

  private remotePresences: GridPresence[] = [];
  private remotePresenceDirtyPadding = 1;
  private readonly remotePresenceRangeScratch: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
  private readonly remotePresenceCursorCellRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };

  private readonly textWidthCache = new LruCache<string, number>(10_000);
  private textLayoutEngine?: TextLayoutEngine;

  private readonly imageResolver: CanvasGridImageResolver | null;
  private readonly imageErrorRetryMs = 250;
  private readonly imageBitmapCacheMaxEntries: number;
  private imageBitmapCacheReadyCount = 0;
  private readonly imageBitmapCache = new Map<
    string,
    | { state: "pending"; promise: Promise<void>; bitmap: null }
    | { state: "ready"; promise: null; bitmap: CanvasImageSource }
    | { state: "missing"; promise: null; bitmap: null }
    | { state: "error"; promise: null; bitmap: null; error: unknown; expiresAtMs: number }
  >();
  private imagePlaceholderPattern: { pattern: CanvasPattern; zoomKey: number } | null = null;
  private destroyed = false;

  private presenceFont: string;
  private theme: GridTheme;
  private readonly defaultCellFontFamily: string;
  private readonly defaultHeaderFontFamily: string;

  // Optional worksheet-level background pattern (tiled) rendered behind cell fills.
  //
  // This is intentionally *not* part of GridTheme because:
  // - the theme is serializable CSS colors only
  // - the background image is workbook data, not a styling token
  private backgroundPatternImage: CanvasImageSource | null = null;
  private backgroundPatternTile: HTMLCanvasElement | null = null;
  private backgroundPatternTileKey: { image: CanvasImageSource; zoom: number; devicePixelRatio: number } | null = null;

  private readonly perfStats: GridPerfStats = {
    enabled: false,
    lastFrameMs: 0,
    cellsPainted: 0,
    cellFetches: 0,
    dirtyRects: { background: 0, content: 0, selection: 0, total: 0 },
    blitUsed: false
  };

  // Per-frame (cleared each render) caches used by the hot-path `renderGridQuadrant` renderer.
  //
  // These are shared across quadrant renders within a single frame to:
  // - avoid repeated `provider.getCell()` calls when logic such as text overflow probing touches
  //   cells outside the currently-rendered quadrant
  // - avoid per-frame allocations (especially nested Map-of-Map row caches)
  private readonly frameCellCache = new Map<number, CellData | null>();
  private readonly frameBlockedCache = new Map<number, boolean>();
  private readonly frameCellCacheNested = new Map<number, Map<number, CellData | null>>();
  private readonly frameBlockedCacheNested = new Map<number, Map<number, boolean>>();
  private frameCacheUsesLinearKeys = true;
  private frameCacheColCount = 0;

  // Test-only instrumentation: counts the number of inner (per-row) Map allocations when we
  // fall back to the nested cache strategy (used only for extremely large grids where a
  // linearized numeric key could exceed Number.MAX_SAFE_INTEGER).
  private __testOnly_rowCacheMapAllocs = 0;

  private readonly viewportListeners = new Set<ViewportListenerEntry>();

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
  private readonly rangeToViewportRectsScratchRects: [Rect, Rect, Rect, Rect] = [
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 },
    { x: 0, y: 0, width: 0, height: 0 }
  ];
  private mergedIndex: MergedCellIndex = EMPTY_MERGED_INDEX;
  private mergedIndexKey: string | null = null;
  private mergedIndexViewport: GridViewportState | null = null;
  private mergedIndexDirty = true;

  constructor(options: CanvasGridRendererOptions) {
    this.provider = options.provider;
    this.prefetchOverscanRows = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanRows);
    this.prefetchOverscanCols = CanvasGridRenderer.sanitizeOverscan(options.prefetchOverscanCols);
    this.theme = resolveGridTheme(options.theme);
    this.imageResolver = options.imageResolver ?? null;
    this.imageBitmapCacheMaxEntries = CanvasGridRenderer.sanitizeImageBitmapCacheMaxEntries(options.imageBitmapCacheMaxEntries);

    const sanitizeFontFamily = (value: string | undefined): string | null => {
      if (typeof value !== "string") return null;
      const trimmed = value.trim();
      return trimmed ? trimmed : null;
    };

    const defaultCellFontFamily = sanitizeFontFamily(options.defaultCellFontFamily) ?? DEFAULT_CELL_FONT_FAMILY;
    this.defaultCellFontFamily = defaultCellFontFamily;
    this.defaultHeaderFontFamily = sanitizeFontFamily(options.defaultHeaderFontFamily) ?? defaultCellFontFamily;
    this.presenceFont = `${12 * this.zoom}px ${ensureSansSerifFallback(this.defaultHeaderFontFamily)}`;

    this.backgroundPatternImage = options.backgroundPatternImage ?? null;
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

  private static sanitizeImageBitmapCacheMaxEntries(value: number | undefined): number {
    // Keep the default fairly small; ImageBitmaps can be large (full source resolution).
    const DEFAULT_MAX = 256;
    if (value === undefined) return DEFAULT_MAX;
    if (!Number.isFinite(value)) return DEFAULT_MAX;
    // Disallow 0 here; the renderer relies on caching ready images to ever draw them.
    return Math.max(1, Math.floor(value));
  }

  getPerfStats(): Readonly<GridPerfStats> {
    return this.perfStats;
  }

  setPerfStatsEnabled(enabled: boolean): void {
    this.perfStats.enabled = enabled;
  }

  /**
   * Subscribe to notifications when the viewport *layout* changes.
   *
   * This intentionally does **not** fire for scroll offset changes (to avoid per-frame work during
   * scrolling), but it does fire for changes that can affect scrollbar layout/metrics such as:
   * - axis size changes (row/col resize, applyAxisSizeOverrides, etc)
   * - frozen pane changes
   * - viewport resize
   * - zoom changes
   */
  subscribeViewport(listener: GridViewportChangeListener, options?: GridViewportSubscriptionOptions): () => void {
    const entry: ViewportListenerEntry = {
      listener,
      options: options ?? {},
      rafId: null,
      timeoutId: null,
      pendingReason: null
    };
    this.viewportListeners.add(entry);

    return () => {
      if (!this.viewportListeners.delete(entry)) return;
      if (entry.rafId !== null) {
        globalThis.cancelAnimationFrame?.(entry.rafId);
        entry.rafId = null;
      }
      if (entry.timeoutId !== null) {
        clearTimeout(entry.timeoutId);
        entry.timeoutId = null;
      }
      entry.pendingReason = null;
    };
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
   * Update the optional worksheet-level tiled background image rendered behind cell fills.
   *
   * The background is painted on the grid/background layer and clipped so it does not render
   * underneath header rows/cols.
   */
  setBackgroundPatternImage(image: CanvasImageSource | null): void {
    const next = image ?? null;
    if (next === this.backgroundPatternImage) return;
    this.backgroundPatternImage = next;
    this.backgroundPatternTile = null;
    this.backgroundPatternTileKey = null;
    this.markAllDirty();
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

  private getBackgroundPatternTile(): HTMLCanvasElement | null {
    const image = this.backgroundPatternImage;
    if (!image) return null;

    const zoom = this.zoom;
    const rawDpr = Number.isFinite(this.devicePixelRatio) && this.devicePixelRatio > 0 ? this.devicePixelRatio : 1;
    const patternTransformSupported =
      typeof (globalThis as any).CanvasPattern !== "undefined" &&
      typeof (globalThis as any).CanvasPattern.prototype?.setTransform === "function";
    // Only bake DPR into the tile size when we can also scale the pattern back down in CSS space
    // (via `CanvasPattern.setTransform`). Otherwise the pattern would repeat at the wrong size.
    const devicePixelRatio = patternTransformSupported ? rawDpr : 1;
    const key = this.backgroundPatternTileKey;
    if (key && key.image === image && key.zoom === zoom && key.devicePixelRatio === devicePixelRatio && this.backgroundPatternTile) {
      return this.backgroundPatternTile;
    }

    const dims = getCanvasImageSourceDimensions(image);
    if (!dims) return null;

    // Scale the offscreen tile by both zoom and DPR so the pattern does not get
    // upscaled by the renderer's HiDPI transform (which would otherwise look
    // pixelated because the grid layer uses `imageSmoothingEnabled=false`).
    const width = Math.max(1, Math.round(dims.width * zoom * devicePixelRatio));
    const height = Math.max(1, Math.round(dims.height * zoom * devicePixelRatio));

    const tile = document.createElement("canvas");
    tile.width = width;
    tile.height = height;

    const ctx = tile.getContext("2d");
    if (!ctx) return null;
    // Background images look better with smoothing when zoomed/scaled; keep this separate
    // from the grid layer's `imageSmoothingEnabled=false` setting.
    ctx.imageSmoothingEnabled = true;
    try {
      ctx.drawImage(image, 0, 0, width, height);
    } catch {
      return null;
    }

    this.backgroundPatternTile = tile;
    this.backgroundPatternTileKey = { image, zoom, devicePixelRatio };
    return tile;
  }

  getZoom(): number {
    return this.zoom;
  }

  setZoom(nextZoom: number, options?: { anchorX?: number; anchorY?: number }): void {
    const clamped = clampGridZoom(nextZoom);
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

    // Apply overrides in bulk to avoid O(n^2) prefix-diff updates when many overrides exist.
    const nextRowOverrides = new Map<number, number>();
    for (const [row, baseHeight] of this.rowHeightOverridesBase) {
      nextRowOverrides.set(row, baseHeight * clamped);
    }
    nextScroll.rows.setOverrides(nextRowOverrides);

    const nextColOverrides = new Map<number, number>();
    for (const [col, baseWidth] of this.colWidthOverridesBase) {
      nextColOverrides.set(col, baseWidth * clamped);
    }
    nextScroll.cols.setOverrides(nextColOverrides);

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
    this.presenceFont = `${12 * this.zoom}px ${ensureSansSerifFallback(this.defaultHeaderFontFamily)}`;

    this.scroll.setScroll(targetScrollX, targetScrollY);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);

    if (this.selectionCtx) {
      this.remotePresenceDirtyPadding = this.getRemotePresenceDirtyPadding(this.selectionCtx);
    }

    this.markAllDirty();
    this.notifyViewportChange("zoom");
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
    this.destroyed = true;
    // Cancel any pending render request scheduled via `requestRender()`.
    if (this.renderRafId !== null) {
      try {
        globalThis.cancelAnimationFrame?.(this.renderRafId);
      } catch {
        // ignore
      }
      this.renderRafId = null;
    }
    this.scheduled = false;
    if (this.unsubscribeProvider) {
      this.unsubscribeProvider();
      this.unsubscribeProvider = undefined;
    }
    if (this.viewportListeners.size > 0) {
      for (const entry of this.viewportListeners) {
        if (entry.rafId !== null) {
          globalThis.cancelAnimationFrame?.(entry.rafId);
          entry.rafId = null;
        }
        if (entry.timeoutId !== null) {
          clearTimeout(entry.timeoutId);
          entry.timeoutId = null;
        }
        entry.pendingReason = null;
      }
      this.viewportListeners.clear();
    }
    this.clearImageCache();

    // Drop caches and release backing stores so a destroyed renderer does not retain large
    // allocations (text layout caches, offscreen canvases, etc) if it remains referenced
    // after teardown (tests, hot reload, view swaps).
    try {
      this.textWidthCache.clear();
    } catch {
      // ignore
    }
    this.textLayoutEngine = undefined;
    this.rowHeightOverridesBase.clear();
    this.colWidthOverridesBase.clear();
    this.referenceHighlights = [];
    this.remotePresences = [];
    this.imagePlaceholderPattern = null;
    this.backgroundPatternImage = null;
    this.backgroundPatternTile = null;
    this.backgroundPatternTileKey = null;
    this.mergedIndex = EMPTY_MERGED_INDEX;
    this.mergedIndexKey = null;
    this.mergedIndexDirty = true;
    this.frameCellCache.clear();
    this.frameBlockedCache.clear();
    this.frameCellCacheNested.clear();
    this.frameBlockedCacheNested.clear();

    // Release any canvas backing stores (these can hold multi-megabyte buffers even after the
    // canvases are removed from the DOM).
    const shrinkCanvas = (canvas: HTMLCanvasElement | undefined) => {
      if (!canvas) return;
      try {
        canvas.width = 0;
        canvas.height = 0;
      } catch {
        // ignore
      }
    };
    shrinkCanvas(this.gridCanvas);
    shrinkCanvas(this.contentCanvas);
    shrinkCanvas(this.selectionCanvas);
    shrinkCanvas(this.blitCanvas);
    this.blitCanvas = undefined;
    this.blitCtx = undefined;
    this.gridCanvas = undefined;
    this.gridCtx = undefined;
    this.contentCanvas = undefined;
    this.contentCtx = undefined;
    this.selectionCanvas = undefined;
    this.selectionCtx = undefined;
  }

  invalidateImage(imageId: string): void {
    const normalizedId = typeof imageId === "string" ? imageId.trim() : "";
    if (normalizedId === "") return;
    const existing = this.imageBitmapCache.get(normalizedId);
    if (existing?.state === "ready") {
      this.imageBitmapCacheReadyCount = Math.max(0, this.imageBitmapCacheReadyCount - 1);
      const bitmap = existing.bitmap as any;
      if (bitmap && typeof bitmap.close === "function") {
        try {
          bitmap.close();
        } catch {
          // Ignore bitmap disposal failures (best-effort).
        }
      }
    }
    this.imageBitmapCache.delete(normalizedId);
    this.markContentDirtyForImageUpdate();
  }

  clearImageCache(): void {
    // Close any decoded bitmaps, then drop all entries (including pending/error/missing).
    //
    // Important: avoid calling `invalidateImage()` per entry here:
    // - `invalidateImage` marks the viewport dirty and schedules a render
    // - tests may stub `requestAnimationFrame` to run synchronously
    // - re-entrant renders can repopulate `imageBitmapCache` while we're iterating it,
    //   causing the loop to never terminate.
    //
    // Instead, dispose any decoded bitmaps and clear the cache in one shot, then request
    // a single repaint at the end.
    for (const entry of this.imageBitmapCache.values()) {
      if (entry.state !== "ready") continue;
      const bitmap = entry.bitmap as any;
      if (bitmap && typeof bitmap.close === "function") {
        try {
          bitmap.close();
        } catch {
          // Ignore bitmap disposal failures (best-effort).
        }
      }
    }

    this.imageBitmapCache.clear();
    this.imageBitmapCacheReadyCount = 0;
    this.imagePlaceholderPattern = null;
    this.markContentDirtyForImageUpdate();
  }

  private markContentDirtyForImageUpdate(): void {
    if (this.destroyed) return;
    if (!this.contentCtx) return;
    const viewport = this.scroll.getViewportState();
    if (viewport.width <= 0 || viewport.height <= 0) return;
    this.dirty.content.markDirty({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    this.requestRender();
  }

  private getOrRequestImageBitmap(imageId: string): CanvasImageSource | null {
    const normalizedId = typeof imageId === "string" ? imageId.trim() : "";
    if (normalizedId === "") return null;

    const existing = this.imageBitmapCache.get(normalizedId);
    if (existing?.state === "ready") {
      // Touch for LRU eviction.
      this.imageBitmapCache.delete(normalizedId);
      this.imageBitmapCache.set(normalizedId, existing);
      return existing.bitmap;
    }
    if (existing?.state === "error") {
      if (existing.expiresAtMs <= Date.now()) {
        this.imageBitmapCache.delete(normalizedId);
      } else {
        return null;
      }
    } else if (existing) {
      return null;
    }

    this.requestImageBitmap(normalizedId);
    return null;
  }

  private requestImageBitmap(imageId: string): void {
    if (this.imageBitmapCache.has(imageId)) return;

    const resolver = this.imageResolver;
    if (!resolver) {
      this.imageBitmapCache.set(imageId, { state: "missing", promise: null, bitmap: null });
      return;
    }

    const placeholder: { state: "pending"; promise: Promise<void>; bitmap: null } = {
      state: "pending",
      // Assigned below.
      promise: Promise.resolve(),
      bitmap: null
    };
    this.imageBitmapCache.set(imageId, placeholder);

    const tryClose = (value: unknown) => {
      const bitmap = value as any;
      if (!bitmap || typeof bitmap.close !== "function") return;
      try {
        bitmap.close();
      } catch {
        // Ignore disposal failures (best-effort).
      }
    };

    const promise = (async () => {
      try {
        const source = await resolver(imageId);
        if (this.destroyed) {
          tryClose(source);
          return;
        }
        if (this.imageBitmapCache.get(imageId) !== placeholder) {
          tryClose(source);
          return;
        }

        if (source == null) {
          this.imageBitmapCache.set(imageId, { state: "missing", promise: null, bitmap: null });
          return;
        }

        const decoded = await this.decodeImageSource(source);
        if (this.destroyed) {
          tryClose(decoded);
          return;
        }
        if (this.imageBitmapCache.get(imageId) !== placeholder) {
          tryClose(decoded);
          return;
        }

        if (!decoded) {
          this.imageBitmapCache.set(imageId, { state: "missing", promise: null, bitmap: null });
          return;
        }

        this.imageBitmapCacheReadyCount++;
        // Replace + touch for LRU ordering (Map#set for an existing key does not
        // update insertion order).
        this.imageBitmapCache.delete(imageId);
        this.imageBitmapCache.set(imageId, { state: "ready", promise: null, bitmap: decoded });
        this.evictImageBitmapsIfNeeded();
      } catch (error) {
        if (this.destroyed) return;
        if (this.imageBitmapCache.get(imageId) !== placeholder) return;
        this.imageBitmapCache.set(imageId, {
          state: "error",
          promise: null,
          bitmap: null,
          error,
          expiresAtMs: Date.now() + this.imageErrorRetryMs,
        });
      } finally {
        // Repaint visible placeholders once the image finishes resolving.
        this.markContentDirtyForImageUpdate();
      }
    })();

    placeholder.promise = promise;
  }

  private evictImageBitmapsIfNeeded(): void {
    while (this.imageBitmapCacheReadyCount > this.imageBitmapCacheMaxEntries) {
      let evicted = false;
      for (const [imageId, entry] of this.imageBitmapCache) {
        if (entry.state !== "ready") continue;
        this.imageBitmapCache.delete(imageId);
        this.imageBitmapCacheReadyCount = Math.max(0, this.imageBitmapCacheReadyCount - 1);
        const bitmap = entry.bitmap as any;
        if (bitmap && typeof bitmap.close === "function") {
          try {
            bitmap.close();
          } catch {
            // Ignore disposal failures (best-effort).
          }
        }
        evicted = true;
        break;
      }
      if (!evicted) return;
    }
  }

  private async decodeImageSource(source: CanvasGridImageSource): Promise<CanvasImageSource | null> {
    if (typeof ImageBitmap !== "undefined" && source instanceof ImageBitmap) {
      return source;
    }

    const create = (globalThis as any).createImageBitmap as ((src: any) => Promise<ImageBitmap>) | undefined;

    if (source instanceof Blob) {
      if (typeof create !== "function") return null;
      try {
        await guardPngBlob(source);
        return await create(source);
      } catch (err) {
        // Chrome can decode some malformed PNGs via `<img>` but rejects
        // `createImageBitmap(blob)` (eg certain real Excel fixtures). Fall back
        // to decoding through an `<img>` + canvas for this specific failure
        // mode so images remain renderable.
        const name = (err as any)?.name;
        if (name !== "InvalidStateError") throw err;

        const fallback = await this.decodeBlobViaImageElement(source, create).catch(() => null);
        if (fallback) return fallback;
        throw err;
      }
    }

    if (source instanceof ArrayBuffer) {
      if (typeof create !== "function") return null;
      guardPngBytes(new Uint8Array(source));
      return await create(new Blob([source]));
    }

    if (source instanceof Uint8Array) {
      if (typeof create !== "function") return null;
      guardPngBytes(source);
      // Copy into a standalone buffer to avoid retaining oversized backing stores.
      const buffer = new ArrayBuffer(source.byteLength);
      new Uint8Array(buffer).set(source);
      return await create(new Blob([buffer]));
    }

    // For drawable sources, prefer `createImageBitmap` when available for performance.
    if (typeof create === "function") {
      try {
        return await create(source as any);
      } catch {
        // Fall back to using the source directly when supported by `drawImage`.
      }
    }

    if (source instanceof HTMLImageElement || source instanceof HTMLCanvasElement) return source;
    if (typeof OffscreenCanvas !== "undefined" && source instanceof OffscreenCanvas) return source;

    // ImageData cannot be drawn directly via `drawImage` without a bitmap conversion.
    return null;
  }

  private async decodeBlobViaImageElement(
    blob: Blob,
    create: (src: any) => Promise<ImageBitmap>,
  ): Promise<CanvasImageSource> {
    if (typeof document === "undefined") {
      throw new Error("Image decode fallback requires DOM APIs (missing document)");
    }
    if (typeof Image !== "function") {
      throw new Error("Image decode fallback requires DOM APIs (missing Image)");
    }
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") {
      throw new Error("Image decode fallback requires URL.createObjectURL");
    }

    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      await new Promise<void>((resolve, reject) => {
        const timeoutMs = 5_000;
        let timeoutId: ReturnType<typeof setTimeout> | null = null;
        let settled = false;

        const finish = (fn: () => void) => {
          if (settled) return;
          settled = true;
          if (timeoutId !== null) {
            try {
              clearTimeout(timeoutId);
            } catch {
              // Ignore clear failures (best-effort).
            }
            timeoutId = null;
          }
          // Drop references to callbacks to allow GC even if the image object remains alive.
          img.onload = null;
          img.onerror = null;
          fn();
        };

        img.onload = () => {
          finish(resolve);
        };
        img.onerror = () => {
          finish(() => reject(new Error("Image decode fallback failed to load <img>")));
        };

        if (typeof setTimeout === "function") {
          timeoutId = setTimeout(() => {
            timeoutId = null;
            finish(() => reject(new Error("Image decode fallback timed out")));
          }, timeoutMs);
        }

        // Assign the src after wiring handlers so we don't miss synchronous load events in tests/polyfills.
        img.src = url;
      });

      const canvas = document.createElement("canvas");
      const width = (img as any).naturalWidth ?? img.width;
      const height = (img as any).naturalHeight ?? img.height;
      canvas.width = Number.isFinite(width) && width > 0 ? width : 1;
      canvas.height = Number.isFinite(height) && height > 0 ? height : 1;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Image decode fallback missing 2D canvas context");
      ctx.drawImage(img, 0, 0);

      try {
        return await create(canvas);
      } catch {
        // If bitmap allocation fails for any reason, fall back to the drawn
        // canvas directly (still drawable via `drawImage`).
        return canvas;
      }
    } finally {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // Ignore revoke failures (best-effort).
      }
    }
  }

  private getImagePlaceholderPattern(ctx: CanvasRenderingContext2D, zoom: number): CanvasPattern | null {
    const createPattern = (ctx as any).createPattern as CanvasRenderingContext2D["createPattern"] | undefined;
    if (typeof createPattern !== "function") return null;

    const tile = Math.max(4, Math.round(6 * zoom));
    const cached = this.imagePlaceholderPattern;
    if (cached && cached.zoomKey === tile) return cached.pattern;

    const canvas = document.createElement("canvas");
    canvas.width = tile * 2;
    canvas.height = tile * 2;
    const pctx = canvas.getContext("2d");
    if (!pctx) return null;

    pctx.fillStyle = "#d0d0d0";
    pctx.fillRect(0, 0, canvas.width, canvas.height);
    pctx.fillStyle = "#e6e6e6";
    pctx.fillRect(0, 0, tile, tile);
    pctx.fillRect(tile, tile, tile, tile);

    const pattern = createPattern.call(ctx, canvas, "repeat");
    if (!pattern) return null;

    this.imagePlaceholderPattern = { pattern, zoomKey: tile };
    return pattern;
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
    this.notifyViewportChange("resize");
  }

  setFrozen(frozenRows: number, frozenCols: number): void {
    this.scroll.setFrozen(frozenRows, frozenCols);
    this.markAllDirty();
    this.notifyViewportChange("frozen");
  }

  setScroll(scrollX: number, scrollY: number): void {
    const before = this.scroll.getScroll();
    this.scroll.setScroll(scrollX, scrollY);
    const aligned = this.alignScrollToDevicePixels(this.scroll.getScroll());
    this.scroll.setScroll(aligned.x, aligned.y);
    const after = this.scroll.getScroll();
    if (before.x !== after.x || before.y !== after.y) this.invalidateForScroll();
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
    if (
      range === this.fillHandleRectMemoRange &&
      viewport === this.fillHandleRectMemoViewport &&
      this.zoom === this.fillHandleRectMemoZoom
    ) {
      return this.fillHandleRectMemoVisible ? this.fillHandleRectMemoRect : null;
    }

    this.fillHandleRectMemoRange = range;
    this.fillHandleRectMemoViewport = viewport;
    this.fillHandleRectMemoZoom = this.zoom;
    this.fillHandleRectMemoVisible = this.fillHandleRectInViewport(range, viewport, this.fillHandleRectMemoRect);

    return this.fillHandleRectMemoVisible ? this.fillHandleRectMemoRect : null;
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

  /**
   * Update the currently active selection range while preserving all other ranges.
   *
   * This is intended for high-frequency interactions (e.g. pointer-driven selection drags)
   * where repeatedly calling {@link setSelectionRanges} would allocate/clamp/expand every
   * range on every pointer move.
   *
   * Returns `true` when the selection state actually changed.
   */
  setActiveSelectionRange(range: CellRange): boolean {
    if (this.selectionRanges.length === 0) {
      const prev = this.selectionRanges.length;
      this.setSelectionRange(range);
      return prev !== this.selectionRanges.length;
    }

    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    let startRow = clampIndex(range.startRow, 0, rowCount);
    let endRow = clampIndex(range.endRow, 0, rowCount);
    let startCol = clampIndex(range.startCol, 0, colCount);
    let endCol = clampIndex(range.endCol, 0, colCount);

    if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
    if (startCol > endCol) [startCol, endCol] = [endCol, startCol];
    if (startRow === endRow || startCol === endCol) return false;

    const provider = this.provider;
    if (provider.getMergedRangeAt || provider.getMergedRangesInRange) {
      const scratch = this.activeSelectionRangeScratch;
      scratch.startRow = startRow;
      scratch.endRow = endRow;
      scratch.startCol = startCol;
      scratch.endCol = endCol;
      const expanded = this.expandRangeToMergedCells(scratch);
      startRow = expanded.startRow;
      endRow = expanded.endRow;
      startCol = expanded.startCol;
      endCol = expanded.endCol;
    }

    const activeIndex = clampIndex(this.activeSelectionIndex, 0, this.selectionRanges.length - 1);
    const activeRange = this.selectionRanges[activeIndex]!;

    if (
      activeRange.startRow === startRow &&
      activeRange.endRow === endRow &&
      activeRange.startCol === startCol &&
      activeRange.endCol === endCol
    ) {
      return false;
    }

    activeRange.startRow = startRow;
    activeRange.endRow = endRow;
    activeRange.startCol = startCol;
    activeRange.endCol = endCol;

    // If we mutate the active range in-place, invalidate fill-handle memoization since it
    // caches by range reference (not by coordinates).
    this.fillHandleRectMemoRange = null;

    // Preserve the existing active cell (Excel-style behavior: anchor cell remains active),
    // but clamp it into the updated active range bounds.
    let selection = this.selection;
    if (selection) {
      const nextRow = clamp(selection.row, startRow, endRow - 1);
      const nextCol = clamp(selection.col, startCol, endCol - 1);
      if (nextRow !== selection.row || nextCol !== selection.col) {
        selection.row = nextRow;
        selection.col = nextCol;
      }
    } else {
      selection = { row: startRow, col: startCol };
      this.selection = selection;
    }

    const resolvedActive = this.getMergedAnchorForCell(selection.row, selection.col);
    if (resolvedActive && (resolvedActive.row !== selection.row || resolvedActive.col !== selection.col)) {
      selection.row = resolvedActive.row;
      selection.col = resolvedActive.col;
    }

    this.activeSelectionIndex = activeIndex;

    // The active range changed, so selection overlays must repaint. We also repaint when the
    // active cell was clamped or resolved into a merged anchor.
    this.markSelectionDirty();

    return true;
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

  /**
   * Returns viewport-space rectangles (relative to the grid canvases) for the provided sheet
   * `CellRange`, clipped to the visible viewport and split across frozen-row/column quadrants.
   *
   * This mirrors the internal selection rendering logic and is intended for overlay renderers
   * (diff/audit/presence/etc) that need pixel-accurate range geometry.
   */
  getRangeRects(range: CellRange): Rect[] {
    const normalized = this.normalizeSelectionRange(range);
    if (!normalized) return [];
    const viewport = this.scroll.getViewportState();
    return this.rangeToViewportRects(normalized, viewport);
  }

  getViewportState(): GridViewportState {
    return this.scroll.getViewportState();
  }

  setFillPreviewRange(range: CellRange | null): void {
    const previousRange = this.fillPreviewRange;
    if (!range) {
      if (!previousRange) return;
      this.fillPreviewRange = null;
      this.markSelectionDirty();
      return;
    }

    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    let startRow = clampIndex(range.startRow, 0, rowCount);
    let endRow = clampIndex(range.endRow, 0, rowCount);
    let startCol = clampIndex(range.startCol, 0, colCount);
    let endCol = clampIndex(range.endCol, 0, colCount);

    if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
    if (startCol > endCol) [startCol, endCol] = [endCol, startCol];
    if (startRow === endRow || startCol === endCol) {
      if (!previousRange) return;
      this.fillPreviewRange = null;
      this.markSelectionDirty();
      return;
    }

    if (
      previousRange &&
      previousRange.startRow === startRow &&
      previousRange.endRow === endRow &&
      previousRange.startCol === startCol &&
      previousRange.endCol === endCol
    ) {
      return;
    }

    const target = this.fillPreviewRangeScratch;
    target.startRow = startRow;
    target.endRow = endRow;
    target.startCol = startCol;
    target.endCol = endCol;
    this.fillPreviewRange = target;
    this.markSelectionDirty();
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
      // Keep runtime axis overrides consistent with the persisted base map; otherwise tiny
      // floating-point differences can leave a "default-sized" override installed.
      this.scroll.rows.deleteSize(row);
    } else {
      this.rowHeightOverridesBase.set(row, height / this.zoom);
      this.scroll.rows.setSize(row, height);
    }
    this.onAxisSizeChanged();
  }

  setColWidth(col: number, width: number): void {
    this.assertColIndex(col);
    if (Math.abs(width - this.scroll.cols.defaultSize) < 1e-6) {
      this.colWidthOverridesBase.delete(col);
      // Keep runtime axis overrides consistent with the persisted base map; otherwise tiny
      // floating-point differences can leave a "default-sized" override installed.
      this.scroll.cols.deleteSize(col);
    } else {
      this.colWidthOverridesBase.set(col, width / this.zoom);
      this.scroll.cols.setSize(col, width);
    }
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
   * Apply many row/column size overrides in one batch, triggering at most one full invalidation.
   *
   * Sizes are specified in CSS pixels at the current zoom.
   *
   * When `resetUnspecified` is set, any existing overrides for the provided axes that are *not*
   * present in the new maps are cleared.
   *
   * Notes:
   * - Indices are applied in ascending order to keep `VariableSizeAxis` updates predictable.
   * - Scroll clamping + device-pixel alignment happen once at the end.
   */
  applyAxisSizeOverrides(
    overrides: { rows?: ReadonlyMap<number, number>; cols?: ReadonlyMap<number, number> },
    options?: { resetUnspecified?: boolean }
  ): void {
    const rows = overrides.rows;
    const cols = overrides.cols;
    const resetUnspecified = options?.resetUnspecified ?? false;
    const zoom = this.zoom;
    const epsilon = 1e-6;
    let changed = false;

    const mapsEqual = (a: ReadonlyMap<number, number>, b: ReadonlyMap<number, number>): boolean => {
      if (a.size !== b.size) return false;
      for (const [key, value] of a) {
        const other = b.get(key);
        if (other === undefined) return false;
        if (Math.abs(other - value) > epsilon) return false;
      }
      return true;
    };

    const applyAxis = (axis: "rows" | "cols") => {
      const sizes = axis === "rows" ? rows : cols;
      if (!sizes) return;

      const baseOverrides = axis === "rows" ? this.rowHeightOverridesBase : this.colWidthOverridesBase;
      const variableAxis = axis === "rows" ? this.scroll.rows : this.scroll.cols;
      const assertIndex =
        axis === "rows" ? (idx: number) => this.assertRowIndex(idx) : (idx: number) => this.assertColIndex(idx);
      const defaultSize = variableAxis.defaultSize;

      // Build the next base override map by either starting from scratch or merging into the
      // current state (when `resetUnspecified` is false).
      const nextBase = resetUnspecified ? new Map<number, number>() : new Map<number, number>(baseOverrides);

      for (const [idx, size] of sizes) {
        assertIndex(idx);
        if (!Number.isFinite(size) || size <= 0) {
          throw new Error(`${axis} size must be a positive finite number, got ${size} for index ${idx}`);
        }

        // Treat "default-sized" entries as a request to clear the override.
        if (Math.abs(size - defaultSize) < epsilon) {
          nextBase.delete(idx);
          continue;
        }

        nextBase.set(idx, size / zoom);
      }

      if (mapsEqual(baseOverrides, nextBase)) return;

      // Update persisted base sizes (zoom=1).
      baseOverrides.clear();
      for (const [idx, baseSize] of nextBase) {
        baseOverrides.set(idx, baseSize);
      }

      // Replace the runtime VariableSizeAxis overrides in bulk using CSS sizes.
      const cssOverrides = new Map<number, number>();
      for (const [idx, baseSize] of nextBase) {
        cssOverrides.set(idx, baseSize * zoom);
      }
      variableAxis.setOverrides(cssOverrides);
      changed = true;
    };

    applyAxis("cols");
    applyAxis("rows");

    if (changed) this.onAxisSizeChanged();
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
    const headerRows = this.headerRowsOverride ?? (viewport.frozenRows > 0 ? 1 : 0);
    const headerCols = this.headerColsOverride ?? (viewport.frozenCols > 0 ? 1 : 0);

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
      const isHeader = row < headerRows || col < headerCols;
      const fontFamily = style?.fontFamily ?? (isHeader ? this.defaultHeaderFontFamily : this.defaultCellFontFamily);
      const fontWeight = style?.fontWeight ?? "400";
      const fontStyle = style?.fontStyle ?? "normal";
      const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle } satisfies FontSpec;

      if (hasRichText) {
        const offsets = buildCodePointIndex(richTextText);
        const textLen = offsets.length - 1;
        const rawRuns = normalizeRichTextRuns(textLen, richText?.runs);

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
    const headerRows = this.headerRowsOverride ?? (viewport.frozenRows > 0 ? 1 : 0);
    const headerCols = this.headerColsOverride ?? (viewport.frozenCols > 0 ? 1 : 0);

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
      const isHeader = row < headerRows || col < headerCols;
      const fontFamily = style?.fontFamily ?? (isHeader ? this.defaultHeaderFontFamily : this.defaultCellFontFamily);
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
        const rawRuns = normalizeRichTextRuns(textLen, richText?.runs);

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

  pickCellAt(viewportX: number, viewportY: number, out?: Selection): Selection | null {
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

    const result = out ?? { row, col };
    result.row = row;
    result.col = col;

    const merged = this.getMergedIndex(viewport).rangeAt(result);
    if (merged) {
      result.row = merged.startRow;
      result.col = merged.startCol;
    }
    return result;
  }

  renderImmediately(): void {
    if (this.destroyed) return;
    this.renderFrame();
  }

  requestRender(): void {
    if (this.destroyed) return;
    if (this.scheduled) return;
    this.scheduled = true;
    this.renderRafId = requestAnimationFrame(() => {
      this.renderRafId = null;
      this.scheduled = false;
      if (this.destroyed) return;
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
    if (this.destroyed) return;
    const perf = this.perfStats;
    const perfEnabled = perf.enabled;
    const frameStart = perfEnabled ? performance.now() : 0;

    const viewport = this.scroll.getViewportState();

    // Reset frame-local caches before any paint work begins.
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    this.frameCacheColCount = colCount;
    this.frameCacheUsesLinearKeys = CanvasGridRenderer.canUseLinearCellCacheKey(rowCount, colCount);
    this.frameCellCache.clear();
    this.frameBlockedCache.clear();
    this.frameCellCacheNested.clear();
    this.frameBlockedCacheNested.clear();
    this.__testOnly_rowCacheMapAllocs = 0;

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

    const backgroundRegions = this.dirty.background.drainInto(this.dirtyRegionsScratch.background);
    const contentRegions = this.dirty.content.drainInto(this.dirtyRegionsScratch.content);
    const selectionRegions = this.dirty.selection.drainInto(this.dirtyRegionsScratch.selection);

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

  private static canUseLinearCellCacheKey(rowCount: number, colCount: number): boolean {
    // A linearized key `row * colCount + col` is only guaranteed unique when every key stays
    // within the safe integer range. For extreme grids (very large rowCount/colCount), we
    // fall back to the previous nested Map-of-Map strategy to preserve correctness.
    if (!Number.isSafeInteger(rowCount) || !Number.isSafeInteger(colCount)) return false;
    if (rowCount <= 0 || colCount <= 0) return true;
    const maxKey = rowCount * colCount - 1;
    return Number.isFinite(maxKey) && maxKey <= Number.MAX_SAFE_INTEGER;
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
    return alignScrollToDevicePixelsUtil(pos, this.scroll.getMaxScroll(), this.devicePixelRatio);
  }

  private allocScrollDirtyRect(): Rect {
    const idx = this.scrollDirtyRectPoolIndex++;
    if (idx < this.scrollDirtyRectPool.length) {
      return this.scrollDirtyRectPool[idx]!;
    }
    const rect: Rect = { x: 0, y: 0, width: 0, height: 0 };
    this.scrollDirtyRectPool.push(rect);
    return rect;
  }

  private markDirtyBothPaddedIntoPool(x: number, y: number, width: number, height: number, paddingPx: number): void {
    const rect = this.allocScrollDirtyRect();
    rect.x = x - paddingPx;
    rect.y = y - paddingPx;
    rect.width = width + paddingPx * 2;
    rect.height = height + paddingPx * 2;
    this.dirty.background.markDirty(rect);
    this.dirty.content.markDirty(rect);
  }

  private markDirtySelectionPaddedIntoPool(x: number, y: number, width: number, height: number, padding: number): void {
    const p = Number.isFinite(padding) ? Math.max(1, Math.floor(padding)) : 1;
    const rect = this.allocScrollDirtyRect();
    rect.x = x - p;
    rect.y = y - p;
    rect.width = width + p * 2;
    rect.height = height + p * 2;
    this.dirty.selection.markDirty(rect);
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
    this.scrollDirtyRectPoolIndex = 0;

    const frozenWidth = viewport.frozenWidth;
    const frozenHeight = viewport.frozenHeight;

    const selectionPadding = 1;
    const borderDirtyPaddingPx = Math.max(1, Math.ceil(2 * this.zoom));

    // Frozen rows + scrollable columns (top-right): horizontal-only scroll.
    {
      const rectX = frozenWidth;
      const rectY = 0;
      const rectW = viewport.width - frozenWidth;
      const rectH = frozenHeight;
      if (rectW > 0 && rectH > 0) {
        const shiftX = deltaX;
        if (shiftX > 0) {
          this.markDirtyBothPaddedIntoPool(rectX, rectY, shiftX, rectH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, rectY, shiftX, rectH, selectionPadding);
        } else if (shiftX < 0) {
          const stripeX = rectX + rectW + shiftX;
          const stripeW = -shiftX;
          this.markDirtyBothPaddedIntoPool(stripeX, rectY, stripeW, rectH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(stripeX, rectY, stripeW, rectH, selectionPadding);
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
          this.markDirtyBothPaddedIntoPool(rectX, rectY, rectW, shiftY, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, rectY, rectW, shiftY, selectionPadding);
        } else if (shiftY < 0) {
          const stripeY = rectY + rectH + shiftY;
          const stripeH = -shiftY;
          this.markDirtyBothPaddedIntoPool(rectX, stripeY, rectW, stripeH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, stripeY, rectW, stripeH, selectionPadding);
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
          this.markDirtyBothPaddedIntoPool(rectX, rectY, shiftX, rectH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, rectY, shiftX, rectH, selectionPadding);
        } else if (shiftX < 0) {
          const stripeX = rectX + rectW + shiftX;
          const stripeW = -shiftX;
          this.markDirtyBothPaddedIntoPool(stripeX, rectY, stripeW, rectH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(stripeX, rectY, stripeW, rectH, selectionPadding);
        }

        if (shiftY > 0) {
          this.markDirtyBothPaddedIntoPool(rectX, rectY, rectW, shiftY, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, rectY, rectW, shiftY, selectionPadding);
        } else if (shiftY < 0) {
          const stripeY = rectY + rectH + shiftY;
          const stripeH = -shiftY;
          this.markDirtyBothPaddedIntoPool(rectX, stripeY, rectW, stripeH, borderDirtyPaddingPx);
          this.markDirtySelectionPaddedIntoPool(rectX, stripeY, rectW, stripeH, selectionPadding);
        }
      }
    }

    if (this.remotePresences.length > 0) {
      // Cursor name badges can overlap into frozen quadrants, which aren't shifted during blit.
      // Mark the (padded) cursor rects dirty in both the previous and next viewport so badges
      // are cleared/redrawn correctly.
      const previousScrollX = viewport.scrollX + deltaX;
      const previousScrollY = viewport.scrollY + deltaY;
      const cursorRect = this.remotePresenceCursorCellRectScratch;

      for (const presence of this.remotePresences) {
        const cursor = presence.cursor;
        if (!cursor) continue;

        if (this.cellRectInViewportIntoWithScroll(cursor.row, cursor.col, viewport, previousScrollX, previousScrollY, cursorRect)) {
          this.markDirtySelectionPaddedIntoPool(
            cursorRect.x,
            cursorRect.y,
            cursorRect.width,
            cursorRect.height,
            this.remotePresenceDirtyPadding
          );
        }
        if (this.cellRectInViewportInto(cursor.row, cursor.col, viewport, cursorRect)) {
          this.markDirtySelectionPaddedIntoPool(
            cursorRect.x,
            cursorRect.y,
            cursorRect.width,
            cursorRect.height,
            this.remotePresenceDirtyPadding
          );
        }
      }
    }

    // Freeze lines are drawn on the selection layer but should not move with scroll. When we blit
    // the selection layer, the previous freeze line pixels get shifted into the scrollable
    // quadrants, leaving "ghost" lines behind. Mark those shifted lines as dirty so they get
    // cleared and redrawn in the correct location.
    const ghostWidth = 6;
    if (viewport.frozenCols > 0 && deltaX !== 0) {
      const ghostX = crispLine(viewport.frozenWidth) + deltaX;
      const rect = this.allocScrollDirtyRect();
      rect.x = ghostX - ghostWidth;
      rect.y = 0;
      rect.width = ghostWidth * 2;
      rect.height = viewport.height;
      this.dirty.selection.markDirty(rect);
    }
    if (viewport.frozenRows > 0 && deltaY !== 0) {
      const ghostY = crispLine(viewport.frozenHeight) + deltaY;
      const rect = this.allocScrollDirtyRect();
      rect.x = 0;
      rect.y = ghostY - ghostWidth;
      rect.width = viewport.width;
      rect.height = ghostWidth * 2;
      this.dirty.selection.markDirty(rect);
    }
  }

  private markDirtyBoth(rect: Rect): void {
    // Similar to provider-update invalidation: explicit borders are strokes centered on cell edges,
    // so they can extend beyond the newly-exposed scroll stripe. Pad slightly to ensure borders
    // are cleared + redrawn correctly when using blit-based scrolling.
    const borderDirtyPaddingPx = Math.max(1, Math.ceil(2 * this.zoom));
    const padded = borderDirtyPaddingPx > 0 ? padRect(rect, borderDirtyPaddingPx) : rect;
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
    // Borders are rendered using strokes centered on cell edges; to ensure incremental renders
    // correctly clear + redraw borders (including thick/double borders that extend beyond the
    // updated cell rect), pad the dirty region slightly.
    //
    // We intentionally keep this as a small constant derived from zoom rather than scanning for
    // the exact maximum border width, which would be expensive for large update ranges.
    const borderDirtyPaddingPx = Math.max(1, Math.ceil(2 * this.zoom));
    for (const rect of rects) {
      const padded = borderDirtyPaddingPx > 0 ? padRect(rect, borderDirtyPaddingPx) : rect;
      this.dirty.background.markDirty(padded);
      this.dirty.content.markDirty(padded);
    }
    this.requestRender();
  }

  private rangeToViewportRectsScratch(range: CellRange, viewport: GridViewportState): number {
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

    const rects = this.rangeToViewportRectsScratchRects;
    let rectCount = 0;

    const addRect = (rowStart: number, rowEnd: number, colStart: number, colEnd: number, scrollRows: boolean, scrollCols: boolean) => {
      if (rowStart >= rowEnd || colStart >= colEnd) return;

      const x1 = colAxis.positionOf(colStart);
      const x2 = colAxis.positionOf(colEnd);
      const y1 = rowAxis.positionOf(rowStart);
      const y2 = rowAxis.positionOf(rowEnd);

      const x = scrollCols ? x1 - viewport.scrollX : x1;
      const y = scrollRows ? y1 - viewport.scrollY : y1;
      const width = x2 - x1;
      const height = y2 - y1;
      const xEnd = x + width;
      const yEnd = y + height;

      const quadrantX = scrollCols ? frozenWidth : 0;
      const quadrantY = scrollRows ? frozenHeight : 0;
      const quadrantWidth = scrollCols ? Math.max(0, viewport.width - frozenWidth) : frozenWidth;
      const quadrantHeight = scrollRows ? Math.max(0, viewport.height - frozenHeight) : frozenHeight;
      const quadrantXEnd = quadrantX + quadrantWidth;
      const quadrantYEnd = quadrantY + quadrantHeight;

      const ix1 = Math.max(x, quadrantX);
      const iy1 = Math.max(y, quadrantY);
      const ix2 = Math.min(xEnd, quadrantXEnd);
      const iy2 = Math.min(yEnd, quadrantYEnd);
      const iWidth = ix2 - ix1;
      const iHeight = iy2 - iy1;
      if (iWidth <= 0 || iHeight <= 0) return;

      const rect = rects[rectCount];
      rect.x = ix1;
      rect.y = iy1;
      rect.width = iWidth;
      rect.height = iHeight;
      rectCount++;
    };

    addRect(rowsFrozenStart, rowsFrozenEnd, colsFrozenStart, colsFrozenEnd, false, false);
    addRect(rowsFrozenStart, rowsFrozenEnd, colsScrollStart, colsScrollEnd, false, true);
    addRect(rowsScrollStart, rowsScrollEnd, colsFrozenStart, colsFrozenEnd, true, false);
    addRect(rowsScrollStart, rowsScrollEnd, colsScrollStart, colsScrollEnd, true, true);

    return rectCount;
  }

  private rangeToViewportRects(range: CellRange, viewport: GridViewportState): Rect[] {
    const count = this.rangeToViewportRectsScratch(range, viewport);
    const scratch = this.rangeToViewportRectsScratchRects;
    const rects: Rect[] = [];
    for (let i = 0; i < count; i++) {
      const rect = scratch[i]!;
      rects.push({ x: rect.x, y: rect.y, width: rect.width, height: rect.height });
    }
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

    const full = this.fullViewportRectScratch;
    full.x = 0;
    full.y = 0;
    full.width = viewport.width;
    full.height = viewport.height;

    let shouldFullRender = regions.length > 8;
    if (!shouldFullRender) {
      for (let i = 0; i < regions.length; i++) {
        const region = regions[i]!;
        if (region.x <= 0 && region.y <= 0 && region.width >= viewport.width && region.height >= viewport.height) {
          shouldFullRender = true;
          break;
        }
      }
    }

    const toRender = shouldFullRender ? this.fullViewportRectListScratch : regions;

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
          existing.x = x1;
          existing.y = y1;
          existing.width = x2 - x1;
          existing.height = y2 - y1;
          merged = existing;
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

    const full = this.fullViewportRectScratch;
    full.x = 0;
    full.y = 0;
    full.width = viewport.width;
    full.height = viewport.height;

    let shouldFullRender = regions.length > 8;
    if (!shouldFullRender) {
      for (let i = 0; i < regions.length; i++) {
        const region = regions[i]!;
        if (region.x <= 0 && region.y <= 0 && region.width >= viewport.width && region.height >= viewport.height) {
          shouldFullRender = true;
          break;
        }
      }
    }

    const toRender = shouldFullRender ? this.fullViewportRectListScratch : regions;

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

    const headerRows = this.headerRowsOverride ?? (viewport.frozenRows > 0 ? 1 : 0);
    const headerCols = this.headerColsOverride ?? (viewport.frozenCols > 0 ? 1 : 0);
    const dataOriginXSheet = headerCols > 0 ? this.scroll.cols.positionOf(headerCols) : 0;
    const dataOriginYSheet = headerRows > 0 ? this.scroll.rows.positionOf(headerRows) : 0;

    const backgroundTile = this.getBackgroundPatternTile();
    let backgroundPattern: CanvasPattern | null = null;
    if (backgroundTile) {
      const pattern = gridCtx.createPattern(backgroundTile, "repeat");
      if (pattern) {
        // If the tile was generated at HiDPI resolution, scale it back down into CSS-pixel space
        // so the pattern repeats at the expected logical size.
        //
        // (When supported, CanvasPattern.setTransform applies in addition to the context's current
        // transform, so scaling by 1/DPR cancels the HiDPI upscaling that `setupHiDpiCanvas` applies.)
        const setTransform = (pattern as any).setTransform as ((transform?: DOMMatrix2DInit) => void) | undefined;
        const dpr = Number.isFinite(this.devicePixelRatio) && this.devicePixelRatio > 0 ? this.devicePixelRatio : 1;
        if (typeof setTransform === "function" && dpr !== 1) {
          try {
            setTransform.call(pattern, { a: 1 / dpr, b: 0, c: 0, d: 1 / dpr, e: 0, f: 0 });
          } catch {
            // Ignore pattern transform failures (best-effort; pattern will still render, just lower quality).
          }
        }
        backgroundPattern = pattern;
      }
    }

    for (const quadrant of quadrants) {
      if (quadrant.rect.width <= 0 || quadrant.rect.height <= 0) continue;
      if (quadrant.maxRowExclusive <= quadrant.minRow || quadrant.maxColExclusive <= quadrant.minCol) continue;

      const intersection = this.intersectionRectScratch;
      {
        const x1 = Math.max(region.x, quadrant.rect.x);
        const y1 = Math.max(region.y, quadrant.rect.y);
        const x2 = Math.min(region.x + region.width, quadrant.rect.x + quadrant.rect.width);
        const y2 = Math.min(region.y + region.height, quadrant.rect.y + quadrant.rect.height);
        const width = x2 - x1;
        const height = y2 - y1;
        if (width <= 0 || height <= 0) continue;
        intersection.x = x1;
        intersection.y = y1;
        intersection.width = width;
        intersection.height = height;
      }

      const sheetX = quadrant.scrollBaseX + (intersection.x - quadrant.originX);
      const sheetY = quadrant.scrollBaseY + (intersection.y - quadrant.originY);
      const sheetXEnd = sheetX + intersection.width;
      const sheetYEnd = sheetY + intersection.height;
      // Treat the intersection end coordinates as *exclusive* bounds. `indexAt(pos)` returns the
      // last index whose start position is `<= pos`, so calling it with an exact cell boundary
      // can otherwise "round up" and include a fully-clipped trailing row/col (wasted work, and
      // can trigger unnecessary image decode requests).
      const epsilon = 1e-6;

      const startRow = this.scroll.rows.indexAt(sheetY, {
        min: quadrant.minRow,
        maxInclusive: quadrant.maxRowExclusive - 1
      });
      const endRow = Math.min(
        this.scroll.rows.indexAt(Math.max(sheetY, sheetYEnd - epsilon), {
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
        this.scroll.cols.indexAt(Math.max(sheetX, sheetXEnd - epsilon), {
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

      if (backgroundPattern) {
        const bodyX = Math.max(intersection.x, quadrant.originX + (dataOriginXSheet - quadrant.scrollBaseX));
        const bodyY = Math.max(intersection.y, quadrant.originY + (dataOriginYSheet - quadrant.scrollBaseY));
        const bodyW = intersection.x + intersection.width - bodyX;
        const bodyH = intersection.y + intersection.height - bodyY;
        if (bodyW > 0 && bodyH > 0) {
          // Align the pattern to the sheet data origin (A1), excluding header rows/cols.
          // We shift the pattern by applying a translate, then offsetting the draw rect
          // back so the filled pixels land in the same viewport rect.
          const tx = quadrant.originX + dataOriginXSheet - quadrant.scrollBaseX;
          const ty = quadrant.originY + dataOriginYSheet - quadrant.scrollBaseY;
          gridCtx.save();
          gridCtx.translate(tx, ty);
          gridCtx.fillStyle = backgroundPattern;
          gridCtx.fillRect(bodyX - tx, bodyY - ty, bodyW, bodyH);
          gridCtx.restore();
        }
      }

      this.renderGridQuadrant(
        quadrant,
        mergedIndex,
        startRow,
        endRow,
        startCol,
        endCol,
        viewport.frozenRows,
        viewport.frozenCols,
        viewport.scrollX,
        viewport.scrollY,
        headerRows,
        headerCols,
        perf
      );

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
    frozenRows: number,
    frozenCols: number,
    scrollX: number,
    scrollY: number,
    headerRows: number,
    headerCols: number,
    perf: GridPerfStats | null
  ): void {
    if (!this.gridCtx || !this.contentCtx) return;
    const gridCtx = this.gridCtx;
    const contentCtx = this.contentCtx;
    const theme = this.theme;
    const gridBg = theme.gridBg;
    const hasBackgroundPattern = this.backgroundPatternImage !== null;
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
    const decorationLineWidth = Math.max(1, zoom);

    // Font specs are part of the text-layout cache key and are returned in layout runs.
    // Avoid mutating a shared object after passing it to the layout engine.
    let fontSpec = { family: this.defaultCellFontFamily, sizePx: 12 * zoom, weight: "400", style: "normal" } satisfies FontSpec;
    let currentFontFamily = "";
    let currentFontSize = -1;
    let currentFontWeight = "";
    let currentFontStyle = "";

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();

    const mergedRanges = mergedIndex.getRanges();
    let hasMerges = false;

    // Reuse a single object for merged range lookups to avoid per-cell allocations in hot paths.
    const mergedCellRef = { row: 0, col: 0 };
    const mergedRangeAt = (row: number, col: number): CellRange | null => {
      mergedCellRef.row = row;
      mergedCellRef.col = col;
      return mergedIndex.rangeAt(mergedCellRef);
    };

    const useLinearKeys = this.frameCacheUsesLinearKeys && this.frameCacheColCount === colCount;

    const getCellCached: (row: number, col: number) => CellData | null = useLinearKeys
      ? (row, col) => {
          const key = row * colCount + col;
          if (this.frameCellCache.has(key)) return this.frameCellCache.get(key) ?? null;
          const cell = this.provider.getCell(row, col);
          if (trackCellFetches) cellFetches += 1;
          this.frameCellCache.set(key, cell);
          return cell;
        }
      : (row, col) => {
          let rowCache = this.frameCellCacheNested.get(row);
          if (!rowCache) {
            rowCache = new Map();
            this.frameCellCacheNested.set(row, rowCache);
            this.__testOnly_rowCacheMapAllocs += 1;
          }
          if (rowCache.has(col)) return rowCache.get(col) ?? null;
          const cell = this.provider.getCell(row, col);
          if (trackCellFetches) cellFetches += 1;
          rowCache.set(col, cell);
          return cell;
        };

    const isBlockedForOverflow: (row: number, col: number) => boolean = useLinearKeys
      ? (row, col) => {
          const key = row * colCount + col;
          if (this.frameBlockedCache.has(key)) return this.frameBlockedCache.get(key) ?? false;

          if (mergedRanges.length > 0 && mergedRangeAt(row, col)) {
            this.frameBlockedCache.set(key, true);
            return true;
          }

          const cell = getCellCached(row, col);
          const value = cell?.value ?? null;
          const richTextText = cell?.richText?.text;
          const blocked =
            Boolean(cell?.image) ||
            (value !== null && value !== "") ||
            (typeof richTextText === "string" && richTextText !== "");
          this.frameBlockedCache.set(key, blocked);
          return blocked;
        }
      : (row, col) => {
          let rowCache = this.frameBlockedCacheNested.get(row);
          if (!rowCache) {
            rowCache = new Map();
            this.frameBlockedCacheNested.set(row, rowCache);
            this.__testOnly_rowCacheMapAllocs += 1;
          }
          if (rowCache.has(col)) return rowCache.get(col) ?? false;

          if (mergedRanges.length > 0 && mergedRangeAt(row, col)) {
            rowCache.set(col, true);
            return true;
          }

          const cell = getCellCached(row, col);
          const value = cell?.value ?? null;
          const richTextText = cell?.richText?.text;
          const blocked =
            Boolean(cell?.image) ||
            (value !== null && value !== "") ||
            (typeof richTextText === "string" && richTextText !== "");
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
      /**
       * Optional row range to use when probing for horizontal text overflow.
       *
       * For extremely tall merged cells, scanning the full merged height can be
       * prohibitively expensive. Since rendering is viewport-clipped, we can
       * safely limit overflow probing to the rows that are actually visible in
       * the current paint pass.
       */
      probeStartRow?: number;
      probeEndRow?: number;
      isHeader: boolean;
    }): void => {
      const { cell, x, y, width, height, spanStartRow, spanEndRow, spanStartCol, spanEndCol, probeStartRow, probeEndRow, isHeader } =
        options;
      const style = cell.style;

      const drawCommentIndicator = (): void => {
        if (!cell.comment) return;
        const resolved = cell.comment.resolved ?? false;
        const maxSize = Math.min(width, height);
        const size = Math.min(maxSize, Math.max(6, maxSize * 0.25));
        if (size <= 0) return;
        contentCtx.save();
        contentCtx.beginPath();
        contentCtx.moveTo(x + width, y);
        contentCtx.lineTo(x + width - size, y);
        contentCtx.lineTo(x + width, y + size);
        contentCtx.closePath();
        contentCtx.fillStyle = resolved ? commentIndicatorResolved : commentIndicator;
        contentCtx.fill();
        contentCtx.restore();
      };

      const image = cell.image;
      const imageId = typeof image?.imageId === "string" ? image.imageId.trim() : "";
      if (image && imageId !== "") {
        const bitmap = this.getOrRequestImageBitmap(imageId);

        const availableWidth = Math.max(0, width - paddingX * 2);
        const availableHeight = Math.max(0, height - paddingY * 2);

        const normalizeDim = (value: unknown): number | null => {
          if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return null;
          return value;
        };

        const bitmapWidth =
          bitmap && typeof (bitmap as any).width === "number"
            ? (bitmap as any).width
            : bitmap && typeof (bitmap as any).naturalWidth === "number"
              ? (bitmap as any).naturalWidth
              : null;
        const bitmapHeight =
          bitmap && typeof (bitmap as any).height === "number"
            ? (bitmap as any).height
            : bitmap && typeof (bitmap as any).naturalHeight === "number"
              ? (bitmap as any).naturalHeight
              : null;

        const srcW = normalizeDim(image.width) ?? normalizeDim(bitmapWidth) ?? null;
        const srcH = normalizeDim(image.height) ?? normalizeDim(bitmapHeight) ?? null;

        let destW = availableWidth;
        let destH = availableHeight;
        if (srcW && srcH && availableWidth > 0 && availableHeight > 0) {
          const scale = Math.min(availableWidth / srcW, availableHeight / srcH);
          destW = srcW * scale;
          destH = srcH * scale;
        }

        const destX = x + paddingX + Math.max(0, (availableWidth - destW) / 2);
        const destY = y + paddingY + Math.max(0, (availableHeight - destH) / 2);

        contentCtx.save();
        contentCtx.beginPath();
        contentCtx.rect(x, y, width, height);
        contentCtx.clip();

        if (bitmap && destW > 0 && destH > 0) {
          contentCtx.imageSmoothingEnabled = true;
          contentCtx.drawImage(bitmap, destX, destY, destW, destH);
        } else if (availableWidth > 0 && availableHeight > 0) {
          // Placeholder: checkerboard + alt text.
          const pattern = this.getImagePlaceholderPattern(contentCtx, zoom);
          if (pattern) {
            contentCtx.fillStyle = pattern;
          } else {
            contentCtx.fillStyle = "#d0d0d0";
          }
          contentCtx.fillRect(destX, destY, destW, destH);

          const rawLabel = image.altText ?? (typeof cell.value === "string" ? cell.value : "");
          const trimmedLabel = rawLabel.trim();
          const label = trimmedLabel !== "" ? trimmedLabel : "[Image]";
          const fontSize = Math.max(10, Math.round(11 * zoom));
          const placeholderFontFamily = ensureSansSerifFallback(
            isHeader ? this.defaultHeaderFontFamily : this.defaultCellFontFamily
          );
          contentCtx.font = `${fontSize}px ${placeholderFontFamily}`;
          contentCtx.fillStyle = "#333333";
          contentCtx.textAlign = "center";
          contentCtx.textBaseline = "middle";
          contentCtx.fillText(label, destX + destW / 2, destY + destH / 2);
          contentCtx.textAlign = "left";
          contentCtx.textBaseline = "alphabetic";
        }

        contentCtx.restore();
        drawCommentIndicator();
        return;
      }

      const richText = cell.richText;
      const richTextText = richText?.text ?? "";
      const hasRichText = Boolean(richText && richTextText);
      const hasValue = cell.value !== null;

      if (hasValue || hasRichText) {
        const fontVariantPosition = style?.fontVariantPosition;
        const fontVariantScale = fontVariantPosition === "subscript" || fontVariantPosition === "superscript" ? 0.75 : 1;
        const baseFontSize = (style?.fontSize ?? 12) * zoom;
        const fontSize = baseFontSize * fontVariantScale;
        const fontFamily = style?.fontFamily ?? (isHeader ? this.defaultHeaderFontFamily : this.defaultCellFontFamily);
        const fontWeight = style?.fontWeight ?? "400";
        const fontStyle = style?.fontStyle ?? "normal";
        const baselineShiftYBase =
          fontVariantPosition === "superscript"
            ? -Math.round(baseFontSize * 0.25)
            : fontVariantPosition === "subscript"
              ? Math.round(baseFontSize * 0.25)
              : 0;

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
          const horizontalAlign = style?.horizontalAlign;
          const baselineShiftY = rotationDeg === 0 ? baselineShiftYBase : 0;

          const availableWidth = Math.max(0, width - paddingX * 2);
          const availableHeight = Math.max(0, height - paddingY * 2);

          const align: CanvasTextAlign = style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
          const layoutAlign =
            align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
              ? (align as "left" | "right" | "center" | "start" | "end")
              : "start";

          const baseDirection = direction === "auto" ? detectBaseDirection(text) : direction;
          const resolvedAlign =
            layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
              ? layoutAlign
              : resolveAlign(layoutAlign, baseDirection);
          const textIndentX = rotationDeg === 0 ? resolveTextIndentPx(style?.textIndentPx, zoom) : 0;
          const indentX = resolvedAlign === "left" || resolvedAlign === "right" ? textIndentX : 0;
          const maxWidth = Math.max(0, availableWidth - indentX);
          const originX = x + paddingX + (resolvedAlign === "left" ? indentX : 0);

          const hasExplicitNewline = EXPLICIT_NEWLINE_RE.test(text);
          const rotationRad = (rotationDeg * Math.PI) / 180;

          type DecorationSegment = { x1: number; x2: number; y: number; color: string; lineWidth: number };

          const alignDecorationY = (pos: number, lineWidth: number): number => {
            // When rotated, the decoration is no longer axis-aligned, so pixel snapping doesn't help.
            if (rotationDeg !== 0) return pos;
            if (lineWidth === 1) return crispLine(pos);
            return pos;
          };

          const drawDecorationSegments = (segments: DecorationSegment[]): void => {
            if (segments.length === 0) return;
            // Ensure borders/dashes from other callers do not leak into text decorations.
            (contentCtx as any).setLineDash?.([]);

            type Bucket = { strokeStyle: string; lineWidth: number; coords: number[] };
            const buckets = new Map<string, Bucket>();

            for (const segment of segments) {
              const key = `${segment.color}|${segment.lineWidth}`;
              let bucket = buckets.get(key);
              if (!bucket) {
                bucket = { strokeStyle: segment.color, lineWidth: segment.lineWidth, coords: [] };
                buckets.set(key, bucket);
              }
              bucket.coords.push(segment.x1, segment.y, segment.x2, segment.y);
            }

            for (const bucket of buckets.values()) {
              contentCtx.strokeStyle = bucket.strokeStyle;
              contentCtx.lineWidth = bucket.lineWidth;
              contentCtx.beginPath();
              const coords = bucket.coords;
              for (let i = 0; i < coords.length; i += 4) {
                contentCtx.moveTo(coords[i]!, coords[i + 1]!);
                contentCtx.lineTo(coords[i + 2]!, coords[i + 3]!);
              }
              contentCtx.stroke();
            }
          };

          const offsets = buildCodePointIndex(text);
          const textLen = offsets.length - 1;
          const rawRuns = normalizeRichTextRuns(textLen, richText?.runs);

          const defaults = {
            family: fontFamily,
            sizePx: fontSize,
            weight: fontWeight,
            style: fontStyle
          } satisfies Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number; style: string };

          const defaultUnderline: UnderlineSpec =
            style?.underlineStyle === "double"
              ? { underline: true, underlineStyle: "double" }
              : style?.underline === true || style?.underlineStyle === "single"
                ? { underline: true, underlineStyle: undefined }
                : { underline: false, underlineStyle: undefined };
          const cellStrike = style?.strike === true;

          const layoutRuns = rawRuns.map((run) => {
            const runStyle =
              run.style && typeof run.style === "object" ? (run.style as RichTextRunStyle) : (undefined as RichTextRunStyle | undefined);
            const underlineSpec = resolveUnderline(runStyle, defaultUnderline);
            const baseRunFont = fontSpecForRichTextStyle(runStyle, defaults, zoom);
            const hasExplicitSize = typeof runStyle?.size_100pt === "number" && Number.isFinite(runStyle.size_100pt);
            return {
              text: sliceByCodePointRange(text, offsets, run.start, run.end),
              font:
                hasExplicitSize && fontVariantScale !== 1 ? { ...baseRunFont, sizePx: baseRunFont.sizePx * fontVariantScale } : baseRunFont,
              color: engineColorToCanvasColor(runStyle?.color),
              underline: underlineSpec.underline,
              underlineStyle: underlineSpec.underlineStyle,
              strike: resolveStrike(runStyle, cellStrike)
            };
          });

          const maxFontSizePx = layoutRuns.reduce((acc, run) => Math.max(acc, run.font.sizePx), defaults.sizePx);
          const lineHeight = Math.ceil(maxFontSizePx * 1.2);
          const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

          contentCtx.save();

          if (wrapMode === "none" && !hasExplicitNewline && rotationDeg === 0) {
            // Fast path: single-line rich text (no wrapping). Uses cached measurements.
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

            let cursorX = originX;
            if (resolvedAlign === "center") {
              cursorX = x + paddingX + (availableWidth - totalWidth) / 2;
            } else if (resolvedAlign === "right") {
              cursorX = x + paddingX + (maxWidth - totalWidth);
            }

            let baselineY = y + paddingY + lineAscent;
            if (verticalAlign === "middle") {
              baselineY = y + height / 2 + (lineAscent - lineDescent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + height - paddingY - lineDescent;
            }
            baselineY += baselineShiftY;

            const underlineSegments: DecorationSegment[] = [];
            const strikeSegments: DecorationSegment[] = [];

            const drawFragmentsAt = (startX: number, options?: { extraSpacePerAdjustableGap?: number; adjustableGapMask?: boolean[] }) => {
              const extraPerGap = options?.extraSpacePerAdjustableGap ?? 0;
              const adjustableMask = options?.adjustableGapMask ?? null;
              let xCursor = startX;
              let gapIdx = 0;
              for (const fragment of fragments) {
                contentCtx.font = toCanvasFontString(fragment.font);
                contentCtx.fillStyle = fragment.color ?? fillStyle;
                contentCtx.fillText(fragment.text, xCursor, baselineY);

                if (fragment.underline) {
                  const underlineOffset = Math.max(1, Math.round(fragment.font.sizePx * 0.08));
                  const decorationLineWidth = Math.max(1, Math.round(fragment.font.sizePx / 16));
                  const underlineY = alignDecorationY(baselineY + underlineOffset, decorationLineWidth);
                  underlineSegments.push({
                    x1: xCursor,
                    x2: xCursor + fragment.width,
                    y: underlineY,
                    color: contentCtx.fillStyle as string,
                    lineWidth: decorationLineWidth
                  });

                  if (fragment.underlineStyle === "double") {
                    const doubleGap = Math.max(2, Math.round(decorationLineWidth * 2));
                    const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap, decorationLineWidth);
                    underlineSegments.push({
                      x1: xCursor,
                      x2: xCursor + fragment.width,
                      y: underlineY2,
                      color: contentCtx.fillStyle as string,
                      lineWidth: decorationLineWidth
                    });
                  }
                }

                if (fragment.strike) {
                  const decorationLineWidth = Math.max(1, Math.round(fragment.font.sizePx / 16));
                  const strikeY = alignDecorationY(
                    baselineY - Math.max(1, Math.round(fragment.ascent * 0.3)),
                    decorationLineWidth
                  );
                  strikeSegments.push({
                    x1: xCursor,
                    x2: xCursor + fragment.width,
                    y: strikeY,
                    color: contentCtx.fillStyle as string,
                    lineWidth: decorationLineWidth
                  });
                }

                xCursor += fragment.width;
                if (extraPerGap !== 0) {
                  const add = adjustableMask ? (adjustableMask[gapIdx] ? extraPerGap : 0) : extraPerGap;
                  if (add) xCursor += add;
                  gapIdx += 1;
                }
              }
            };

            if (horizontalAlign === "fill" && totalWidth > 0 && maxWidth > 0) {
              // Excel-style "Fill" alignment: repeat the single-line rich text horizontally to fill
              // the available width. This is intentionally clipped to the cell rect (unlike the
              // default overflow behavior for long left/right-aligned text).
              const startX = originX;
              const maxRepeats = 512;
              const repeats = Math.min(maxRepeats, Math.max(1, Math.ceil(maxWidth / totalWidth)));

              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();

              for (let i = 0; i < repeats; i++) {
                drawFragmentsAt(startX + i * totalWidth);
              }

              contentCtx.restore();
            } else {
            const shouldClip = totalWidth > maxWidth;
            let clipX = x;
            let clipWidth = width;

            const rowProbeStart = probeStartRow ?? spanStartRow;
            const rowProbeEnd = probeEndRow ?? spanEndRow;

            if (
              shouldClip &&
              (resolvedAlign === "left" || resolvedAlign === "right") &&
              totalWidth > width - (paddingX + indentX)
            ) {
              const requiredExtra = paddingX + indentX + totalWidth - width;
              if (requiredExtra > 0) {
                if (resolvedAlign === "left") {
                  let extra = 0;
                  for (
                    let probeCol = spanEndCol, steps = 0;
                    probeCol < colCount && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                    probeCol++, steps++
                  ) {
                    let blocked = false;
                    for (let r = rowProbeStart; r < rowProbeEnd; r++) {
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
                    for (let r = rowProbeStart; r < rowProbeEnd; r++) {
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

            if (shouldClip) {
              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(clipX, y, clipWidth, height);
              contentCtx.clip();
              drawFragmentsAt(cursorX);
              contentCtx.restore();
            } else {
              drawFragmentsAt(cursorX);
            }
            }

            if (underlineSegments.length > 0 || strikeSegments.length > 0) {
              // Excel clips text decorations (underline/strike) to the cell's own rect, even when
              // the text itself overflows into adjacent empty cells.
              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();
              drawDecorationSegments(underlineSegments);
              drawDecorationSegments(strikeSegments);
              contentCtx.restore();
            }
          } else if (layoutEngine && maxWidth > 0) {
            const layout = layoutEngine.layout({
              runs: layoutRuns.map((r) => ({
                text: r.text,
                font: r.font,
                color: r.color,
                underline: r.underline,
                underlineStyle: r.underlineStyle,
                strike: r.strike
              })),
              text: undefined,
              font: defaults,
              maxWidth,
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

            const shouldClip = layout.width > maxWidth || layout.height > availableHeight || rotationDeg !== 0;

            const drawLayout = () => {
              const underlineSegments: DecorationSegment[] = [];
              const strikeSegments: DecorationSegment[] = [];

              const justifyEnabled =
                horizontalAlign === "justify" && rotationDeg === 0 && resolvedAlign === "left" && layout.lines.length > 1;
              const whitespaceRe = /^\s+$/;

              for (let i = 0; i < layout.lines.length; i++) {
                const line = layout.lines[i];
                const baselineY = originY + i * layout.lineHeight + line.ascent + baselineShiftY;

                // Preserve non-justified semantics for the last line (Excel-like).
                const justifyLine = justifyEnabled && i < layout.lines.length - 1 && line.width > 0 && maxWidth > line.width;

                if (justifyLine) {
                  const tokens: Array<{
                    text: string;
                    font: FontSpec;
                    color?: string;
                    underline?: boolean;
                    underlineStyle?: "single" | "double";
                    strike?: boolean;
                    isWhitespace: boolean;
                    adjustableGap: boolean;
                  }> = [];

                  for (const run of line.runs as Array<{
                    text: string;
                    font: FontSpec;
                    color?: string;
                    underline?: boolean;
                    underlineStyle?: "single" | "double";
                    strike?: boolean;
                  }>) {
                    if (!run.text) continue;
                    const parts = run.text.split(/(\s+)/).filter((p) => p.length > 0);
                    for (const part of parts) {
                      tokens.push({
                        text: part,
                        font: run.font,
                        color: run.color,
                        underline: run.underline,
                        underlineStyle: run.underlineStyle,
                        strike: run.strike,
                        isWhitespace: whitespaceRe.test(part),
                        adjustableGap: false
                      });
                    }
                  }

                  let gapCount = 0;
                  for (let ti = 0; ti < tokens.length; ti++) {
                    const tok = tokens[ti]!;
                    if (!tok.isWhitespace) continue;
                    if (ti === 0 || ti === tokens.length - 1) continue;
                    const left = tokens[ti - 1];
                    const right = tokens[ti + 1];
                    if (!left || !right) continue;
                    if (left.isWhitespace || right.isWhitespace) continue;
                    tok.adjustableGap = true;
                    gapCount += 1;
                  }

                  if (gapCount > 0) {
                    const extra = maxWidth - line.width;
                    const extraPerGap = extra > 0 ? extra / gapCount : 0;

                    let xCursor = originX;
                    for (const token of tokens) {
                      if (!token.text) continue;
                      const measurement = layoutEngine.measure(token.text, token.font);
                      const width = measurement?.width ?? contentCtx.measureText(token.text).width;
                      const ascent = measurement?.ascent ?? token.font.sizePx * 0.8;
                      const descent = measurement?.descent ?? token.font.sizePx * 0.2;
                      const extraAdvance = token.adjustableGap ? extraPerGap : 0;
                      const advance = token.isWhitespace ? width + extraAdvance : width;

                      contentCtx.font = toCanvasFontString(token.font);
                      contentCtx.fillStyle = token.color ?? fillStyle;

                      if (!token.isWhitespace) {
                        contentCtx.fillText(token.text, xCursor, baselineY);
                      }

                      if (token.underline) {
                        const underlineOffset = Math.max(1, Math.round(token.font.sizePx * 0.08));
                        const lineWidth = Math.max(1, Math.round(token.font.sizePx / 16));
                        const underlineY = alignDecorationY(baselineY + underlineOffset, lineWidth);
                        underlineSegments.push({
                          x1: xCursor,
                          x2: xCursor + advance,
                          y: underlineY,
                          color: contentCtx.fillStyle as string,
                          lineWidth
                        });

                        if (token.underlineStyle === "double") {
                          const doubleGap = Math.max(2, Math.round(lineWidth * 2));
                          const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap, lineWidth);
                          underlineSegments.push({
                            x1: xCursor,
                            x2: xCursor + advance,
                            y: underlineY2,
                            color: contentCtx.fillStyle as string,
                            lineWidth
                          });
                        }
                      }

                      if (token.strike) {
                        const lineWidth = Math.max(1, Math.round(token.font.sizePx / 16));
                        const strikeY = alignDecorationY(baselineY - Math.max(1, Math.round(ascent * 0.3)), lineWidth);
                        strikeSegments.push({
                          x1: xCursor,
                          x2: xCursor + advance,
                          y: strikeY,
                          color: contentCtx.fillStyle as string,
                          lineWidth
                        });
                      }

                      xCursor += advance;
                    }
                    continue;
                  }
                }

                let xCursor = originX + line.x;

                for (const run of line.runs as Array<{
                  text: string;
                  font: FontSpec;
                  color?: string;
                  underline?: boolean;
                  underlineStyle?: "single" | "double";
                  strike?: boolean;
                }>) {
                  const measurement = layoutEngine.measure(run.text, run.font);
                  contentCtx.font = toCanvasFontString(run.font);
                  contentCtx.fillStyle = run.color ?? fillStyle;
                  contentCtx.fillText(run.text, xCursor, baselineY);

                  if (run.underline) {
                    const underlineOffset = Math.max(1, Math.round(run.font.sizePx * 0.08));
                    const lineWidth = Math.max(1, Math.round(run.font.sizePx / 16));
                    const underlineY = alignDecorationY(baselineY + underlineOffset, lineWidth);
                    underlineSegments.push({
                      x1: xCursor,
                      x2: xCursor + measurement.width,
                      y: underlineY,
                      color: contentCtx.fillStyle as string,
                      lineWidth
                    });

                    if (run.underlineStyle === "double") {
                      const doubleGap = Math.max(2, Math.round(lineWidth * 2));
                      const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap, lineWidth);
                      underlineSegments.push({
                        x1: xCursor,
                        x2: xCursor + measurement.width,
                        y: underlineY2,
                        color: contentCtx.fillStyle as string,
                        lineWidth
                      });
                    }
                  }

                  if (run.strike) {
                    const lineWidth = Math.max(1, Math.round(run.font.sizePx / 16));
                    const strikeY = alignDecorationY(
                      baselineY - Math.max(1, Math.round(measurement.ascent * 0.3)),
                      lineWidth
                    );
                    strikeSegments.push({
                      x1: xCursor,
                      x2: xCursor + measurement.width,
                      y: strikeY,
                      color: contentCtx.fillStyle as string,
                      lineWidth
                    });
                  }

                  xCursor += measurement.width;
                }
              }

              return { underlineSegments, strikeSegments };
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

              const { underlineSegments, strikeSegments } = drawLayout();
              if (underlineSegments.length > 0 || strikeSegments.length > 0) {
                drawDecorationSegments(underlineSegments);
                drawDecorationSegments(strikeSegments);
              }
              contentCtx.restore();
            } else {
              const { underlineSegments, strikeSegments } = drawLayout();
              if (underlineSegments.length > 0 || strikeSegments.length > 0) {
                contentCtx.save();
                contentCtx.beginPath();
                contentCtx.rect(x, y, width, height);
                contentCtx.clip();
                drawDecorationSegments(underlineSegments);
                drawDecorationSegments(strikeSegments);
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
          const horizontalAlign = style?.horizontalAlign;
          const baselineShiftY = rotationDeg === 0 ? baselineShiftYBase : 0;
          const underlineStyle = style?.underlineStyle;
          const underline =
            style?.underline === true || underlineStyle === "single" || underlineStyle === "double";
          const doubleUnderline = underlineStyle === "double";
          const strike = style?.strike === true;
          const hasTextDecorations = underline || strike;
          const decorationLineWidth = hasTextDecorations ? Math.max(1, Math.round(fontSize / 12)) : 0;
          const alignDecorationY = (pos: number): number => {
            if (rotationDeg !== 0) return pos;
            if (decorationLineWidth === 1) return crispLine(pos);
            return pos;
          };
          const prepareDecorationStroke = () => {
            if (!hasTextDecorations) return;
            if (contentCtx.strokeStyle !== currentTextFill) {
              contentCtx.strokeStyle = currentTextFill;
            }
            contentCtx.lineWidth = decorationLineWidth;
            // Ensure borders/dashes from other callers do not leak into text decorations.
            (contentCtx as any).setLineDash?.([]);
          };

          const availableWidth = Math.max(0, width - paddingX * 2);
          const availableHeight = Math.max(0, height - paddingY * 2);
          const lineHeight = Math.ceil(fontSize * 1.2);
          const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

          const align: CanvasTextAlign = style?.textAlign ?? (typeof cell.value === "number" ? "end" : "start");
          const layoutAlign =
            align === "left" || align === "right" || align === "center" || align === "start" || align === "end"
              ? (align as "left" | "right" | "center" | "start" | "end")
              : "start";

          const baseDirection =
            direction === "auto" ? (typeof cell.value === "string" ? detectBaseDirection(text) : "ltr") : direction;
          const resolvedAlign =
            layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
              ? layoutAlign
              : resolveAlign(layoutAlign, baseDirection);
          const textIndentX = rotationDeg === 0 ? resolveTextIndentPx(style?.textIndentPx, zoom) : 0;
          const indentX = resolvedAlign === "left" || resolvedAlign === "right" ? textIndentX : 0;
          const maxWidth = Math.max(0, availableWidth - indentX);
          const originX = x + paddingX + (resolvedAlign === "left" ? indentX : 0);

          const hasExplicitNewline = EXPLICIT_NEWLINE_RE.test(text);
          const rotationRad = (rotationDeg * Math.PI) / 180;

          if (wrapMode === "none" && !hasExplicitNewline && rotationDeg === 0) {
            const measurement = layoutEngine?.measure(text, fontSpec);
            const textWidth = measurement?.width ?? contentCtx.measureText(text).width;
            const ascent = measurement?.ascent ?? fontSize * 0.8;
            const descent = measurement?.descent ?? fontSize * 0.2;

            let textX = originX;
            if (resolvedAlign === "center") {
              textX = x + paddingX + (availableWidth - textWidth) / 2;
            } else if (resolvedAlign === "right") {
              textX = x + paddingX + (maxWidth - textWidth);
            }

            let baselineY = y + paddingY + ascent;
            if (verticalAlign === "middle") {
              baselineY = y + height / 2 + (ascent - descent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + height - paddingY - descent;
            }
            baselineY += baselineShiftY;

            if (horizontalAlign === "fill" && textWidth > 0 && maxWidth > 0) {
              // Excel-style "Fill" alignment: repeat the single-line text horizontally to fill
              // the available width. This is intentionally clipped to the cell rect (unlike the
              // default overflow behavior for long left/right-aligned text).
              const startX = originX;
              const maxRepeats = 512;
              const repeats = Math.min(maxRepeats, Math.max(1, Math.ceil(maxWidth / textWidth)));

              contentCtx.save();
              contentCtx.beginPath();
              contentCtx.rect(x, y, width, height);
              contentCtx.clip();

              for (let i = 0; i < repeats; i++) {
                contentCtx.fillText(text, startX + i * textWidth, baselineY);
              }

              if (hasTextDecorations) {
                prepareDecorationStroke();
                const repeatedWidth = repeats * textWidth;

                if (underline) {
                  const underlineOffset = Math.max(1, Math.round(descent * 0.5));
                  const underlineY = alignDecorationY(baselineY + underlineOffset);
                  contentCtx.beginPath();
                  contentCtx.moveTo(startX, underlineY);
                  contentCtx.lineTo(startX + repeatedWidth, underlineY);
                  if (doubleUnderline) {
                    const doubleGap = Math.max(2, Math.round(decorationLineWidth * 2));
                    const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap);
                    contentCtx.moveTo(startX, underlineY2);
                    contentCtx.lineTo(startX + repeatedWidth, underlineY2);
                  }
                  contentCtx.stroke();
                }

                if (strike) {
                  const strikeY = alignDecorationY(baselineY - Math.max(1, Math.round(ascent * 0.3)));
                  contentCtx.beginPath();
                  contentCtx.moveTo(startX, strikeY);
                  contentCtx.lineTo(startX + repeatedWidth, strikeY);
                  contentCtx.stroke();
                }
              }

              contentCtx.restore();
            } else {
              const shouldClip = textWidth > maxWidth;
              if (shouldClip) {
                let clipX = x;
                let clipWidth = width;

                const rowProbeStart = probeStartRow ?? spanStartRow;
                const rowProbeEnd = probeEndRow ?? spanEndRow;

              if ((resolvedAlign === "left" || resolvedAlign === "right") && textWidth > width - (paddingX + indentX)) {
                const requiredExtra = paddingX + indentX + textWidth - width;
                if (requiredExtra > 0) {
                  if (resolvedAlign === "left") {
                    let extra = 0;
                    for (
                      let probeCol = spanEndCol, steps = 0;
                      probeCol < colCount && steps < MAX_TEXT_OVERFLOW_COLUMNS && extra < requiredExtra;
                      probeCol++, steps++
                    ) {
                      let blocked = false;
                      for (let r = rowProbeStart; r < rowProbeEnd; r++) {
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
                      for (let r = rowProbeStart; r < rowProbeEnd; r++) {
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
              if (hasTextDecorations && textWidth > 0) {
                prepareDecorationStroke();
                if (underline) {
                  const underlineOffset = Math.max(1, Math.round(descent * 0.5));
                  const underlineY = alignDecorationY(baselineY + underlineOffset);
                  contentCtx.beginPath();
                  contentCtx.moveTo(textX, underlineY);
                  contentCtx.lineTo(textX + textWidth, underlineY);
                  if (doubleUnderline) {
                    const doubleGap = Math.max(2, Math.round(decorationLineWidth * 2));
                    const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap);
                    contentCtx.moveTo(textX, underlineY2);
                    contentCtx.lineTo(textX + textWidth, underlineY2);
                  }
                  contentCtx.stroke();
                }
                if (strike) {
                  const strikeY = alignDecorationY(baselineY - Math.max(1, Math.round(ascent * 0.3)));
                  contentCtx.beginPath();
                  contentCtx.moveTo(textX, strikeY);
                  contentCtx.lineTo(textX + textWidth, strikeY);
                  contentCtx.stroke();
                }
              }
              contentCtx.restore();
            } else {
              contentCtx.fillText(text, textX, baselineY);
              if (hasTextDecorations && textWidth > 0) {
                contentCtx.save();
                contentCtx.beginPath();
                contentCtx.rect(x, y, width, height);
                contentCtx.clip();
                prepareDecorationStroke();
                if (underline) {
                  const underlineOffset = Math.max(1, Math.round(descent * 0.5));
                  const underlineY = alignDecorationY(baselineY + underlineOffset);
                  contentCtx.beginPath();
                  contentCtx.moveTo(textX, underlineY);
                  contentCtx.lineTo(textX + textWidth, underlineY);
                  if (doubleUnderline) {
                    const doubleGap = Math.max(2, Math.round(decorationLineWidth * 2));
                    const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap);
                    contentCtx.moveTo(textX, underlineY2);
                    contentCtx.lineTo(textX + textWidth, underlineY2);
                  }
                  contentCtx.stroke();
                }
                if (strike) {
                  const strikeY = alignDecorationY(baselineY - Math.max(1, Math.round(ascent * 0.3)));
                  contentCtx.beginPath();
                  contentCtx.moveTo(textX, strikeY);
                  contentCtx.lineTo(textX + textWidth, strikeY);
                  contentCtx.stroke();
                }
                contentCtx.restore();
              }
            }
            }
          } else if (layoutEngine && maxWidth > 0) {
            const layout = layoutEngine.layout({
              text,
              font: fontSpec,
              maxWidth,
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

            const justifyEnabled =
              horizontalAlign === "justify" && rotationDeg === 0 && resolvedAlign === "left" && layout.lines.length > 1;
            const justifiedLineWidths: number[] | null = justifyEnabled ? new Array(layout.lines.length) : null;

              const drawLayoutText = () => {
                if (!justifyEnabled) {
                  drawTextLayout(contentCtx, layout, originX, originY + baselineShiftY);
                  return;
                }

              const whitespaceRe = /^\s+$/;
              const measureWidth = (token: string): number => {
                const measured = layoutEngine.measure(token, fontSpec);
                return measured?.width ?? contentCtx.measureText(token).width;
              };

              for (let i = 0; i < layout.lines.length; i++) {
                const line = layout.lines[i];
                const baselineY = originY + i * layout.lineHeight + line.ascent + baselineShiftY;

                // Preserve non-justified semantics for the last line.
                if (i === layout.lines.length - 1) {
                  if (justifiedLineWidths) justifiedLineWidths[i] = line.width;
                  contentCtx.fillText(line.text, originX + line.x, baselineY);
                  continue;
                }

                const extra = maxWidth - line.width;
                if (!(extra > 0) || !line.text) {
                  if (justifiedLineWidths) justifiedLineWidths[i] = line.width;
                  contentCtx.fillText(line.text, originX + line.x, baselineY);
                  continue;
                }

                const tokens = line.text.split(/(\s+)/).filter((t) => t.length > 0);
                if (tokens.length === 0) {
                  if (justifiedLineWidths) justifiedLineWidths[i] = line.width;
                  continue;
                }

                let gapCount = 0;
                const adjustableGap = new Array(tokens.length).fill(false);
                for (let ti = 0; ti < tokens.length; ti++) {
                  const tok = tokens[ti];
                  if (!whitespaceRe.test(tok)) continue;
                  if (ti === 0 || ti === tokens.length - 1) continue;
                  const left = tokens[ti - 1];
                  const right = tokens[ti + 1];
                  if (whitespaceRe.test(left) || whitespaceRe.test(right)) continue;
                  adjustableGap[ti] = true;
                  gapCount += 1;
                }

                if (gapCount === 0) {
                  if (justifiedLineWidths) justifiedLineWidths[i] = line.width;
                  contentCtx.fillText(line.text, originX + line.x, baselineY);
                  continue;
                }

                const extraPerGap = extra / gapCount;
                if (justifiedLineWidths) justifiedLineWidths[i] = maxWidth;

                let xCursor = originX;
                for (let ti = 0; ti < tokens.length; ti++) {
                  const tok = tokens[ti];
                  if (!tok) continue;
                  const tokWidth = measureWidth(tok);

                  if (whitespaceRe.test(tok)) {
                    xCursor += tokWidth + (adjustableGap[ti] ? extraPerGap : 0);
                    continue;
                  }

                  contentCtx.fillText(tok, xCursor, baselineY);
                  xCursor += tokWidth;
                }
              }
            };

            const shouldClip = layout.width > maxWidth || layout.height > availableHeight || rotationDeg !== 0;
            const drawLayoutDecorations = () => {
              if (!hasTextDecorations) return;
              prepareDecorationStroke();
              if (underline) {
                const doubleGap = doubleUnderline ? Math.max(2, Math.round(decorationLineWidth * 2)) : 0;
                contentCtx.beginPath();
                for (let i = 0; i < layout.lines.length; i++) {
                  const line = layout.lines[i];
                  if (line.width <= 0) continue;
                  const x1 = originX + line.x;
                  const effectiveWidth = justifiedLineWidths?.[i] ?? line.width;
                  const x2 = x1 + effectiveWidth;
                  const baselineY = originY + i * layout.lineHeight + line.ascent + baselineShiftY;
                  const underlineOffset = Math.max(1, Math.round(line.descent * 0.5));
                  const underlineY = alignDecorationY(baselineY + underlineOffset);
                  contentCtx.moveTo(x1, underlineY);
                  contentCtx.lineTo(x2, underlineY);
                  if (doubleUnderline) {
                    const underlineY2 = alignDecorationY(baselineY + underlineOffset + doubleGap);
                    contentCtx.moveTo(x1, underlineY2);
                    contentCtx.lineTo(x2, underlineY2);
                  }
                }
                contentCtx.stroke();
              }
              if (strike) {
                contentCtx.beginPath();
                for (let i = 0; i < layout.lines.length; i++) {
                  const line = layout.lines[i];
                  if (line.width <= 0) continue;
                  const x1 = originX + line.x;
                  const effectiveWidth = justifiedLineWidths?.[i] ?? line.width;
                  const x2 = x1 + effectiveWidth;
                  const baselineY = originY + i * layout.lineHeight + line.ascent + baselineShiftY;
                  const strikeY = alignDecorationY(baselineY - Math.max(1, Math.round(line.ascent * 0.3)));
                  contentCtx.moveTo(x1, strikeY);
                  contentCtx.lineTo(x2, strikeY);
                }
                contentCtx.stroke();
              }
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

              drawLayoutText();
              drawLayoutDecorations();
              contentCtx.restore();
            } else {
              drawLayoutText();
              drawLayoutDecorations();
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

      drawCommentIndicator();
    };

    const diagonalEntries: Array<{ rect: Rect; up?: unknown; down?: unknown }> = [];

    // Render merged regions (fill + text) first so we can skip their constituent cells below.
    for (const range of mergedRanges) {
      if (range.startRow >= endRow || range.endRow <= startRow) continue;
      if (range.startCol >= endCol || range.endCol <= startCol) continue;
      hasMerges = true;
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
      const diagonalBorders = anchorStyle?.diagonalBorders;
      if (diagonalBorders && (diagonalBorders.up || diagonalBorders.down)) {
        diagonalEntries.push({ rect: { x, y, width, height }, up: diagonalBorders.up, down: diagonalBorders.down });
      }

      const fill = anchorStyle?.fill ?? (isHeader ? headerBg : undefined);
      const fillToDraw = fill && (fill !== gridBg || (hasBackgroundPattern && !isHeader)) ? fill : null;
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
          probeStartRow: Math.max(startRow, range.startRow),
          probeEndRow: Math.min(endRow, range.endRow),
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

        if (hasMerges && mergedRangeAt(row, col)) {
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
        const diagonalBorders = style?.diagonalBorders;
        if (diagonalBorders && (diagonalBorders.up || diagonalBorders.down)) {
          diagonalEntries.push({ rect: { x, y, width: colWidth, height: rowHeight }, up: diagonalBorders.up, down: diagonalBorders.down });
        }

        // Background fill (grid layer).
        const fill = style?.fill ?? (isHeader ? headerBg : undefined);
        const fillToDraw = fill && (fill !== gridBg || (hasBackgroundPattern && !isHeader)) ? fill : null;
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
    (gridCtx as any).setLineDash?.([]);

    const drawCellBorders = () => {
      // Excel-like border rendering:
      // - Borders are collapsed so shared edges are drawn once (deterministic conflict resolution).
      // - Double borders are rendered as two parallel strokes.
      // - Merged ranges are drawn as a single rect using the anchor cell's border specs.
      // - Segments are batched by stroke config (color/width/dash/double) to avoid per-segment strokes.

      const ORIENT_H = 0 as const;
      const ORIENT_V = 1 as const;
      type BorderOrientation = typeof ORIENT_H | typeof ORIENT_V;

      type BorderWinner = {
        spec: CellBorderSpec;
        totalWidth: number;
        styleRank: number;
        sourceRow: number;
        sourceCol: number;
        x1: number;
        y1: number;
        x2: number;
        y2: number;
        orientation: BorderOrientation;
      };

      const isCellBorder = (value: unknown): value is CellBorderSpec => {
        if (!value || typeof value !== "object") return false;
        const v = value as any;
        if (typeof v.color !== "string" || v.color.trim() === "") return false;
        if (typeof v.width !== "number" || !Number.isFinite(v.width) || v.width <= 0) return false;
        return v.style === "solid" || v.style === "dashed" || v.style === "dotted" || v.style === "double";
      };

      const styleRank = (style: CellBorderSpec["style"]): number => {
        // Deterministic style ordering when widths tie.
        // (Higher wins.)
        switch (style) {
          case "double":
            return 4;
          case "solid":
            return 3;
          case "dashed":
            return 2;
          case "dotted":
            return 1;
          default:
            return 0;
        }
      };

      const dashForStyle = (style: CellBorderSpec["style"], lineWidth: number): number[] => {
        if (style === "dashed") {
          const unit = Math.max(1, lineWidth);
          return [4 * unit, 3 * unit];
        }
        if (style === "dotted") {
          const unit = Math.max(1, lineWidth);
          return [unit, 2 * unit];
        }
        return [];
      };

      const alignPos = (pos: number, lineWidth: number): number => {
        // Align strokes to device pixels for crisp output when the width is effectively an integer.
        //
        // For non-integer widths (common at fractional zoom levels), snapping would shift borders
        // away from their true geometric position and can collapse distinct widths (e.g. double
        // border widths of 2 vs 3). In those cases we draw at the exact coordinate and let the
        // browser antialias as needed.
        //
        // Mirrors the logic asserted in `CanvasGridRenderer.cellFormatting.test.ts` for integer widths
        // (odd widths use half-pixel alignment).
        const roundedWidth = Math.round(lineWidth);
        if (Math.abs(lineWidth - roundedWidth) > 1e-3) return pos;
        const roundedPos = Math.round(pos);
        return roundedWidth % 2 === 1 ? roundedPos + 0.5 : roundedPos;
      };

      // Use numeric keys for edge winner maps to avoid allocating per-edge string keys.
      // Keys are relative to the current rendered window (startRow/startCol), so they stay small.
      const rowSpan = endRow - startRow;
      const colSpan = endCol - startCol;
      const horizontalKeySpace = (rowSpan + 1) * colSpan;
      const hKey = (rowBoundary: number, col: number): number => (rowBoundary - startRow) * colSpan + (col - startCol);
      const vKey = (row: number, colBoundary: number): number =>
        horizontalKeySpace + (row - startRow) * (colSpan + 1) + (colBoundary - startCol);

      const winners = new Map<number, BorderWinner>();
      const consider = (
        key: number,
        spec: CellBorderSpec,
        totalWidth: number,
        styleRankValue: number,
        sourceRow: number,
        sourceCol: number,
        x1: number,
        y1: number,
        x2: number,
        y2: number,
        orientation: BorderOrientation
      ) => {
        const existing = winners.get(key);
        if (!existing) {
          winners.set(key, { spec, totalWidth, styleRank: styleRankValue, sourceRow, sourceCol, x1, y1, x2, y2, orientation });
          return;
        }

        if (totalWidth !== existing.totalWidth) {
          if (totalWidth > existing.totalWidth) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
          return;
        }

        if (styleRankValue !== existing.styleRank) {
          if (styleRankValue > existing.styleRank) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
          return;
        }

        // Final tie-breaker: prefer the bottom/right cell (deterministic).
        const candidateTie = orientation === ORIENT_H ? sourceRow /* bottom cell */ : sourceCol /* right cell */;
        const existingTie = existing.orientation === ORIENT_H ? existing.sourceRow : existing.sourceCol;
        if (candidateTie !== existingTie) {
          if (candidateTie > existingTie) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
          return;
        }

        // Fully deterministic fallback (should rarely/never happen).
        if (sourceRow !== existing.sourceRow) {
          if (sourceRow > existing.sourceRow) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
          return;
        }
        if (sourceCol !== existing.sourceCol) {
          if (sourceCol > existing.sourceCol) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
          return;
        }
        if (spec.color !== existing.spec.color) {
          if (spec.color > existing.spec.color) {
            existing.spec = spec;
            existing.totalWidth = totalWidth;
            existing.styleRank = styleRankValue;
            existing.sourceRow = sourceRow;
            existing.sourceCol = sourceCol;
            existing.x1 = x1;
            existing.y1 = y1;
            existing.x2 = x2;
            existing.y2 = y2;
            existing.orientation = orientation;
          }
        }
      };

      const addBorder = (
        key: number,
        spec: CellBorderSpec,
        sourceRow: number,
        sourceCol: number,
        x1: number,
        y1: number,
        x2: number,
        y2: number,
        orientation: BorderOrientation
      ) => {
        const totalWidth = spec.width * zoom;
        if (!Number.isFinite(totalWidth) || totalWidth <= 0) return;
        consider(
          key,
          spec,
          totalWidth,
          styleRank(spec.style),
          sourceRow,
          sourceCol,
          x1,
          y1,
          x2,
          y2,
          orientation
        );
      };

      // 1) Collect per-cell borders (skip merged cells, which are handled separately below).
      let borderRowYSheet = startRowYSheet;
      for (let row = startRow; row < endRow; row++) {
        const rowHeight = rowAxis.getSize(row);
        const y = borderRowYSheet - quadrant.scrollBaseY + quadrant.originY;

        let borderColXSheet = startColXSheet;
        for (let col = startCol; col < endCol; col++) {
          const colWidth = colAxis.getSize(col);
          const x = borderColXSheet - quadrant.scrollBaseX + quadrant.originX;

          if (hasMerges && mergedRangeAt(row, col)) {
            borderColXSheet += colWidth;
            continue;
          }

          const cell = getCellCached(row, col);
          const borders = cell?.style?.borders;
          if (borders) {
            const top = borders.top;
            const right = borders.right;
            const bottom = borders.bottom;
            const left = borders.left;

            if (isCellBorder(top)) {
              addBorder(hKey(row, col), top, row, col, x, y, x + colWidth, y, ORIENT_H);
            }
            if (isCellBorder(bottom)) {
              addBorder(hKey(row + 1, col), bottom, row, col, x, y + rowHeight, x + colWidth, y + rowHeight, ORIENT_H);
            }
            if (isCellBorder(left)) {
              addBorder(vKey(row, col), left, row, col, x, y, x, y + rowHeight, ORIENT_V);
            }
            if (isCellBorder(right)) {
              addBorder(vKey(row, col + 1), right, row, col, x + colWidth, y, x + colWidth, y + rowHeight, ORIENT_V);
            }
          }

          borderColXSheet += colWidth;
        }

        borderRowYSheet += rowHeight;
      }

      // When frozen panes are enabled, the grid is rendered in quadrants with different scroll bases.
      // At scrollX/scrollY === 0, the frozen quadrants abut the scrollable quadrants in *sheet space*,
      // meaning cell borders can conflict on the shared frozen boundary (Excel-like collapsed borders).
      //
      // When scrollX/scrollY > 0, the boundary is a "jump" in sheet space (the scrolled-out region is not
      // visible), so borders from frozen and scrollable panes are not adjacent and should not be collapsed.
      if (scrollY === 0 && frozenRows > 0) {
        // Shared horizontal boundary between the last frozen row (frozenRows-1) and the first scrollable row (frozenRows).
        if (startRow === frozenRows && startRow > 0) {
          // We are rendering the scrollable pane directly below the frozen rows; include the row above's bottom borders.
          const rowAbove = startRow - 1;
          const yBoundary = startRowYSheet - quadrant.scrollBaseY + quadrant.originY;
          let xSheet = startColXSheet;
          for (let col = startCol; col < endCol; col++) {
            const colWidth = colAxis.getSize(col);
            const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
            if (mergedRanges.length > 0 && mergedRangeAt(rowAbove, col)) {
              xSheet += colWidth;
              continue;
            }
            const cell = getCellCached(rowAbove, col);
            const bottom = cell?.style?.borders?.bottom;
            if (isCellBorder(bottom)) {
              addBorder(hKey(startRow, col), bottom, rowAbove, col, x, yBoundary, x + colWidth, yBoundary, ORIENT_H);
            }
            xSheet += colWidth;
          }

          // Also include merged-range bottom borders that land exactly on the frozen boundary (ranges above the boundary).
          if (mergedRanges.length > 0) {
            for (const range of mergedRanges) {
              if (range.endRow !== startRow) continue;
              if (range.startRow >= startRow) continue;
              const segStartCol = Math.max(startCol, range.startCol);
              const segEndCol = Math.min(endCol, range.endCol);
              if (segStartCol >= segEndCol) continue;
              const anchorCell = getCellCached(range.startRow, range.startCol);
              const bottom = anchorCell?.style?.borders?.bottom;
              if (!isCellBorder(bottom)) continue;
              let mergeXSheet = colAxis.positionOf(segStartCol);
              for (let col = segStartCol; col < segEndCol; col++) {
                const colWidth = colAxis.getSize(col);
                const x = mergeXSheet - quadrant.scrollBaseX + quadrant.originX;
                addBorder(hKey(startRow, col), bottom, range.startRow, range.startCol, x, yBoundary, x + colWidth, yBoundary, ORIENT_H);
                mergeXSheet += colWidth;
              }
            }
          }
        } else if (endRow === frozenRows && endRow < rowCount) {
          // We are rendering the frozen pane; include the first scrollable row's top borders.
          const rowBelow = endRow;
          const yBoundary = borderRowYSheet - quadrant.scrollBaseY + quadrant.originY;
          let xSheet = startColXSheet;
          for (let col = startCol; col < endCol; col++) {
            const colWidth = colAxis.getSize(col);
            const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
            if (mergedRanges.length > 0 && mergedRangeAt(rowBelow, col)) {
              xSheet += colWidth;
              continue;
            }
            const cell = getCellCached(rowBelow, col);
            const top = cell?.style?.borders?.top;
            if (isCellBorder(top)) {
              addBorder(hKey(endRow, col), top, rowBelow, col, x, yBoundary, x + colWidth, yBoundary, ORIENT_H);
            }
            xSheet += colWidth;
          }

          // Also include merged-range top borders that land exactly on the frozen boundary (ranges below the boundary).
          if (mergedRanges.length > 0) {
            for (const range of mergedRanges) {
              if (range.startRow !== endRow) continue;
              if (range.endRow <= endRow) continue;
              const segStartCol = Math.max(startCol, range.startCol);
              const segEndCol = Math.min(endCol, range.endCol);
              if (segStartCol >= segEndCol) continue;
              const anchorCell = getCellCached(range.startRow, range.startCol);
              const top = anchorCell?.style?.borders?.top;
              if (!isCellBorder(top)) continue;
              let mergeXSheet = colAxis.positionOf(segStartCol);
              for (let col = segStartCol; col < segEndCol; col++) {
                const colWidth = colAxis.getSize(col);
                const x = mergeXSheet - quadrant.scrollBaseX + quadrant.originX;
                addBorder(hKey(endRow, col), top, range.startRow, range.startCol, x, yBoundary, x + colWidth, yBoundary, ORIENT_H);
                mergeXSheet += colWidth;
              }
            }
          }
        }
      }

      if (scrollX === 0 && frozenCols > 0) {
        // Shared vertical boundary between the last frozen col (frozenCols-1) and the first scrollable col (frozenCols).
        if (startCol === frozenCols && startCol > 0) {
          // We are rendering the scrollable pane to the right of the frozen cols; include the col to the left's right borders.
          const colLeft = startCol - 1;
          const xBoundary = startColXSheet - quadrant.scrollBaseX + quadrant.originX;
          let ySheet = startRowYSheet;
          for (let row = startRow; row < endRow; row++) {
            const rowHeight = rowAxis.getSize(row);
            const y = ySheet - quadrant.scrollBaseY + quadrant.originY;
            if (mergedRanges.length > 0 && mergedRangeAt(row, colLeft)) {
              ySheet += rowHeight;
              continue;
            }
            const cell = getCellCached(row, colLeft);
            const right = cell?.style?.borders?.right;
            if (isCellBorder(right)) {
              addBorder(vKey(row, startCol), right, row, colLeft, xBoundary, y, xBoundary, y + rowHeight, ORIENT_V);
            }
            ySheet += rowHeight;
          }

          // Also include merged-range right borders that land exactly on the frozen boundary (ranges left of the boundary).
          if (mergedRanges.length > 0) {
            for (const range of mergedRanges) {
              if (range.endCol !== startCol) continue;
              if (range.startCol >= startCol) continue;
              const segStartRow = Math.max(startRow, range.startRow);
              const segEndRow = Math.min(endRow, range.endRow);
              if (segStartRow >= segEndRow) continue;
              const anchorCell = getCellCached(range.startRow, range.startCol);
              const right = anchorCell?.style?.borders?.right;
              if (!isCellBorder(right)) continue;
              let mergeYSheet = rowAxis.positionOf(segStartRow);
              for (let row = segStartRow; row < segEndRow; row++) {
                const rowHeight = rowAxis.getSize(row);
                const y = mergeYSheet - quadrant.scrollBaseY + quadrant.originY;
                addBorder(vKey(row, startCol), right, range.startRow, range.startCol, xBoundary, y, xBoundary, y + rowHeight, ORIENT_V);
                mergeYSheet += rowHeight;
              }
            }
          }
        } else if (endCol === frozenCols && endCol < colCount) {
          // We are rendering the frozen pane; include the first scrollable col's left borders.
          const colRight = endCol;
          const xBoundary = colAxis.positionOf(endCol) - quadrant.scrollBaseX + quadrant.originX;
          let ySheet = startRowYSheet;
          for (let row = startRow; row < endRow; row++) {
            const rowHeight = rowAxis.getSize(row);
            const y = ySheet - quadrant.scrollBaseY + quadrant.originY;
            if (mergedRanges.length > 0 && mergedRangeAt(row, colRight)) {
              ySheet += rowHeight;
              continue;
            }
            const cell = getCellCached(row, colRight);
            const left = cell?.style?.borders?.left;
            if (isCellBorder(left)) {
              addBorder(vKey(row, endCol), left, row, colRight, xBoundary, y, xBoundary, y + rowHeight, ORIENT_V);
            }
            ySheet += rowHeight;
          }

          // Also include merged-range left borders that land exactly on the frozen boundary (ranges right of the boundary).
          if (mergedRanges.length > 0) {
            for (const range of mergedRanges) {
              if (range.startCol !== endCol) continue;
              if (range.endCol <= endCol) continue;
              const segStartRow = Math.max(startRow, range.startRow);
              const segEndRow = Math.min(endRow, range.endRow);
              if (segStartRow >= segEndRow) continue;
              const anchorCell = getCellCached(range.startRow, range.startCol);
              const left = anchorCell?.style?.borders?.left;
              if (!isCellBorder(left)) continue;
              let mergeYSheet = rowAxis.positionOf(segStartRow);
              for (let row = segStartRow; row < segEndRow; row++) {
                const rowHeight = rowAxis.getSize(row);
                const y = mergeYSheet - quadrant.scrollBaseY + quadrant.originY;
                addBorder(vKey(row, endCol), left, range.startRow, range.startCol, xBoundary, y, xBoundary, y + rowHeight, ORIENT_V);
                mergeYSheet += rowHeight;
              }
            }
          }
        }
      }

      // 2) Collect merged range borders by drawing the perimeter of the merged rect.
      if (hasMerges) {
        for (const range of mergedRanges) {
          if (range.startRow >= endRow || range.endRow <= startRow) continue;
          if (range.startCol >= endCol || range.endCol <= startCol) continue;
          const anchorRow = range.startRow;
          const anchorCol = range.startCol;
          const anchorCell = getCellCached(anchorRow, anchorCol);
          const borders = anchorCell?.style?.borders;
          if (!borders) continue;

          const top = borders.top;
          const right = borders.right;
          const bottom = borders.bottom;
          const left = borders.left;

          const yTop = rowAxis.positionOf(range.startRow) - quadrant.scrollBaseY + quadrant.originY;
          const yBottom = rowAxis.positionOf(range.endRow) - quadrant.scrollBaseY + quadrant.originY;
          const xLeft = colAxis.positionOf(range.startCol) - quadrant.scrollBaseX + quadrant.originX;
          const xRight = colAxis.positionOf(range.endCol) - quadrant.scrollBaseX + quadrant.originX;

          // Top edge (split per column).
          if (isCellBorder(top) && range.startRow >= startRow && range.startRow <= endRow) {
            const segStartCol = Math.max(startCol, range.startCol);
            const segEndCol = Math.min(endCol, range.endCol);
            let xSheet = colAxis.positionOf(segStartCol);
            for (let col = segStartCol; col < segEndCol; col++) {
              const colWidth = colAxis.getSize(col);
              const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
              addBorder(hKey(range.startRow, col), top, anchorRow, anchorCol, x, yTop, x + colWidth, yTop, ORIENT_H);
              xSheet += colWidth;
            }
          }

          // Bottom edge (split per column).
          if (isCellBorder(bottom) && range.endRow >= startRow && range.endRow <= endRow) {
            const segStartCol = Math.max(startCol, range.startCol);
            const segEndCol = Math.min(endCol, range.endCol);
            let xSheet = colAxis.positionOf(segStartCol);
            for (let col = segStartCol; col < segEndCol; col++) {
              const colWidth = colAxis.getSize(col);
              const x = xSheet - quadrant.scrollBaseX + quadrant.originX;
              addBorder(hKey(range.endRow, col), bottom, anchorRow, anchorCol, x, yBottom, x + colWidth, yBottom, ORIENT_H);
              xSheet += colWidth;
            }
          }

          // Left edge (split per row).
          if (isCellBorder(left) && range.startCol >= startCol && range.startCol <= endCol) {
            const segStartRow = Math.max(startRow, range.startRow);
            const segEndRow = Math.min(endRow, range.endRow);
            let ySheet = rowAxis.positionOf(segStartRow);
            for (let row = segStartRow; row < segEndRow; row++) {
              const rowHeight = rowAxis.getSize(row);
              const y = ySheet - quadrant.scrollBaseY + quadrant.originY;
              addBorder(vKey(row, range.startCol), left, anchorRow, anchorCol, xLeft, y, xLeft, y + rowHeight, ORIENT_V);
              ySheet += rowHeight;
            }
          }

          // Right edge (split per row).
          if (isCellBorder(right) && range.endCol >= startCol && range.endCol <= endCol) {
            const segStartRow = Math.max(startRow, range.startRow);
            const segEndRow = Math.min(endRow, range.endRow);
            let ySheet = rowAxis.positionOf(segStartRow);
            for (let row = segStartRow; row < segEndRow; row++) {
              const rowHeight = rowAxis.getSize(row);
              const y = ySheet - quadrant.scrollBaseY + quadrant.originY;
              addBorder(vKey(row, range.endCol), right, anchorRow, anchorCol, xRight, y, xRight, y + rowHeight, ORIENT_V);
              ySheet += rowHeight;
            }
          }
        }
      }

      if (winners.size === 0) return;

      // 3) Batch draw by stroke config.
      type StrokeGroup = {
        strokeStyle: string;
        lineWidth: number;
        lineDash: number[];
        dashKey: string;
        lineCap: CanvasLineCap;
        double: boolean;
        segments: BorderWinner[];
      };

      const groups = new Map<string, StrokeGroup>();
      for (const winner of winners.values()) {
        const strokeStyle = winner.spec.color;
        if (winner.spec.style === "double") {
          const lineWidth = winner.totalWidth / 3;
          if (!Number.isFinite(lineWidth) || lineWidth <= 0) continue;
          const dashKey = "solid";
          const key = `${strokeStyle}|${lineWidth}|double`;
          let group = groups.get(key);
          if (!group) {
            group = { strokeStyle, lineWidth, lineDash: [], dashKey, lineCap: "butt", double: true, segments: [] };
            groups.set(key, group);
          }
          group.segments.push(winner);
        } else {
          const lineWidth = winner.totalWidth;
          if (!Number.isFinite(lineWidth) || lineWidth <= 0) continue;
          const style = winner.spec.style;
          const unit = style === "solid" ? 0 : Math.max(1, lineWidth);
          // Include the resolved dash unit in the key so we don't need to allocate a dash array just
          // to determine whether `setLineDash` needs to change.
          const dashKey = style === "solid" ? "solid" : `${style}:${unit}`;
          const key = `${strokeStyle}|${lineWidth}|${dashKey}|single`;
          let group = groups.get(key);
          if (!group) {
            const lineDash = style === "solid" ? [] : dashForStyle(style, lineWidth);
            group = {
              strokeStyle,
              lineWidth,
              lineDash,
              dashKey,
              lineCap: style === "dotted" ? "round" : "butt",
              double: false,
              segments: []
            };
            groups.set(key, group);
          }
          group.segments.push(winner);
        }
      }

      if (groups.size === 0) return;

      const orderedGroups = [...groups.values()].sort((a, b) => {
        if (a.dashKey !== b.dashKey) return a.dashKey < b.dashKey ? -1 : 1;
        if (a.strokeStyle !== b.strokeStyle) return a.strokeStyle < b.strokeStyle ? -1 : 1;
        if (a.lineWidth !== b.lineWidth) return a.lineWidth - b.lineWidth;
        if (a.double !== b.double) return a.double ? 1 : -1;
        return 0;
      });

      let currentStrokeStyle = "";
      let currentLineWidth = -1;
      let currentDashKey = "";
      let currentLineCap: CanvasLineCap = "butt";

      for (const group of orderedGroups) {
        if (group.strokeStyle !== currentStrokeStyle) {
          gridCtx.strokeStyle = group.strokeStyle;
          currentStrokeStyle = group.strokeStyle;
        }
        if (group.lineWidth !== currentLineWidth) {
          gridCtx.lineWidth = group.lineWidth;
          currentLineWidth = group.lineWidth;
        }
        if (group.lineCap !== currentLineCap) {
          gridCtx.lineCap = group.lineCap;
          currentLineCap = group.lineCap;
        }
        if (group.dashKey !== currentDashKey) {
          (gridCtx as any).setLineDash?.(group.lineDash);
          currentDashKey = group.dashKey;
        }

        if (!group.double) {
          gridCtx.beginPath();
          for (const seg of group.segments) {
            if (seg.orientation === ORIENT_H) {
              const y = alignPos(seg.y1, group.lineWidth);
              gridCtx.moveTo(seg.x1, y);
              gridCtx.lineTo(seg.x2, y);
            } else {
              const x = alignPos(seg.x1, group.lineWidth);
              gridCtx.moveTo(x, seg.y1);
              gridCtx.lineTo(x, seg.y2);
            }
          }
          gridCtx.stroke();
          continue;
        }

        const offset = group.lineWidth;
        gridCtx.beginPath();
        for (const seg of group.segments) {
          if (seg.orientation === ORIENT_H) {
            const y1 = alignPos(seg.y1 - offset, group.lineWidth);
            const y2 = alignPos(seg.y1 + offset, group.lineWidth);
            gridCtx.moveTo(seg.x1, y1);
            gridCtx.lineTo(seg.x2, y1);
            gridCtx.moveTo(seg.x1, y2);
            gridCtx.lineTo(seg.x2, y2);
          } else {
            const x1 = alignPos(seg.x1 - offset, group.lineWidth);
            const x2 = alignPos(seg.x1 + offset, group.lineWidth);
            gridCtx.moveTo(x1, seg.y1);
            gridCtx.lineTo(x1, seg.y2);
            gridCtx.moveTo(x2, seg.y1);
            gridCtx.lineTo(x2, seg.y2);
          }
        }
        gridCtx.stroke();
      }

      // Avoid leaking dashed patterns / round caps into diagonal borders (drawn next).
      (gridCtx as any).setLineDash?.([]);
      gridCtx.lineCap = "butt";
    };

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
    } else {
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

    // Explicit cell borders (grid layer).
    drawCellBorders();

    // Diagonal borders (grid layer), drawn last so they render above gridlines + side borders (Excel-like).
    if (diagonalEntries.length === 0) return;

    type DiagonalSegment = { x1: number; y1: number; x2: number; y2: number };
    type DiagonalStrokeGroup = {
      strokeStyle: string;
      lineWidth: number;
      lineDash: number[];
      lineCap: CanvasLineCap;
      clipRects: Rect[];
      clipRectKeys: Set<string>;
      segments: DiagonalSegment[];
    };

    const isCellBorder = (value: unknown): value is CellBorderSpec => {
      if (!value || typeof value !== "object") return false;
      const v = value as any;
      if (typeof v.color !== "string" || v.color.trim() === "") return false;
      if (typeof v.width !== "number" || !Number.isFinite(v.width) || v.width <= 0) return false;
      return v.style === "solid" || v.style === "dashed" || v.style === "dotted" || v.style === "double";
    };

    const dashForStyle = (style: CellBorderSpec["style"], lineWidth: number): number[] => {
      if (style === "dashed") {
        const unit = Math.max(1, lineWidth);
        return [4 * unit, 3 * unit];
      }
      if (style === "dotted") {
        const unit = Math.max(1, lineWidth);
        return [unit, 2 * unit];
      }
      return [];
    };

    const groups = new Map<string, DiagonalStrokeGroup>();

    const addGroupSegment = (options: {
      rect: Rect;
      strokeStyle: string;
      lineWidth: number;
      lineDash: number[];
      lineCap: CanvasLineCap;
      segment: DiagonalSegment;
    }): void => {
      const dashKey = options.lineDash.length > 0 ? options.lineDash.join(",") : "solid";
      const key = `${options.strokeStyle}|${options.lineWidth}|${dashKey}|${options.lineCap}`;
      let group = groups.get(key);
      if (!group) {
        group = {
          strokeStyle: options.strokeStyle,
          lineWidth: options.lineWidth,
          lineDash: options.lineDash,
          lineCap: options.lineCap,
          clipRects: [],
          clipRectKeys: new Set<string>(),
          segments: []
        };
        groups.set(key, group);
      }
      const rectKey = `${options.rect.x},${options.rect.y},${options.rect.width},${options.rect.height}`;
      if (!group.clipRectKeys.has(rectKey)) {
        group.clipRectKeys.add(rectKey);
        group.clipRects.push(options.rect);
      }
      group.segments.push(options.segment);
    };

    const addDiagonal = (rect: Rect, border: unknown, direction: "up" | "down"): void => {
      if (!isCellBorder(border)) return;
      const strokeStyle = border.color;
      const totalWidth = border.width * zoom;
      if (!Number.isFinite(totalWidth) || totalWidth <= 0) return;
      const lineCap: CanvasLineCap = border.style === "dotted" ? "round" : "butt";

      const x1 = rect.x;
      const y1 = direction === "up" ? rect.y + rect.height : rect.y;
      const x2 = rect.x + rect.width;
      const y2 = direction === "up" ? rect.y : rect.y + rect.height;

      if (border.style === "double") {
        const lineWidth = totalWidth / 3;
        if (!Number.isFinite(lineWidth) || lineWidth <= 0) return;
        const dx = x2 - x1;
        const dy = y2 - y1;
        const len = Math.hypot(dx, dy);
        if (!Number.isFinite(len) || len === 0) return;
        const nx = -dy / len;
        const ny = dx / len;
        const offset = lineWidth;

        addGroupSegment({
          rect,
          strokeStyle,
          lineWidth,
          lineDash: [],
          lineCap,
          segment: { x1: x1 + nx * offset, y1: y1 + ny * offset, x2: x2 + nx * offset, y2: y2 + ny * offset }
        });
        addGroupSegment({
          rect,
          strokeStyle,
          lineWidth,
          lineDash: [],
          lineCap,
          segment: { x1: x1 - nx * offset, y1: y1 - ny * offset, x2: x2 - nx * offset, y2: y2 - ny * offset }
        });
        return;
      }

      addGroupSegment({
        rect,
        strokeStyle,
        lineWidth: totalWidth,
        lineDash: dashForStyle(border.style, totalWidth),
        lineCap,
        segment: { x1, y1, x2, y2 }
      });
    };

    for (const entry of diagonalEntries) {
      if (entry.up) addDiagonal(entry.rect, entry.up, "up");
      if (entry.down) addDiagonal(entry.rect, entry.down, "down");
    }

    for (const group of groups.values()) {
      gridCtx.save();
      gridCtx.strokeStyle = group.strokeStyle;
      gridCtx.lineWidth = group.lineWidth;
      gridCtx.lineCap = group.lineCap;
      (gridCtx as any).setLineDash?.(group.lineDash);

      // Clip to the union of cell rects for this group to keep strokes inside cells/merged cells.
      gridCtx.beginPath();
      for (const rect of group.clipRects) {
        gridCtx.rect(rect.x, rect.y, rect.width, rect.height);
      }
      gridCtx.clip();

      gridCtx.beginPath();
      for (const segment of group.segments) {
        gridCtx.moveTo(segment.x1, segment.y1);
        gridCtx.lineTo(segment.x2, segment.y2);
      }
      gridCtx.stroke();
      gridCtx.restore();
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
      const rowsRange = this.scroll.rows.visibleRange(sheetY, intersection.height, {
        min: quadrant.minRow,
        maxExclusive: quadrant.maxRowExclusive
      });
      const colsRange = this.scroll.cols.visibleRange(sheetX, intersection.width, {
        min: quadrant.minCol,
        maxExclusive: quadrant.maxColExclusive
      });
      const startRow = rowsRange.start;
      const endRow = rowsRange.end;
      const startCol = colsRange.start;
      const endCol = colsRange.end;

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
      const clipX1 = intersection.x;
      const clipY1 = intersection.y;
      const clipX2 = intersection.x + intersection.width;
      const clipY2 = intersection.y + intersection.height;
      for (const rect of rects) {
        const x1 = Math.max(rect.x, clipX1);
        const y1 = Math.max(rect.y, clipY1);
        const x2 = Math.min(rect.x + rect.width, clipX2);
        const y2 = Math.min(rect.y + rect.height, clipY2);
        const width = x2 - x1;
        const height = y2 - y1;
        if (width <= 0 || height <= 0) continue;
        ctx.rect(x1, y1, width, height);
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
    const viewport = this.scroll.getViewportState();
    const headerRows = this.headerRowsOverride ?? (viewport.frozenRows > 0 ? 1 : 0);
    const headerCols = this.headerColsOverride ?? (viewport.frozenCols > 0 ? 1 : 0);

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
          const isHeader = row < headerRows || col < headerCols;
          const fontFamily = style?.fontFamily ?? (isHeader ? this.defaultHeaderFontFamily : this.defaultCellFontFamily);
          const fontWeight = style?.fontWeight ?? "400";
          const fontStyle = style?.fontStyle ?? "normal";
          const fontSpec = { family: fontFamily, sizePx: fontSize, weight: fontWeight, style: fontStyle } satisfies FontSpec;
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
            const textIndentX = resolveTextIndentPx(style?.textIndentPx, zoom);
            const indentX = resolvedAlign === "left" || resolvedAlign === "right" ? textIndentX : 0;
            const maxWidth = Math.max(0, availableWidth - indentX);
            const originX = x + paddingX + (resolvedAlign === "left" ? indentX : 0);

            const measurement = layoutEngine?.measure(text, fontSpec);
            const textWidth = measurement?.width ?? ctx.measureText(text).width;
            const ascent = measurement?.ascent ?? fontSize * 0.8;
            const descent = measurement?.descent ?? fontSize * 0.2;

            let textX = originX;
            if (resolvedAlign === "center") {
              textX = x + paddingX + (availableWidth - textWidth) / 2;
            } else if (resolvedAlign === "right") {
              textX = x + paddingX + (maxWidth - textWidth);
            }

            let baselineY = y + paddingY + ascent;
            if (verticalAlign === "middle") {
              baselineY = y + rowHeight / 2 + (ascent - descent) / 2;
            } else if (verticalAlign === "bottom") {
              baselineY = y + rowHeight - paddingY - descent;
            }

            const shouldClip = textWidth > maxWidth;
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
            const baseDirection = direction === "auto" ? detectBaseDirection(text) : direction;
            const resolvedAlign =
              layoutAlign === "left" || layoutAlign === "right" || layoutAlign === "center"
                ? layoutAlign
                : resolveAlign(layoutAlign, baseDirection);
            const textIndentX = rotationDeg === 0 ? resolveTextIndentPx(style?.textIndentPx, zoom) : 0;
            const indentX = resolvedAlign === "left" || resolvedAlign === "right" ? textIndentX : 0;
            const maxWidth = Math.max(0, availableWidth - indentX);
            const originX = x + paddingX + (resolvedAlign === "left" ? indentX : 0);
            if (maxWidth <= 0) {
              // Indent fully consumes the available width.
              // Nothing to render, but keep iterating the grid.
            } else {
              const layout = layoutEngine.layout({
                text,
                font: fontSpec,
                maxWidth,
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

              const shouldClip = layout.width > maxWidth || layout.height > availableHeight || rotationDeg !== 0;

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

    const clipX1 = intersection.x;
    const clipY1 = intersection.y;
    const clipX2 = intersection.x + intersection.width;
    const clipY2 = intersection.y + intersection.height;
    const rangeRectsScratch = this.rangeToViewportRectsScratchRects;

    const transientRange = this.rangeSelection;
    if (transientRange) {
      const rectCount = this.rangeToViewportRectsScratch(transientRange, viewport);

      ctx.fillStyle = this.theme.selectionFill;
      for (let i = 0; i < rectCount; i++) {
        const rect = rangeRectsScratch[i]!;
        const x1 = Math.max(rect.x, clipX1);
        const y1 = Math.max(rect.y, clipY1);
        const x2 = Math.min(rect.x + rect.width, clipX2);
        const y2 = Math.min(rect.y + rect.height, clipY2);
        const width = x2 - x1;
        const height = y2 - y1;
        if (width <= 0 || height <= 0) continue;
        ctx.fillRect(x1, y1, width, height);
      }

      ctx.strokeStyle = this.theme.selectionBorder;
      ctx.lineWidth = 2;
      for (let i = 0; i < rectCount; i++) {
        const rect = rangeRectsScratch[i]!;
        if (!rectsOverlap(rect, intersection)) continue;
        if (rect.width <= 2 || rect.height <= 2) continue;
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
      }
    }

    const fillPreview = this.fillPreviewRange;
    if (fillPreview) {
      const rectCount = this.rangeToViewportRectsScratch(fillPreview, viewport);

      if (rectCount > 0) {
        ctx.save();

        ctx.fillStyle = this.theme.selectionFill;
        ctx.globalAlpha = 0.6;
        for (let i = 0; i < rectCount; i++) {
          const rect = rangeRectsScratch[i]!;
          const x1 = Math.max(rect.x, clipX1);
          const y1 = Math.max(rect.y, clipY1);
          const x2 = Math.min(rect.x + rect.width, clipX2);
          const y2 = Math.min(rect.y + rect.height, clipY2);
          const width = x2 - x1;
          const height = y2 - y1;
          if (width <= 0 || height <= 0) continue;
          ctx.fillRect(x1, y1, width, height);
        }

        ctx.strokeStyle = this.theme.selectionBorder;
        ctx.globalAlpha = 1;
        ctx.lineWidth = 2;
        ctx.setLineDash([5, 4]);
        for (let i = 0; i < rectCount; i++) {
          const rect = rangeRectsScratch[i]!;
          if (!rectsOverlap(rect, intersection)) continue;
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
        const rectCount = this.rangeToViewportRectsScratch(range, viewport);
        if (rectCount === 0) return;

        ctx.save();

        ctx.fillStyle = this.theme.selectionFill;
        ctx.globalAlpha = options.fillAlpha;
        for (let i = 0; i < rectCount; i++) {
          const rect = rangeRectsScratch[i]!;
          const x1 = Math.max(rect.x, clipX1);
          const y1 = Math.max(rect.y, clipY1);
          const x2 = Math.min(rect.x + rect.width, clipX2);
          const y2 = Math.min(rect.y + rect.height, clipY2);
          const width = x2 - x1;
          const height = y2 - y1;
          if (width <= 0 || height <= 0) continue;
          ctx.fillRect(x1, y1, width, height);
        }

        ctx.strokeStyle = this.theme.selectionBorder;
        ctx.globalAlpha = options.strokeAlpha;
        ctx.lineWidth = options.strokeWidth;
        const inset = options.strokeWidth / 2;
        for (let i = 0; i < rectCount; i++) {
          const rect = rangeRectsScratch[i]!;
          if (!rectsOverlap(rect, intersection)) continue;
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
        const handleRect = this.getFillHandleRect();
        if (handleRect) {
          const x1 = Math.max(handleRect.x, intersection.x);
          const y1 = Math.max(handleRect.y, intersection.y);
          const x2 = Math.min(handleRect.x + handleRect.width, intersection.x + intersection.width);
          const y2 = Math.min(handleRect.y + handleRect.height, intersection.y + intersection.height);
          const width = x2 - x1;
          const height = y2 - y1;
          if (width > 0 && height > 0) {
            ctx.fillStyle = this.theme.selectionHandle;
            ctx.fillRect(x1, y1, width, height);
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

        const activeRectCount = this.rangeToViewportRectsScratch(activeCellRange, viewport);
        if (activeRectCount > 0) {
          ctx.strokeStyle = this.theme.selectionBorder;
          ctx.lineWidth = 2;
          for (let i = 0; i < activeRectCount; i++) {
            const rect = rangeRectsScratch[i]!;
            if (!rectsOverlap(rect, intersection)) continue;
            if (rect.width <= 2 || rect.height <= 2) continue;
            ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
          }
        }
      }
    }

    // Formula reference highlights: render on top of all selection overlays so
    // colored outlines remain visible while editing formulas.
    if (this.referenceHighlights.length > 0) {
      const strokeRects = (rects: Rect[], rectCount: number, lineWidth: number) => {
        const inset = lineWidth / 2;
        for (let i = 0; i < rectCount; i++) {
          const rect = rects[i]!;
          if (!rectsOverlap(rect, intersection)) continue;
          if (rect.width <= lineWidth || rect.height <= lineWidth) continue;
          ctx.strokeRect(rect.x + inset, rect.y + inset, rect.width - lineWidth, rect.height - lineWidth);
        }
      };

      ctx.save();

      ctx.lineWidth = 2;
      ctx.setLineDash([4, 3]);

      for (const highlight of this.referenceHighlights) {
        if (highlight.active) continue;
        const rectCount = this.rangeToViewportRectsScratch(highlight.range, viewport);
        if (rectCount === 0) continue;
        ctx.strokeStyle = highlight.color;
        strokeRects(rangeRectsScratch, rectCount, ctx.lineWidth);
      }

      const hasActive = this.referenceHighlights.some((h) => h.active);
      if (hasActive) {
        ctx.lineWidth = 3;
        ctx.setLineDash([]);
        for (const highlight of this.referenceHighlights) {
          if (!highlight.active) continue;
          const rectCount = this.rangeToViewportRectsScratch(highlight.range, viewport);
          if (rectCount === 0) continue;
          ctx.strokeStyle = highlight.color;
          strokeRects(rangeRectsScratch, rectCount, ctx.lineWidth);
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
    const rangeRectsScratch = this.rangeToViewportRectsScratchRects;
    const rangeScratch = this.remotePresenceRangeScratch;

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

          rangeScratch.startRow = clampedStartRow;
          rangeScratch.endRow = Math.min(rowCount, clampedEndRow + 1);
          rangeScratch.startCol = clampedStartCol;
          rangeScratch.endCol = Math.min(colCount, clampedEndCol + 1);
          const rectCount = this.rangeToViewportRectsScratch(rangeScratch, viewport);
          if (rectCount === 0) continue;

          ctx.globalAlpha = selectionFillAlpha;
          for (let i = 0; i < rectCount; i++) {
            const rect = rangeRectsScratch[i]!;
            ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
          }

          ctx.globalAlpha = selectionStrokeAlpha;
          for (let i = 0; i < rectCount; i++) {
            const rect = rangeRectsScratch[i]!;
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
        const cursorCellRect = this.remotePresenceCursorCellRectScratch;
        if (!this.cellRectInViewportInto(presence.cursor.row, presence.cursor.col, viewport, cursorCellRect)) continue;

        const cursorRange = mergedIndex.rangeAt(presence.cursor);
        let cursorRectCount = 0;
        if (cursorRange) {
          cursorRectCount = this.rangeToViewportRectsScratch(cursorRange, viewport);
        } else {
          rangeScratch.startRow = presence.cursor.row;
          rangeScratch.endRow = presence.cursor.row + 1;
          rangeScratch.startCol = presence.cursor.col;
          rangeScratch.endCol = presence.cursor.col + 1;
          cursorRectCount = this.rangeToViewportRectsScratch(rangeScratch, viewport);
        }
        if (cursorRectCount === 0) continue;

        ctx.globalAlpha = 1;
        ctx.strokeStyle = color;
        ctx.lineWidth = cursorStrokeWidth;
        for (let i = 0; i < cursorRectCount; i++) {
          const rect = rangeRectsScratch[i]!;
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

  private cellRectInViewportInto(
    row: number,
    col: number,
    viewport: GridViewportState,
    out: Rect,
    options?: { clampToViewport?: boolean }
  ): boolean {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return false;

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const colX = colAxis.positionOf(col);
    const rowY = rowAxis.positionOf(row);
    const width = colAxis.getSize(col);
    const height = rowAxis.getSize(row);

    const scrollCols = col >= viewport.frozenCols;
    const scrollRows = row >= viewport.frozenRows;
    const x = scrollCols ? colX - viewport.scrollX : colX;
    const y = scrollRows ? rowY - viewport.scrollY : rowY;

    if (options?.clampToViewport === false) {
      out.x = x;
      out.y = y;
      out.width = width;
      out.height = height;
      return true;
    }

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const quadrantX = scrollCols ? frozenWidthClamped : 0;
    const quadrantY = scrollRows ? frozenHeightClamped : 0;
    const quadrantWidth = scrollCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped;
    const quadrantHeight = scrollRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped;

    const x1 = Math.max(x, quadrantX);
    const y1 = Math.max(y, quadrantY);
    const x2 = Math.min(x + width, quadrantX + quadrantWidth);
    const y2 = Math.min(y + height, quadrantY + quadrantHeight);
    const clippedWidth = x2 - x1;
    const clippedHeight = y2 - y1;
    if (clippedWidth <= 0 || clippedHeight <= 0) return false;

    out.x = x1;
    out.y = y1;
    out.width = clippedWidth;
    out.height = clippedHeight;
    return true;
  }

  private cellRectInViewportIntoWithScroll(
    row: number,
    col: number,
    viewport: GridViewportState,
    scrollX: number,
    scrollY: number,
    out: Rect,
    options?: { clampToViewport?: boolean }
  ): boolean {
    const rowCount = this.getRowCount();
    const colCount = this.getColCount();
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return false;

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const colX = colAxis.positionOf(col);
    const rowY = rowAxis.positionOf(row);
    const width = colAxis.getSize(col);
    const height = rowAxis.getSize(row);

    const scrollCols = col >= viewport.frozenCols;
    const scrollRows = row >= viewport.frozenRows;
    const x = scrollCols ? colX - scrollX : colX;
    const y = scrollRows ? rowY - scrollY : rowY;

    if (options?.clampToViewport === false) {
      out.x = x;
      out.y = y;
      out.width = width;
      out.height = height;
      return true;
    }

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);
    const quadrantX = scrollCols ? frozenWidthClamped : 0;
    const quadrantY = scrollRows ? frozenHeightClamped : 0;
    const quadrantWidth = scrollCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped;
    const quadrantHeight = scrollRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped;

    const x1 = Math.max(x, quadrantX);
    const y1 = Math.max(y, quadrantY);
    const x2 = Math.min(x + width, quadrantX + quadrantWidth);
    const y2 = Math.min(y + height, quadrantY + quadrantHeight);
    const clippedWidth = x2 - x1;
    const clippedHeight = y2 - y1;
    if (clippedWidth <= 0 || clippedHeight <= 0) return false;

    out.x = x1;
    out.y = y1;
    out.width = clippedWidth;
    out.height = clippedHeight;
    return true;
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
    const rect = this.selectionDirtyRectScratch;
    rect.x = 0;
    rect.y = 0;
    rect.width = viewport.width;
    rect.height = viewport.height;
    this.dirty.selection.markDirty(rect);
    this.requestRender();
  }

  private fillPreviewOverlayRects(range: CellRange | null, viewport: GridViewportState): Rect[] {
    if (!range) return [];
    return this.rangeToViewportRects(range, viewport);
  }

  private fillHandleRectInViewport(range: CellRange, viewport: GridViewportState, out: Rect): boolean {
    const handleSize = DEFAULT_FILL_HANDLE_SIZE_PX * this.zoom;
    const handleRow = range.endRow - 1;
    const handleCol = range.endCol - 1;

    const rowAxis = this.scroll.rows;
    const colAxis = this.scroll.cols;

    const colX = colAxis.positionOf(handleCol);
    const rowY = rowAxis.positionOf(handleRow);
    const cellWidth = colAxis.getSize(handleCol);
    const cellHeight = rowAxis.getSize(handleRow);
    if (cellWidth < handleSize || cellHeight < handleSize) return false;

    const scrollCols = handleCol >= viewport.frozenCols;
    const scrollRows = handleRow >= viewport.frozenRows;
    const cellX = scrollCols ? colX - viewport.scrollX : colX;
    const cellY = scrollRows ? rowY - viewport.scrollY : rowY;

    const handleX = cellX + cellWidth - handleSize / 2;
    const handleY = cellY + cellHeight - handleSize / 2;

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

    const quadrantX = scrollCols ? frozenWidthClamped : 0;
    const quadrantY = scrollRows ? frozenHeightClamped : 0;
    const quadrantWidth = scrollCols ? Math.max(0, viewport.width - frozenWidthClamped) : frozenWidthClamped;
    const quadrantHeight = scrollRows ? Math.max(0, viewport.height - frozenHeightClamped) : frozenHeightClamped;

    const x1 = Math.max(handleX, quadrantX);
    const y1 = Math.max(handleY, quadrantY);
    const x2 = Math.min(handleX + handleSize, quadrantX + quadrantWidth);
    const y2 = Math.min(handleY + handleSize, quadrantY + quadrantHeight);
    const width = x2 - x1;
    const height = y2 - y1;
    if (width <= 0 || height <= 0) return false;

    out.x = x1;
    out.y = y1;
    out.width = width;
    out.height = height;
    return true;
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
        const getMergedRangeAt = provider.getMergedRangeAt.bind(provider);
        const startRow = expanded.startRow;
        const endRow = expanded.endRow;
        const startCol = expanded.startCol;
        const endCol = expanded.endCol;

        const lastRow = endRow - 1;
        const lastCol = endCol - 1;

        // Scan the range perimeter (not the interior), but avoid O(range height/width) work
        // for extremely tall/wide merged ranges by skipping over spans when we detect a merge.
        //
        // We only need to consider merges that *cross* the current range boundary (and therefore
        // touch its perimeter). Merges fully contained in `expanded` do not affect the resulting
        // expanded bounds.
        const normalizeMerge = (candidate: CellRange | null): CellRange | null => {
          const normalized = candidate ? this.normalizeSelectionRange(candidate) : null;
          if (!normalized) return null;
          if (normalized.endRow - normalized.startRow <= 1 && normalized.endCol - normalized.startCol <= 1) return null;
          return normalized;
        };

        const scanVerticalEdge = (col: number) => {
          for (let row = startRow; row < endRow; ) {
            const normalized = normalizeMerge(getMergedRangeAt(row, col));
            if (normalized) {
              addMerge(normalized);
              // Jump to the first row after this merged region.
              row = Math.max(row + 1, normalized.endRow);
              continue;
            }
            row++;
          }
        };

        const scanHorizontalEdge = (row: number) => {
          for (let col = startCol; col < endCol; ) {
            const normalized = normalizeMerge(getMergedRangeAt(row, col));
            if (normalized) {
              addMerge(normalized);
              // Jump to the first column after this merged region.
              col = Math.max(col + 1, normalized.endCol);
              continue;
            }
            col++;
          }
        };

        // Left and right edges.
        scanVerticalEdge(startCol);
        if (lastCol !== startCol) scanVerticalEdge(lastCol);

        // Top and bottom edges.
        scanHorizontalEdge(startRow);
        if (lastRow !== startRow) scanHorizontalEdge(lastRow);
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
      this.mergedIndexViewport = null;
      this.mergedIndexDirty = false;
      return this.mergedIndex;
    }

    if (!this.mergedIndexDirty && this.mergedIndexViewport === viewport) {
      return this.mergedIndex;
    }

    const queryRanges = this.getMergedQueryRanges(viewport);
    const key = queryRanges.map((range) => `${range.startRow},${range.endRow},${range.startCol},${range.endCol}`).join("|");

    if (!this.mergedIndexDirty && this.mergedIndexKey === key) {
      this.mergedIndexViewport = viewport;
      return this.mergedIndex;
    }

    // Only materialize row->span entries for rows that are actually visible in the current viewport.
    //
    // We also include the row immediately above the main visible range, because gridline/border
    // rendering checks `isInteriorHorizontalGridline(index, startRow - 1, col)` for the top-most
    // visible row. Without this, merged regions that start above the viewport would incorrectly
    // draw a horizontal gridline at the top edge of the viewport.
    //
    // This still keeps indexing costs proportional to viewport size (O(visible rows)).
    const indexedRowRanges: IndexedRowRange[] = [];
    if (viewport.width > 0 && viewport.height > 0) {
      const frozenHeight = Math.min(viewport.height, viewport.frozenHeight);
      const frozenRowsRange =
        viewport.frozenRows === 0 || frozenHeight === 0
          ? { start: 0, end: 0 }
          : this.scroll.rows.visibleRange(0, frozenHeight, { min: 0, maxExclusive: viewport.frozenRows });
      if (frozenRowsRange.end > frozenRowsRange.start) {
        indexedRowRanges.push({ startRow: frozenRowsRange.start, endRow: frozenRowsRange.end });
      }

      const mainRows = viewport.main.rows;
      if (mainRows.end > mainRows.start) {
        indexedRowRanges.push({ startRow: Math.max(0, mainRows.start - 1), endRow: mainRows.end });
      }
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

    this.mergedIndex = new MergedCellIndex([...merges.values()], indexedRowRanges);
    this.mergedIndexKey = key;
    this.mergedIndexViewport = viewport;
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
    this.notifyViewportChange("axisSize");
  }

  private notifyViewportChange(reason: GridViewportChangeReason): void {
    if (this.viewportListeners.size === 0) return;
    let viewport: GridViewportState | null = null;
    const getViewport = () => {
      if (!viewport) viewport = this.scroll.getViewportState();
      return viewport;
    };

    for (const entry of this.viewportListeners) {
      const debounceMs = entry.options.debounceMs;
      if (typeof debounceMs === "number" && Number.isFinite(debounceMs) && debounceMs >= 0) {
        entry.pendingReason = reason;
        if (entry.timeoutId !== null) clearTimeout(entry.timeoutId);
        entry.timeoutId = setTimeout(() => {
          entry.timeoutId = null;
          const pendingReason = entry.pendingReason;
          entry.pendingReason = null;
          if (!pendingReason) return;
          entry.listener({ viewport: this.scroll.getViewportState(), reason: pendingReason });
        }, debounceMs);
        continue;
      }

      if (entry.options.animationFrame) {
        entry.pendingReason = reason;
        if (entry.rafId !== null) continue;
        // Some unit tests stub `requestAnimationFrame` to invoke the callback synchronously.
        // If that happens and we assign `entry.rafId` *after* `requestAnimationFrame` returns,
        // the callback's `entry.rafId = null` can be overwritten with the returned id, leaving
        // the subscription stuck in a "pending RAF" state (future viewport changes won't
        // schedule).
        //
        // Track whether the callback executed synchronously and only persist the returned id
        // when it hasn't yet run.
        let executedSynchronously = false;
        const rafId = requestAnimationFrame(() => {
          executedSynchronously = true;
          entry.rafId = null;
          const pendingReason = entry.pendingReason;
          entry.pendingReason = null;
          if (!pendingReason) return;
          entry.listener({ viewport: this.scroll.getViewportState(), reason: pendingReason });
        });
        entry.rafId = executedSynchronously ? null : rafId;
        continue;
      }

      entry.listener({ viewport: getViewport(), reason });
    }
  }
}
