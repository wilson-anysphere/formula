import type { Anchor, DrawingObject, DrawingTransform, ImageEntry, ImageStore, Rect } from "./types";
import { ImageBitmapCache } from "./imageBitmapCache";
import { graphicFramePlaceholderLabel, isGraphicFrame, parseShapeRenderSpec, type ShapeRenderSpec } from "./shapeRenderer";
import { parseDrawingMLShapeText, type ShapeTextLayout, type ShapeTextRun } from "./drawingml/shapeText";
import { getDrawingMLPicturePropsFromPicXml } from "./drawingml/pictureProps";
import {
  getResizeHandleCentersInto,
  getRotationHandleCenterInto,
  RESIZE_HANDLE_SIZE_PX,
  ROTATION_HANDLE_SIZE_PX,
  type ResizeHandleCenter,
  type RotationHandleCenter,
} from "./selectionHandles";
import { degToRad } from "./transform";
import { DrawingSpatialIndex } from "./spatialIndex";

import { EMU_PER_INCH, PX_PER_INCH } from "../shared/emu.js";

export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH;
export { EMU_PER_INCH, PX_PER_INCH };

function resolveZoom(raw: number | undefined): number {
  return typeof raw === "number" && Number.isFinite(raw) && raw > 0 ? raw : 1;
}

export function emuToPx(emu: number, zoom?: number): number {
  const z = resolveZoom(zoom);
  return (emu / EMU_PER_PX) * z;
}

export function pxToEmu(px: number, zoom?: number): number {
  const z = resolveZoom(zoom);
  const emu = (px * EMU_PER_PX) / z;
  // `Math.round` rounds half values toward +∞, which makes the conversion non-odd:
  // `Math.round(-0.5) === -0` while `Math.round(0.5) === 1`. For reversible interactions
  // (e.g. ArrowRight then ArrowLeft at non-1x zoom), we want `pxToEmu(-x) === -pxToEmu(x)`.
  // Achieve that by rounding half values away from zero.
  const rounded = Math.round(Math.abs(emu));
  if (!Number.isFinite(rounded) || rounded === 0) return 0;
  return emu < 0 ? -rounded : rounded;
}

type CssVarStyle = Pick<CSSStyleDeclaration, "getPropertyValue">;

type OverlayColorTokens = {
  placeholderChartStroke: string;
  placeholderOtherStroke: string;
  placeholderGraphicFrameStroke: string;
  placeholderLabel: string;
  selectionStroke: string;
  selectionHandleFill: string;
};

type ShapeTextCacheEntry = {
  rawXml: string;
  parsed: ShapeTextLayout | null;
  hasText: boolean;
  spec: ShapeRenderSpec | null;
  // Cached line layout to avoid re-measuring/wrapping on every render.
  lines: RenderedLine[] | null;
  linesMaxWidth: number;
  linesWrap: boolean;
  linesZoom: number;
  linesDefaultColor: string;
};

const DEFAULT_OVERLAY_COLOR_TOKENS: OverlayColorTokens = {
  placeholderChartStroke: "blue",
  placeholderOtherStroke: "cyan",
  placeholderGraphicFrameStroke: "magenta",
  placeholderLabel: "black",
  selectionStroke: "blue",
  selectionHandleFill: "white",
};

const LINE_DASH_NONE: number[] = [];
const LINE_DASH_PLACEHOLDER: number[] = [4, 2];

const SHAPE_TEXT_INNER_RECT_SCRATCH: Rect = { x: 0, y: 0, width: 0, height: 0 };
const SHAPE_TEXT_LAYOUT_OPTS_SCRATCH: { wrap: boolean; defaultColor: string; zoom: number } = {
  wrap: true,
  defaultColor: "",
  zoom: 1,
};

function getRootCssStyle(root: unknown): CssVarStyle | null {
  if (!root || typeof getComputedStyle !== "function") return null;
  try {
    return getComputedStyle(root as any);
  } catch {
    return null;
  }
}

function parseCssVarFunction(value: unknown): { name: string; fallback: string | null } | null {
  const trimmed = String(value ?? "").trim();
  if (!trimmed.startsWith("var(") || !trimmed.endsWith(")")) return null;

  const inner = trimmed.slice(4, -1).trim();
  if (!inner.startsWith("--")) return null;

  let depth = 0;
  let comma = -1;
  let inSingle = false;
  let inDouble = false;

  for (let i = 0; i < inner.length; i += 1) {
    const ch = inner[i];
    const prev = i > 0 ? inner[i - 1] : "";

    if (inSingle) {
      if (ch === "'" && prev !== "\\") inSingle = false;
      continue;
    }
    if (inDouble) {
      if (ch === '"' && prev !== "\\") inDouble = false;
      continue;
    }

    if (ch === "'") {
      inSingle = true;
      continue;
    }
    if (ch === '"') {
      inDouble = true;
      continue;
    }

    if (ch === "(") depth += 1;
    else if (ch === ")") depth = Math.max(0, depth - 1);
    else if (ch === "," && depth === 0) {
      comma = i;
      break;
    }
  }

  const name = comma === -1 ? inner : inner.slice(0, comma);
  const varName = name.trim();
  if (!varName.startsWith("--")) return null;

  const fallbackValue = comma === -1 ? null : inner.slice(comma + 1).trim() || null;
  return { name: varName, fallback: fallbackValue };
}

function resolveCssValue(
  style: CssVarStyle | null,
  value: unknown,
  fallback: string,
  maxDepth = 8,
  seen?: Set<string>,
): string {
  let current = String(value ?? "").trim();
  let lastFallback: string | null = null;
  // Perf: resolving solid colors can be on the per-frame hot path when many
  // shapes are present. Avoid allocating a `Set` unless we actually encounter
  // a `var(--token)` indirection.
  let seenSet = seen;

  for (let depth = 0; depth < maxDepth; depth += 1) {
    const parsed = parseCssVarFunction(current);
    if (!parsed) return current || fallback;

    const nextName = parsed.name;
    if (parsed.fallback != null) lastFallback = parsed.fallback;

    if (!seenSet) seenSet = new Set<string>();

    // Handle cycles and enforce a max indirection depth.
    if (seenSet.has(nextName)) {
      current = parsed.fallback ?? lastFallback ?? "";
      continue;
    }
    seenSet.add(nextName);

    let nextValue = "";
    if (style) {
      const raw = style.getPropertyValue(nextName);
      nextValue = typeof raw === "string" ? raw.trim() : "";
    }
    if (nextValue) {
      current = nextValue;
      continue;
    }

    const fb = parsed.fallback ?? lastFallback;
    if (fb != null) return resolveCssValue(style, fb, fallback, maxDepth, seenSet);
    return fallback;
  }

  if (lastFallback != null) return resolveCssValue(style, lastFallback, fallback, maxDepth, seenSet);
  return fallback;
}

function resolveCssVarFromStyle(style: CssVarStyle | null, varName: string, fallback: string): string {
  if (!style) return fallback;
  const start = style.getPropertyValue(varName);
  const trimmed = typeof start === "string" ? start.trim() : "";
  if (!trimmed) return fallback;
  return resolveCssValue(style, trimmed, fallback, 8, new Set([varName]));
}

function resolveOverlayColorTokens(style: CssVarStyle | null): OverlayColorTokens {
  const textPrimary = resolveCssVarFromStyle(style, "--text-primary", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderLabel);
  const selectionBorder = resolveCssVarFromStyle(
    style,
    "--selection-border",
    DEFAULT_OVERLAY_COLOR_TOKENS.selectionStroke,
  );
  const bgPrimary = resolveCssVarFromStyle(style, "--bg-primary", DEFAULT_OVERLAY_COLOR_TOKENS.selectionHandleFill);
  return {
    placeholderChartStroke: resolveCssVarFromStyle(style, "--chart-series-1", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderChartStroke),
    placeholderOtherStroke: resolveCssVarFromStyle(style, "--chart-series-2", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderOtherStroke),
    placeholderGraphicFrameStroke: resolveCssVarFromStyle(
      style,
      "--chart-series-3",
      DEFAULT_OVERLAY_COLOR_TOKENS.placeholderGraphicFrameStroke,
    ),
    placeholderLabel: resolveCssVarFromStyle(style, "--formula-grid-cell-text", textPrimary),
    selectionStroke: resolveCssVarFromStyle(style, "--formula-grid-selection-border", selectionBorder),
    selectionHandleFill: resolveCssVarFromStyle(style, "--formula-grid-bg", bgPrimary),
  };
}

function resolveCanvasColor(style: CssVarStyle | null, input: string, fallback: string): string {
  return resolveCssValue(style, input, fallback);
}

export interface GridGeometry {
  /** Sheet-space pixel origin for the top-left of a cell. */
  cellOriginPx(cell: { row: number; col: number }): { x: number; y: number };
  /** Pixel size of a cell. */
  cellSizePx(cell: { row: number; col: number }): { width: number; height: number };
}

export interface Viewport {
  scrollX: number;
  scrollY: number;
  width: number;
  height: number;
  dpr: number;
  /**
   * Sheet zoom factor (1 = 100%).
   *
   * This is applied when converting DrawingML EMU-derived geometry into
   * screen-space pixels.
   *
   * Defaults to 1 for legacy/non-zoomed grids.
   */
  zoom?: number;
  /**
   * Frozen pane counts in *sheet-space* (i.e. they do not include any synthetic
   * header rows/cols used by shared-grid mode).
   */
  frozenRows?: number;
  frozenCols?: number;
  /**
   * Frozen pane extents in *viewport* coordinates (pixels).
   *
   * In shared-grid mode the underlying grid viewport state typically reports
   * frozen extents including header rows/cols. If `headerOffsetX/Y` are also
   * provided, the effective frozen content size becomes:
   *
   *   frozenContentWidth  = frozenWidthPx  - headerOffsetX
   *   frozenContentHeight = frozenHeightPx - headerOffsetY
   */
  frozenWidthPx?: number;
  frozenHeightPx?: number;
  /**
   * Optional viewport-space offsets for grid headers (row/col headers).
   *
   * When provided, drawing objects are shifted by this amount so they are
   * rendered under headers rather than on top of them.
   */
  headerOffsetX?: number;
  headerOffsetY?: number;
}

export interface ChartRenderer {
  renderToCanvas(ctx: CanvasRenderingContext2D, chartId: string, rect: Rect): void;
  /**
   * Optional cache-pruning hook.
   *
   * Some chart renderers (e.g. `ChartRendererAdapter`) keep per-chart offscreen surfaces
   * around to avoid re-rendering on scroll. `DrawingOverlay` calls this when the set of
   * chart objects changes so implementations can drop surfaces for deleted or off-sheet
   * charts.
   */
  pruneSurfaces?(keep: ReadonlySet<string>): void;
  destroy?(): void;
}

const A1_CELL = { row: 0, col: 0 };

export function anchorToRectPx(anchor: Anchor, geom: GridGeometry, zoom?: number): Rect {
  switch (anchor.type) {
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        x: origin.x + emuToPx(anchor.from.offset.xEmu, zoom),
        y: origin.y + emuToPx(anchor.from.offset.yEmu, zoom),
        width: emuToPx(anchor.size.cx, zoom),
        height: emuToPx(anchor.size.cy, zoom),
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);

      const x1 = fromOrigin.x + emuToPx(anchor.from.offset.xEmu, zoom);
      const y1 = fromOrigin.y + emuToPx(anchor.from.offset.yEmu, zoom);

      // In DrawingML, `to` specifies the cell *containing* the bottom-right
      // corner (i.e. the first cell strictly outside the shape when the corner
      // lies on a grid boundary). The absolute end point is therefore the
      // origin of the `to` cell plus the offsets.
      const x2 = toOrigin.x + emuToPx(anchor.to.offset.xEmu, zoom);
      const y2 = toOrigin.y + emuToPx(anchor.to.offset.yEmu, zoom);

      return {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        width: Math.abs(x2 - x1),
        height: Math.abs(y2 - y1),
      };
    }
    case "absolute": {
      // In DrawingML, absolute anchors are worksheet-space coordinates whose
      // origin is the top-left of the cell grid (A1), not the top-left of the
      // grid UI root (which may include row/column headers).
      //
      // Use the A1 origin from the grid geometry so drawings align with
      // oneCell/twoCell anchors when the overlay canvas covers the full grid
      // surface.
      const origin = geom.cellOriginPx(A1_CELL);
      return {
        x: origin.x + emuToPx(anchor.pos.xEmu, zoom),
        y: origin.y + emuToPx(anchor.pos.yEmu, zoom),
        width: emuToPx(anchor.size.cx, zoom),
        height: emuToPx(anchor.size.cy, zoom),
      };
    }
  }
}

export function effectiveScrollForAnchor(
  anchor: Anchor,
  viewport: Pick<Viewport, "scrollX" | "scrollY" | "frozenRows" | "frozenCols">,
): { scrollX: number; scrollY: number } {
  const frozenRows = viewport.frozenRows ?? 0;
  const frozenCols = viewport.frozenCols ?? 0;

  if (anchor.type === "absolute") {
    return { scrollX: viewport.scrollX, scrollY: viewport.scrollY };
  }

  const from = anchor.from.cell;
  return {
    scrollX: from.col < frozenCols ? 0 : viewport.scrollX,
    scrollY: from.row < frozenRows ? 0 : viewport.scrollY,
  };
}

export class DrawingOverlay {
  private readonly ctx: CanvasRenderingContext2D;
  private readonly bitmapCache = new ImageBitmapCache({ negativeCacheMs: 250 });
  private readonly shapeTextCache = new Map<number, ShapeTextCacheEntry>();
  private shapeTextCachePruneSource: DrawingObject[] | null = null;
  private shapeTextCachePruneLength = 0;
  private readonly spatialIndex = new DrawingSpatialIndex();
  private requestRender: (() => void) | null = null;
  private readonly onBitmapReady = () => {
    this.scheduleHydrationRerender();
  };
  private resizeMemo: { cssWidth: number; cssHeight: number; dpr: number; pixelWidth: number; pixelHeight: number } | null = null;
  private selectedId: number | null = null;
  private preloadAbort: AbortController | null = typeof AbortController !== "undefined" ? new AbortController() : null;
  private preloadCount = 0;
  private destroyed = false;
  private cssVarStyle: CssVarStyle | null | undefined = undefined;
  private colorTokens: OverlayColorTokens | null = null;
  private chartSurfacePruneSource: DrawingObject[] | null = null;
  private chartSurfacePruneLength = 0;
  private readonly chartSurfaceKeep = new Set<string>();
  private themeObserver: MutationObserver | null = null;
  private lastRenderArgs: { objects: DrawingObject[]; viewport: Viewport; options?: { drawObjects?: boolean } } | null = null;
  private readonly lastRenderViewportScratch: Viewport = { scrollX: 0, scrollY: 0, width: 0, height: 0, dpr: 1 };
  private readonly pendingImageHydrations = new Map<string, Promise<ImageEntry | undefined>>();
  private readonly imageHydrationNegativeCache = new Map<string, number>();
  private imageHydrationEpoch = 0;
  private hydrationRerenderScheduled = false;
  // Scratch state to avoid per-frame allocations during render().
  private readonly viewportRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly queryRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly visibleObjectsScratch: DrawingObject[] = [];
  private readonly queryCandidatesScratch: DrawingObject[] = [];
  private readonly screenRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly aabbScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly localRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly selectedScreenRectScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly selectedAabbScratch: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly resizeHandleCentersScratch: ResizeHandleCenter[] = [];
  private readonly rotationHandleCenterScratch: RotationHandleCenter = { x: 0, y: 0 };
  private readonly dashPatternScratch: number[] = [];
  private readonly paneLayoutScratch: PaneLayout = {
    frozenRows: 0,
    frozenCols: 0,
    headerOffsetX: 0,
    headerOffsetY: 0,
    quadrants: {
      topLeft: { x: 0, y: 0, width: 0, height: 0 },
      topRight: { x: 0, y: 0, width: 0, height: 0 },
      bottomLeft: { x: 0, y: 0, width: 0, height: 0 },
      bottomRight: { x: 0, y: 0, width: 0, height: 0 },
    },
  };

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly images: ImageStore,
    private readonly geom: GridGeometry,
    private readonly chartRenderer?: ChartRenderer,
    requestRender?: (() => void) | null,
    private readonly cssVarRoot: HTMLElement | null = null,
  ) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("overlay canvas 2d context not available");
    this.ctx = ctx;
    this.requestRender = requestRender ?? null;
    this.installThemeObserver();
  }

  private installThemeObserver(): void {
    if (typeof MutationObserver !== "function") return;
    const doc = (this.canvas as any)?.ownerDocument ?? (globalThis as any)?.document ?? null;
    const docRoot = doc?.documentElement ?? null;

    const roots: unknown[] = [];
    if (docRoot) roots.push(docRoot);
    if (this.cssVarRoot && this.cssVarRoot !== docRoot) roots.push(this.cssVarRoot);
    if (roots.length === 0) return;

    try {
      this.themeObserver?.disconnect();
      this.themeObserver = new MutationObserver(() => {
        // Theme changes affect the colors used for placeholders + selection handles.
        // Invalidate cached tokens and request a redraw using the most recent render
        // arguments so overlays update immediately (without requiring a scroll).
        this.refreshThemeTokens();
        this.scheduleHydrationRerender();
      });

      // Custom properties can be toggled via classes/inline styles on the grid root
      // (in addition to `<html data-theme=...>`). Observe both so overlays remain in sync.
      const attributeFilter = ["style", "class", "data-theme", "data-reduced-motion"];
      for (const root of roots) {
        this.themeObserver.observe(root as any, { attributes: true, attributeFilter });
      }
    } catch {
      this.themeObserver = null;
    }
  }

  resize(viewport: Viewport): void {
    if (this.destroyed) return;
    const cssWidth = Number.isFinite(viewport.width) ? viewport.width : 0;
    const cssHeight = Number.isFinite(viewport.height) ? viewport.height : 0;
    const dpr = Number.isFinite(viewport.dpr) && viewport.dpr > 0 ? viewport.dpr : 1;

    const pixelWidth = Math.max(0, Math.floor(cssWidth * dpr));
    const pixelHeight = Math.max(0, Math.floor(cssHeight * dpr));

    const memo = this.resizeMemo;
    if (
      memo &&
      memo.cssWidth === cssWidth &&
      memo.cssHeight === cssHeight &&
      memo.dpr === dpr &&
      memo.pixelWidth === pixelWidth &&
      memo.pixelHeight === pixelHeight
    ) {
      return;
    }

    this.resizeMemo = { cssWidth, cssHeight, dpr, pixelWidth, pixelHeight };
    this.canvas.width = pixelWidth;
    this.canvas.height = pixelHeight;
    this.canvas.style.width = `${cssWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  refreshThemeTokens(): void {
    this.cssVarStyle = undefined;
    this.colorTokens = null;
  }

  /**
   * Invalidates the cached spatial index (used for viewport culling).
   *
   * Call this when the underlying grid geometry changes (e.g. row/col resize) even
   * if the `DrawingObject[]` reference is unchanged.
   */
  invalidateSpatialIndex(): void {
    this.spatialIndex.invalidate();
  }

  private getCssVarStyle(): CssVarStyle | null {
    if (this.cssVarStyle !== undefined) return this.cssVarStyle;
    const root = this.cssVarRoot ?? (this.canvas as any)?.ownerDocument?.documentElement ?? (globalThis as any)?.document?.documentElement ?? null;
    this.cssVarStyle = getRootCssStyle(root);
    return this.cssVarStyle;
  }

  setRequestRender(requestRender: (() => void) | null): void {
    this.requestRender = requestRender;
  }

  private scheduleHydrationRerender(): void {
    if (this.hydrationRerenderScheduled) return;
    this.hydrationRerenderScheduled = true;

    const schedule =
      typeof queueMicrotask === "function"
        ? queueMicrotask
        : (cb: () => void) => {
            void Promise.resolve().then(cb);
          };

    schedule(() => {
      this.hydrationRerenderScheduled = false;
      if (this.destroyed) return;
      const cb = this.requestRender;
      if (cb) {
        try {
          const result = cb() as unknown;
          // `requestRender` is expected to be synchronous, but unit tests sometimes stub it with an async
          // mock (returning a Promise). Swallow async rejections to avoid unhandled promise rejection
          // noise when image hydration triggers follow-up renders.
          if (typeof (result as { then?: unknown } | null)?.then === "function") {
            void Promise.resolve(result).catch(() => {
              // Best-effort: rendering hooks should never throw from cache callbacks.
            });
          }
        } catch {
          // Best-effort: rendering hooks should never throw from cache callbacks.
        }
        return;
      }

      const last = this.lastRenderArgs;
      if (!last) return;
      try {
        const result = this.render(last.objects, last.viewport, last.options) as unknown;
        // `render()` is synchronous, but tests can stub it as an async mock. Swallow async rejections
        // so we don't surface unhandled promise rejections from microtask-based hydration rerenders.
        if (typeof (result as { then?: unknown } | null)?.then === "function") {
          void Promise.resolve(result).catch(() => {
            // Best-effort: avoid throwing from microtask callbacks.
          });
        }
      } catch {
        // Best-effort: avoid throwing from microtask callbacks.
      }
    });
  }

  private hydrateImage(imageId: string): void {
    const id = String(imageId ?? "");
    if (!id) return;
    if (typeof this.images.getAsync !== "function") return;

    // Avoid hammering async stores (e.g. IndexedDB) if an image cannot be found.
    const cachedUntil = this.imageHydrationNegativeCache.get(id);
    if (cachedUntil != null) {
      if (cachedUntil > Date.now()) return;
      this.imageHydrationNegativeCache.delete(id);
    }

    if (this.pendingImageHydrations.has(id)) return;

    const epoch = this.imageHydrationEpoch;
    const promise = Promise.resolve()
      .then(() => {
        // Avoid starting async store work if the overlay has been torn down or the
        // caller cleared image caches while we were waiting for this microtask turn.
        if (this.destroyed) return undefined;
        if (epoch !== this.imageHydrationEpoch) return undefined;
        return this.images.getAsync!(id);
      })
      .catch(() => undefined);

    this.pendingImageHydrations.set(id, promise);

    void promise.then((entry) => {
      this.pendingImageHydrations.delete(id);
      // The overlay can be torn down (sheet close, split pane destruction, tests/hot reload)
      // while an async hydration is still in-flight. Avoid mutating shared image stores after
      // teardown so the decoded bytes can be released promptly.
      if (this.destroyed) return;
      // If callers cleared image caches (e.g. applyState swapping workbook snapshots) while a
      // hydration request was already in-flight, ignore the stale result so we don't repopulate
      // caches with bytes from the previous epoch.
      if (epoch !== this.imageHydrationEpoch) return;
      if (!entry) {
        this.imageHydrationNegativeCache.set(id, Date.now() + 250);
        return;
      }

      // Ensure subsequent sync `get()` calls can resolve without awaiting `getAsync`.
      try {
        if (!this.images.get(id)) {
          this.images.set(entry);
        }
      } catch {
        // Best-effort: ignore caching failures.
      }

      // Kick off bitmap decode so we can re-render once ready.
      // Note: the bitmap cache already negative-caches failures to avoid tight retry loops.
      if (entry.bytes.byteLength > 0) {
        this.bitmapCache.getOrRequest(entry, this.onBitmapReady);
      }

      this.scheduleHydrationRerender();
    });
  }

  render(objects: DrawingObject[], viewport: Viewport, options?: { drawObjects?: boolean }): void {
    if (this.destroyed) return;
    let completed = false;
    const drawObjects = options?.drawObjects !== false;
    // Keep the latest render args around so async image hydration can trigger a follow-up render
    // once bytes are available (without relying on callers to poll/refresh).
    const lastViewport = this.lastRenderViewportScratch;
    lastViewport.scrollX = viewport.scrollX;
    lastViewport.scrollY = viewport.scrollY;
    lastViewport.width = viewport.width;
    lastViewport.height = viewport.height;
    lastViewport.dpr = viewport.dpr;
    lastViewport.zoom = viewport.zoom;
    lastViewport.frozenRows = viewport.frozenRows;
    lastViewport.frozenCols = viewport.frozenCols;
    lastViewport.frozenWidthPx = viewport.frozenWidthPx;
    lastViewport.frozenHeightPx = viewport.frozenHeightPx;
    lastViewport.headerOffsetX = viewport.headerOffsetX;
    lastViewport.headerOffsetY = viewport.headerOffsetY;

    const lastArgs = this.lastRenderArgs;
    if (lastArgs) {
      lastArgs.objects = objects;
      lastArgs.options = options;
    } else {
      // Allocate once and reuse across renders to avoid per-frame `{ ...viewport }` cloning.
      this.lastRenderArgs = { objects, viewport: lastViewport, options };
    }

    const chartRenderer = this.chartRenderer;
    if (chartRenderer && typeof chartRenderer.pruneSurfaces === "function") {
      const sourceChanged =
        this.chartSurfacePruneSource !== objects || this.chartSurfacePruneLength !== objects.length;
      if (sourceChanged) {
        this.chartSurfacePruneSource = objects;
        this.chartSurfacePruneLength = objects.length;
        this.chartSurfaceKeep.clear();
        for (const obj of objects) {
          if (obj.kind.type !== "chart") continue;
          const chartId = (obj.kind as any).chartId;
          if (typeof chartId === "string" && chartId.length > 0) this.chartSurfaceKeep.add(chartId);
        }
        chartRenderer.pruneSurfaces(this.chartSurfaceKeep);
      }
    }

    try {
    const ctx = this.ctx;
    ctx.clearRect(0, 0, viewport.width, viewport.height);

    const cssVarStyle = this.getCssVarStyle();
    const colors = this.colorTokens ?? (this.colorTokens = resolveOverlayColorTokens(cssVarStyle));
    const zoom = resolveZoom(viewport.zoom);

    const paneLayout = resolvePaneLayout(viewport, this.geom, this.paneLayoutScratch);
    const viewportRect = this.viewportRectScratch;
    viewportRect.x = 0;
    viewportRect.y = 0;
    viewportRect.width = viewport.width;
    viewportRect.height = viewport.height;

    // Spatial index: compute a small candidate list for the current viewport rather
    // than scanning every drawing on each render.
    this.spatialIndex.rebuild(objects, this.geom, zoom);
    const ordered = this.visibleObjectsScratch;
    ordered.length = 0;
    const candidatesScratch = this.queryCandidatesScratch;
    const queryRect = this.queryRectScratch;
    const frozenContentWidth = paneLayout.quadrants.topLeft.width;
    const frozenContentHeight = paneLayout.quadrants.topLeft.height;
    const frozenRows = paneLayout.frozenRows;
    const frozenCols = paneLayout.frozenCols;

    // Collect visible candidates per quadrant (avoid allocating per-query rects).
    let w = paneLayout.quadrants.topLeft.width;
    let h = paneLayout.quadrants.topLeft.height;
    if (w > 0 && h > 0) {
      queryRect.x = 0;
      queryRect.y = 0;
      queryRect.width = w;
      queryRect.height = h;
      const candidates = this.spatialIndex.query(queryRect, candidatesScratch);
      for (const obj of candidates) {
        const anchor = obj.anchor;
        if (anchor.type === "absolute") continue;
        if (anchor.from.cell.row < frozenRows && anchor.from.cell.col < frozenCols) ordered.push(obj);
      }
    }

    w = paneLayout.quadrants.topRight.width;
    h = paneLayout.quadrants.topRight.height;
    if (w > 0 && h > 0) {
      queryRect.x = viewport.scrollX + frozenContentWidth;
      queryRect.y = 0;
      queryRect.width = w;
      queryRect.height = h;
      const candidates = this.spatialIndex.query(queryRect, candidatesScratch);
      for (const obj of candidates) {
        const anchor = obj.anchor;
        if (anchor.type === "absolute") continue;
        if (anchor.from.cell.row < frozenRows && anchor.from.cell.col >= frozenCols) ordered.push(obj);
      }
    }

    w = paneLayout.quadrants.bottomLeft.width;
    h = paneLayout.quadrants.bottomLeft.height;
    if (w > 0 && h > 0) {
      queryRect.x = 0;
      queryRect.y = viewport.scrollY + frozenContentHeight;
      queryRect.width = w;
      queryRect.height = h;
      const candidates = this.spatialIndex.query(queryRect, candidatesScratch);
      for (const obj of candidates) {
        const anchor = obj.anchor;
        if (anchor.type === "absolute") continue;
        if (anchor.from.cell.row >= frozenRows && anchor.from.cell.col < frozenCols) ordered.push(obj);
      }
    }

    w = paneLayout.quadrants.bottomRight.width;
    h = paneLayout.quadrants.bottomRight.height;
    if (w > 0 && h > 0) {
      queryRect.x = viewport.scrollX + frozenContentWidth;
      queryRect.y = viewport.scrollY + frozenContentHeight;
      queryRect.width = w;
      queryRect.height = h;
      const candidates = this.spatialIndex.query(queryRect, candidatesScratch);
      for (const obj of candidates) {
        const anchor = obj.anchor;
        if (anchor.type === "absolute") {
          ordered.push(obj);
          continue;
        }
        if (anchor.from.cell.row >= frozenRows && anchor.from.cell.col >= frozenCols) ordered.push(obj);
      }
    }

    // Release references from the query scratch array before drawing (we no longer need it).
    candidatesScratch.length = 0;

    const selectedId = this.selectedId;
    let selectedScreenRect: Rect | null = null;
    let selectedClipRect: Rect | null = null;
    let selectedAabb: Rect | null = null;
    let selectedTransform: DrawingTransform | undefined = undefined;
    const screenRectScratch = this.screenRectScratch;
    const aabbScratch = this.aabbScratch;
    const localRectScratch = this.localRectScratch;
    let selectedDrawRotationHandle = true;
    const headerOffsetX = paneLayout.headerOffsetX;
    const headerOffsetY = paneLayout.headerOffsetY;
    const scrollXBase = viewport.scrollX;
    const scrollYBase = viewport.scrollY;
    const qTopLeft = paneLayout.quadrants.topLeft;
    const qTopRight = paneLayout.quadrants.topRight;
    const qBottomLeft = paneLayout.quadrants.bottomLeft;
    const qBottomRight = paneLayout.quadrants.bottomRight;

    if (drawObjects) {
      // First pass: kick off image decodes for all visible images without awaiting so
      // multiple images can decode concurrently on cold render.
      for (const obj of ordered) {
        if (obj.kind.type !== "image") continue;

        const rect = this.spatialIndex.getRect(obj.id) ?? anchorToRectPx(obj.anchor, this.geom, zoom);
        // Zero-size drawings are invisible and do not need bitmap decoding. Avoid starting an image
        // decode (`createImageBitmap`) when there is no renderable area.
        if (rect.width <= 0 || rect.height <= 0) continue;
        const anchor = obj.anchor;
        let scrollX = scrollXBase;
        let scrollY = scrollYBase;
        let clipRect = qBottomRight;
        if (anchor.type !== "absolute") {
          const inFrozenRows = anchor.from.cell.row < frozenRows;
          const inFrozenCols = anchor.from.cell.col < frozenCols;
          scrollX = inFrozenCols ? 0 : scrollXBase;
          scrollY = inFrozenRows ? 0 : scrollYBase;
          if (inFrozenRows) {
            clipRect = inFrozenCols ? qTopLeft : qTopRight;
          } else {
            clipRect = inFrozenCols ? qBottomLeft : qBottomRight;
          }
        }
        screenRectScratch.x = rect.x - scrollX + headerOffsetX;
        screenRectScratch.y = rect.y - scrollY + headerOffsetY;
        screenRectScratch.width = rect.width;
        screenRectScratch.height = rect.height;
        // Zero-size drawings can exist transiently (tests, partially-hydrated/corrupt documents).
        // Skip decoding/raster work for them since there's nothing to draw.
        if (!(screenRectScratch.width > 0 && screenRectScratch.height > 0)) continue;
        const aabb = getAabbForObject(screenRectScratch, obj.transform, aabbScratch);

        // Avoid starting decode work for degenerate drawings (0x0 sized images).
        // See the similar guard in the per-object render loop below.
        if (!Number.isFinite(screenRectScratch.width) || !Number.isFinite(screenRectScratch.height)) continue;
        if (screenRectScratch.width <= 0 || screenRectScratch.height <= 0) continue;

        if (clipRect.width <= 0 || clipRect.height <= 0) continue;
        if (!intersects(aabb, clipRect)) continue;
        if (!intersects(clipRect, viewportRect)) continue;
        if (!intersects(aabb, viewportRect)) continue;

        const imageId = obj.kind.imageId;
        const entry = this.images.get(imageId);
        if (!entry) {
          // Start best-effort hydration early so async image byte loading can overlap bitmap decode.
          this.hydrateImage(imageId);
          continue;
        }
        this.bitmapCache.getOrRequest(entry, this.onBitmapReady);
      }

      for (const obj of ordered) {
        const rect = this.spatialIndex.getRect(obj.id) ?? anchorToRectPx(obj.anchor, this.geom, zoom);
        if (rect.width <= 0 || rect.height <= 0) continue;
        const anchor = obj.anchor;
        let scrollX = scrollXBase;
        let scrollY = scrollYBase;
        let clipRect = qBottomRight;
        if (anchor.type !== "absolute") {
          const inFrozenRows = anchor.from.cell.row < frozenRows;
          const inFrozenCols = anchor.from.cell.col < frozenCols;
          scrollX = inFrozenCols ? 0 : scrollXBase;
          scrollY = inFrozenRows ? 0 : scrollYBase;
          if (inFrozenRows) {
            clipRect = inFrozenCols ? qTopLeft : qTopRight;
          } else {
            clipRect = inFrozenCols ? qBottomLeft : qBottomRight;
          }
        }
        screenRectScratch.x = rect.x - scrollX + headerOffsetX;
        screenRectScratch.y = rect.y - scrollY + headerOffsetY;
        screenRectScratch.width = rect.width;
        screenRectScratch.height = rect.height;
        if (!(screenRectScratch.width > 0 && screenRectScratch.height > 0)) continue;
        const aabb = getAabbForObject(screenRectScratch, obj.transform, aabbScratch);

        if (selectedId != null && obj.id === selectedId) {
          const selScreen = this.selectedScreenRectScratch;
          selScreen.x = screenRectScratch.x;
          selScreen.y = screenRectScratch.y;
          selScreen.width = screenRectScratch.width;
          selScreen.height = screenRectScratch.height;
          selectedScreenRect = selScreen;
          selectedClipRect = clipRect;
          const selAabb = this.selectedAabbScratch;
          selAabb.x = aabb.x;
          selAabb.y = aabb.y;
          selAabb.width = aabb.width;
          selAabb.height = aabb.height;
          selectedAabb = selAabb;
          selectedTransform = obj.transform;
          // Charts behave like Excel chart objects: movable/resizable but not rotatable.
          selectedDrawRotationHandle = obj.kind.type !== "chart";
        }

        // Skip degenerate objects; they are invisible and do not need decoding or placeholder rendering.
        if (screenRectScratch.width <= 0 || screenRectScratch.height <= 0) continue;

        if (clipRect.width <= 0 || clipRect.height <= 0) continue;
        // Skip objects that are fully outside of their pane quadrant.
        if (!intersects(aabb, clipRect)) continue;
        // Paranoia: clip rects are expected to be within the viewport, but keep the
        // early-out for callers providing custom layouts.
        if (!intersects(clipRect, viewportRect)) continue;

        if (!intersects(aabb, viewportRect)) {
          continue;
        }

        if (obj.kind.type === "image") {
          // Avoid kicking off async image decoding work for degenerate drawings.
          // Zero-size drawing objects can show up in corrupted/imported documents and in unit tests.
          // `ctx.drawImage` would be a no-op, so skip decoding to keep render work bounded (and keep
          // tests that stub `createImageBitmap` deterministic).
          if (!Number.isFinite(screenRectScratch.width) || !Number.isFinite(screenRectScratch.height)) continue;
          if (screenRectScratch.width <= 0 || screenRectScratch.height <= 0) continue;
          const imageId = obj.kind.imageId;
          const entry = this.images.get(imageId);
          if (!entry) {
            // Best-effort: hydrate missing bytes from a persistent store (e.g. IndexedDB). Do not
            // await here: awaiting would delay placeholder rendering until the async lookup
            // resolves, making the object temporarily disappear. Instead, render a placeholder
            // immediately and schedule a follow-up render when bytes arrive.
            this.hydrateImage(imageId);

            // Image metadata can arrive before the bytes are hydrated into the ImageStore
            // (e.g. collaboration metadata received before IndexedDB hydration). Render a
            // placeholder box so the image remains visible/selectable until bytes load.
            pushClipRect(ctx, clipRect);
            try {
              ctx.strokeStyle = colors.placeholderOtherStroke;
              ctx.lineWidth = 1;
              ctx.setLineDash(LINE_DASH_PLACEHOLDER);
              if (hasNonIdentityTransform(obj.transform)) {
                drawTransformedRect(ctx, screenRectScratch, obj.transform);
              } else {
                ctx.strokeRect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
              }
              ctx.setLineDash(LINE_DASH_NONE);
              ctx.fillStyle = colors.placeholderLabel;
              ctx.globalAlpha = 0.6;
              ctx.font = "12px sans-serif";
              ctx.fillText("missing image", screenRectScratch.x + 4, screenRectScratch.y + 14);
            } finally {
              ctx.restore();
            }
            continue;
          }

          const bitmap =
            screenRectScratch.width > 0 && screenRectScratch.height > 0
              ? this.bitmapCache.getOrRequest(entry, this.onBitmapReady)
              : null;
          if (bitmap) {
            const picXml = obj.preserved?.["xlsx.pic_xml"];
            const props = typeof picXml === "string" && picXml.length > 0 ? getDrawingMLPicturePropsFromPicXml(picXml) : null;
            const crop = props?.crop;
            const outline = props?.outline;

            // Compute crop in source space (DrawingML `srcRect` uses 0..100000 percentages).
            let hasCrop = false;
            let sx = 0;
            let sy = 0;
            let sw = 0;
            let sh = 0;
            if (crop) {
              const bw = (bitmap as any).width;
              const bh = (bitmap as any).height;
              if (Number.isFinite(bw) && Number.isFinite(bh) && bw > 0 && bh > 0) {
                const left = (bw * crop.l) / 100_000;
                const top = (bh * crop.t) / 100_000;
                const right = (bw * crop.r) / 100_000;
                const bottom = (bh * crop.b) / 100_000;
                const croppedW = bw - left - right;
                const croppedH = bh - top - bottom;
                if (croppedW > 0 && croppedH > 0) {
                  hasCrop = true;
                  sx = left;
                  sy = top;
                  sw = croppedW;
                  sh = croppedH;
                }
              }
            }

            // Compute outline style (DrawingML uses EMU for width; 1pt = 12700 EMU).
            const outlineColor = outline?.color;
            const outlineWidthPx =
              outline?.widthEmu != null && Number.isFinite(outline.widthEmu) ? emuToPx(outline.widthEmu, zoom) : 1;
            const hasOutline = Boolean(outlineColor) && Number.isFinite(outlineWidthPx) && outlineWidthPx > 0;

            pushClipRect(ctx, clipRect);
            try {
              if (hasNonIdentityTransform(obj.transform)) {
                pushObjectTransform(ctx, screenRectScratch, obj.transform, localRectScratch);
                try {
                  // Clip to the (possibly rotated/flipped) image bounds so we don't
                  // overdraw neighboring cells when transforms extend beyond the anchor.
                  //
                  // Use an extra save/restore around the clip so we can draw the picture
                  // outline (if present) without being affected by the clip.
                  ctx.save();
                  try {
                    ctx.beginPath();
                    ctx.rect(localRectScratch.x, localRectScratch.y, localRectScratch.width, localRectScratch.height);
                    ctx.clip();
                    if (hasCrop) {
                      ctx.drawImage(
                        bitmap,
                        sx,
                        sy,
                        sw,
                        sh,
                        localRectScratch.x,
                        localRectScratch.y,
                        localRectScratch.width,
                        localRectScratch.height,
                      );
                    } else {
                      ctx.drawImage(
                        bitmap,
                        localRectScratch.x,
                        localRectScratch.y,
                        localRectScratch.width,
                        localRectScratch.height,
                      );
                    }
                  } finally {
                    ctx.restore();
                  }

                  if (hasOutline) {
                    ctx.strokeStyle = outlineColor!;
                    ctx.lineWidth = outlineWidthPx;
                    ctx.setLineDash(LINE_DASH_NONE);
                    ctx.strokeRect(localRectScratch.x, localRectScratch.y, localRectScratch.width, localRectScratch.height);
                  }
                } finally {
                  ctx.restore();
                }
              } else {
                if (hasCrop) {
                  ctx.drawImage(
                    bitmap,
                    sx,
                    sy,
                    sw,
                    sh,
                    screenRectScratch.x,
                    screenRectScratch.y,
                    screenRectScratch.width,
                    screenRectScratch.height,
                  );
                } else {
                  ctx.drawImage(bitmap, screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
                }

                if (hasOutline) {
                  ctx.strokeStyle = outlineColor!;
                  ctx.lineWidth = outlineWidthPx;
                  ctx.setLineDash(LINE_DASH_NONE);
                  ctx.strokeRect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
                }
              }
            } finally {
              ctx.restore();
            }
            continue;
          }

          // Fall through to placeholder rendering while decode is in-flight.
        }

        if (obj.kind.type === "chart") {
          const chartId = obj.kind.chartId;
          if (this.chartRenderer && typeof chartId === "string" && chartId.length > 0) {
            let rendered = false;
            pushClipRect(ctx, clipRect);
            try {
              try {
                if (hasNonIdentityTransform(obj.transform)) {
                  pushObjectTransform(ctx, screenRectScratch, obj.transform, localRectScratch);
                  try {
                    ctx.beginPath();
                    ctx.rect(localRectScratch.x, localRectScratch.y, localRectScratch.width, localRectScratch.height);
                    ctx.clip();
                    this.chartRenderer!.renderToCanvas(ctx, chartId, localRectScratch);
                    rendered = true;
                  } finally {
                    ctx.restore();
                  }
                } else {
                  ctx.beginPath();
                  ctx.rect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
                  ctx.clip();
                  this.chartRenderer!.renderToCanvas(ctx, chartId, screenRectScratch);
                  rendered = true;
                }
              } catch {
                rendered = false;
              }
            } finally {
              ctx.restore();
            }

            if (rendered) continue;
          }
        }

        if (obj.kind.type === "shape") {
          const rawXml = (obj.kind as any).rawXml ?? (obj.kind as any).raw_xml;
          const rawXmlText = typeof rawXml === "string" ? rawXml : "";

          // Parse `<xdr:txBody>` / shape render spec once and cache; avoid reparsing XML on every frame.
          let cachedText = this.shapeTextCache.get(obj.id);
          if (!cachedText || cachedText.rawXml !== rawXmlText) {
            const parsed = parseDrawingMLShapeText(rawXmlText);
            const hasText =
              parsed != null &&
              parsed.textRuns.some((run) => {
                const text = run.text;
                return typeof text === "string" && /\S/.test(text);
              });
            let spec: ShapeRenderSpec | null = null;
            try {
              spec = rawXmlText ? parseShapeRenderSpec(rawXmlText) : null;
            } catch {
              spec = null;
            }
            cachedText = {
              rawXml: rawXmlText,
              parsed,
              hasText,
              spec,
              lines: null,
              linesMaxWidth: Number.NaN,
              linesWrap: true,
              linesZoom: Number.NaN,
              linesDefaultColor: "",
            };
            this.shapeTextCache.set(obj.id, cachedText);
          }
          const textLayout = cachedText.parsed;
          const textParsed = textLayout !== null;
          const hasText = cachedText.hasText;
          const canRenderText = hasText && typeof (ctx as any).measureText === "function";

          let rendered = false;
          const spec = cachedText.spec;

          if (spec) {
            pushClipRect(ctx, clipRect);
            try {
              try {
                if (hasNonIdentityTransform(obj.transform)) {
                  pushObjectTransform(ctx, screenRectScratch, obj.transform, localRectScratch);
                  try {
                    drawShape(ctx, localRectScratch, spec, colors, cssVarStyle, zoom, this.dashPatternScratch, canRenderText ? null : undefined);
                    if (canRenderText) {
                      renderShapeText(ctx, localRectScratch, textLayout!, { defaultColor: colors.placeholderLabel, zoom }, cachedText);
                    }
                  } finally {
                    ctx.restore();
                  }
                } else {
                  drawShape(ctx, screenRectScratch, spec, colors, cssVarStyle, zoom, this.dashPatternScratch, canRenderText ? null : undefined);
                  if (canRenderText) {
                    renderShapeText(ctx, screenRectScratch, textLayout!, { defaultColor: colors.placeholderLabel, zoom }, cachedText);
                  }
                }
                rendered = true;
              } catch {
                rendered = false;
              }
            } finally {
              ctx.restore();
            }
            if (rendered) continue;
          }

          // If we couldn't render the shape geometry but we did successfully parse text,
          // still render the text within the anchored bounds (and skip placeholders).
          if (canRenderText) {
            pushClipRect(ctx, clipRect);
            try {
              if (hasNonIdentityTransform(obj.transform)) {
                pushObjectTransform(ctx, screenRectScratch, obj.transform, localRectScratch);
                try {
                  ctx.beginPath();
                  ctx.rect(localRectScratch.x, localRectScratch.y, localRectScratch.width, localRectScratch.height);
                  ctx.clip();
                  renderShapeText(ctx, localRectScratch, textLayout!, { defaultColor: colors.placeholderLabel, zoom }, cachedText);
                } finally {
                  ctx.restore();
                }
              } else {
                ctx.beginPath();
                ctx.rect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
                ctx.clip();
                renderShapeText(ctx, screenRectScratch, textLayout!, { defaultColor: colors.placeholderLabel, zoom }, cachedText);
              }
            } finally {
              ctx.restore();
            }
            continue;
          }

          // Shape parsed but has no text: keep an empty bounds placeholder (no label).
          if (textParsed) {
            pushClipRect(ctx, clipRect);
            try {
              ctx.strokeStyle = colors.placeholderOtherStroke;
              ctx.lineWidth = 1;
              ctx.setLineDash(LINE_DASH_PLACEHOLDER);
              if (hasNonIdentityTransform(obj.transform)) {
                drawTransformedRect(ctx, screenRectScratch, obj.transform);
              } else {
                ctx.strokeRect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
              }
            } finally {
              ctx.restore();
            }
            continue;
          }
        }

        // Placeholder rendering for shapes/charts/unknown.
        pushClipRect(ctx, clipRect);
        try {
          const rawXml =
            // Some integration layers still pass through snake_case from the Rust model.
            (obj.kind as any).rawXml ?? (obj.kind as any).raw_xml;
          const isGFrame = isGraphicFrame(rawXml);
          const isUnknown = obj.kind.type === "unknown";

          ctx.strokeStyle =
            obj.kind.type === "chart"
              ? colors.placeholderChartStroke
              : isUnknown || isGFrame
                ? colors.placeholderGraphicFrameStroke
                : colors.placeholderOtherStroke;
          ctx.lineWidth = 1;
          ctx.setLineDash(LINE_DASH_PLACEHOLDER);
          if (hasNonIdentityTransform(obj.transform)) {
            drawTransformedRect(ctx, screenRectScratch, obj.transform);
          } else {
            ctx.strokeRect(screenRectScratch.x, screenRectScratch.y, screenRectScratch.width, screenRectScratch.height);
          }
          ctx.setLineDash(LINE_DASH_NONE);
          ctx.fillStyle = colors.placeholderLabel;
          ctx.globalAlpha = 0.6;
          ctx.font = "12px sans-serif";
          const explicitLabel = obj.kind.label?.trim();
          const placeholderLabel =
            explicitLabel && explicitLabel.length > 0
              ? explicitLabel
              : // Avoid labeling chart placeholders as "GraphicFrame" — charts already have a distinct kind.
                obj.kind.type !== "chart"
                ? graphicFramePlaceholderLabel(rawXml) ?? obj.kind.type
                : obj.kind.type;
          ctx.fillText(placeholderLabel, screenRectScratch.x + 4, screenRectScratch.y + 14);
        } finally {
          ctx.restore();
        }
      }
    }

    // Selection overlay.
    if (selectedId != null) {
      if (selectedScreenRect && selectedClipRect && selectedAabb) {
        if (selectedClipRect.width > 0 && selectedClipRect.height > 0 && intersects(selectedAabb, selectedClipRect)) {
          pushClipRect(ctx, selectedClipRect);
          try {
            drawSelection(
              ctx,
              selectedScreenRect,
              colors,
              selectedTransform,
              selectedDrawRotationHandle,
              this.resizeHandleCentersScratch,
              this.rotationHandleCenterScratch,
            );
          } finally {
            ctx.restore();
          }
        }
      } else {
        // Fallback: selection can still be rendered when `drawObjects` is disabled.
        let selected: DrawingObject | null = null;
        for (let i = 0; i < ordered.length; i += 1) {
          const obj = ordered[i]!;
          if (obj.id === selectedId) {
            selected = obj;
            break;
          }
        }
        if (selected) {
          const rect = this.spatialIndex.getRect(selected.id) ?? anchorToRectPx(selected.anchor, this.geom, zoom);
          const anchor = selected.anchor;
          let scrollX = scrollXBase;
          let scrollY = scrollYBase;
          let clipRect = qBottomRight;
          if (anchor.type !== "absolute") {
            const inFrozenRows = anchor.from.cell.row < frozenRows;
            const inFrozenCols = anchor.from.cell.col < frozenCols;
            scrollX = inFrozenCols ? 0 : scrollXBase;
            scrollY = inFrozenRows ? 0 : scrollYBase;
            if (inFrozenRows) {
              clipRect = inFrozenCols ? qTopLeft : qTopRight;
            } else {
              clipRect = inFrozenCols ? qBottomLeft : qBottomRight;
            }
          }

          screenRectScratch.x = rect.x - scrollX + headerOffsetX;
          screenRectScratch.y = rect.y - scrollY + headerOffsetY;
          screenRectScratch.width = rect.width;
          screenRectScratch.height = rect.height;

          const selectionAabb = getAabbForObject(screenRectScratch, selected.transform, aabbScratch);
          if (clipRect.width > 0 && clipRect.height > 0 && intersects(selectionAabb, clipRect)) {
            pushClipRect(ctx, clipRect);
            try {
              drawSelection(
                ctx,
                screenRectScratch,
                colors,
                selected.transform,
                selected.kind.type !== "chart",
                this.resizeHandleCentersScratch,
                this.rotationHandleCenterScratch,
              );
            } finally {
              ctx.restore();
            }
          }
        }
      }
    }

    completed = true;
    } finally {
      // Release object references eagerly so we don't accidentally retain the most recently visible
      // object set when rendering pauses.
      this.visibleObjectsScratch.length = 0;
      this.queryCandidatesScratch.length = 0;
      // Prune cached shape text layouts for shapes that no longer exist.
      //
      if (completed && this.shapeTextCache.size > 0) {
        // `shapeTextCache` is keyed by drawing id and can otherwise grow unbounded across
        // delete/undo/redo sessions. Avoid per-frame allocations by only pruning when the
        // object list changes (or appears to have been mutated in place).
        const sourceChanged =
          this.shapeTextCachePruneSource !== objects || this.shapeTextCachePruneLength !== objects.length;
        this.shapeTextCachePruneSource = objects;
        this.shapeTextCachePruneLength = objects.length;

        if (sourceChanged) {
          for (const id of this.shapeTextCache.keys()) {
            const obj = this.spatialIndex.getObject(id);
            if (!obj || obj.kind.type !== "shape") {
              this.shapeTextCache.delete(id);
            }
          }
        }
      }
    }
  }

  /**
   * Optional selection id used for rendering handles.
   *
   * This is intentionally not part of the constructor API so the overlay can be
   * integrated incrementally with different state management approaches.
   */
  setSelectedId(id: number | null): void {
    this.selectedId = id;
  }

  /**
   * Eagerly decode an ImageEntry into an ImageBitmap.
   *
   * Intended for insertion flows so the first render after inserting a picture
   * can reuse an already-decoding promise.
   */
  preloadImage(entry: ImageEntry): Promise<ImageBitmap> {
    if (this.destroyed) {
      // Treat preloads as best-effort and allow callers to treat teardown as an abort.
      return Promise.reject(createAbortError());
    }
    // Preloads are optional / best-effort and can be invalidated by:
    // - overlay destruction (switching workbooks, hot reload, tests)
    // - cache clears/invalidation (applyState/image updates)
    //
    // Use an AbortSignal so that if the overlay is torn down before the decode resolves, any
    // resulting ImageBitmap is deterministically closed instead of being returned to a dropped
    // Promise chain (which would leak the decoded bitmap in memory).
    const signal = this.preloadAbort?.signal;
    this.preloadCount += 1;
    const p = this.bitmapCache.get(entry, signal ? { signal } : undefined);
    return p.finally(() => {
      this.preloadCount = Math.max(0, this.preloadCount - 1);
    });
  }

  /**
   * Invalidate a decoded bitmap for an image id.
   *
   * This is used by higher-level stores (e.g. DocumentController-backed images) when the
   * underlying bytes for an existing image id change (undo/redo, overwrite, etc).
   */
  invalidateImage(imageId: string): void {
    if (this.destroyed) return;
    // If the bitmap bytes change while we are mid-render or mid-preload, the old decode result can
    // arrive after the cache entry has been invalidated. Abort any in-flight consumers first so the
    // stale ImageBitmap is closed when the decode eventually resolves.
    if (this.preloadCount > 0) {
      this.preloadAbort?.abort();
      this.preloadAbort = typeof AbortController !== "undefined" ? new AbortController() : null;
    }
    const id = String(imageId ?? "");
    this.imageHydrationNegativeCache.delete(id);
    this.bitmapCache.invalidate(id);
  }

  /**
   * Clear all cached decoded bitmaps.
   *
   * Useful after loading a new workbook snapshot where all image ids/bytes may change.
   */
  clearImageCache(): void {
    if (this.destroyed) return;
    // Bump the epoch so any async hydrations that were already queued (or in-flight)
    // do not repopulate the caches after we've cleared them.
    this.imageHydrationEpoch += 1;
    // When callers clear the cache (e.g. applying a new document snapshot), ensure any in-flight
    // decodes from older renders/preloads don't leak their ImageBitmaps after the cache entry is
    // dropped.
    this.preloadAbort?.abort();
    this.preloadAbort = typeof AbortController !== "undefined" ? new AbortController() : null;
    this.bitmapCache.clear();
    this.pendingImageHydrations.clear();
    this.imageHydrationNegativeCache.clear();
  }

  destroy(): void {
    // Cancel any in-flight render and release cached bitmap resources.
    this.destroyed = true;
    this.requestRender = null;
    this.preloadAbort?.abort();
    this.preloadAbort = null;
    this.chartRenderer?.destroy?.();
    this.themeObserver?.disconnect();
    this.themeObserver = null;
    this.bitmapCache.clear();
    this.shapeTextCache.clear();
    this.shapeTextCachePruneSource = null;
    this.shapeTextCachePruneLength = 0;
    this.resizeMemo = null;
    this.chartSurfacePruneSource = null;
    this.chartSurfacePruneLength = 0;
    this.chartSurfaceKeep.clear();
    // Release any cached drawing object references (spatial index stores objects + rects and
    // scratch arrays referencing bucket lists from the last query).
    this.spatialIndex.dispose();
    this.cssVarStyle = undefined;
    this.colorTokens = null;
    this.selectedId = null;
    this.lastRenderArgs = null;
    this.pendingImageHydrations.clear();
    this.imageHydrationNegativeCache.clear();
    this.hydrationRerenderScheduled = false;
    this.visibleObjectsScratch.length = 0;
    this.queryCandidatesScratch.length = 0;
    this.resizeHandleCentersScratch.length = 0;

    // Release the canvas backing store even if the DOM element is still referenced
    // by a long-lived owner (tests/hot reload). Setting width/height resets the
    // internal bitmap allocation.
    try {
      this.canvas.width = 0;
      this.canvas.height = 0;
    } catch {
      // Best-effort: ignore canvas reset failures (e.g. mocked canvases).
    }
  }

  /**
   * Alias for `destroy()` (matches other UI controller teardown semantics).
   */
  dispose(): void {
    this.destroy();
  }
}

function createAbortError(): Error {
  const err = new Error("The operation was aborted.");
  err.name = "AbortError";
  return err;
}

function intersects(a: Rect, b: Rect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): transform is DrawingTransform {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

type CachedTrig = { rotationDeg: number; cos: number; sin: number };

const overlayTrigCache = new WeakMap<DrawingTransform, CachedTrig>();

function getTransformTrig(transform: DrawingTransform): CachedTrig {
  const cached = overlayTrigCache.get(transform);
  const rot = transform.rotationDeg;
  if (cached && cached.rotationDeg === rot) return cached;
  const radians = degToRad(rot);
  const next: CachedTrig = { rotationDeg: rot, cos: Math.cos(radians), sin: Math.sin(radians) };
  overlayTrigCache.set(transform, next);
  return next;
}

function rectToAabb(rect: Rect, transform: DrawingTransform, out?: Rect): Rect {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;

  const trig = getTransformTrig(transform);
  const cos = trig.cos;
  const sin = trig.sin;

  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  // Corner 1: (-hw, -hh)
  let x = -hw;
  let y = -hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  let wx = cx + (x * cos - y * sin);
  let wy = cy + (x * sin + y * cos);
  if (wx < minX) minX = wx;
  if (wx > maxX) maxX = wx;
  if (wy < minY) minY = wy;
  if (wy > maxY) maxY = wy;

  // Corner 2: (hw, -hh)
  x = hw;
  y = -hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  wx = cx + (x * cos - y * sin);
  wy = cy + (x * sin + y * cos);
  if (wx < minX) minX = wx;
  if (wx > maxX) maxX = wx;
  if (wy < minY) minY = wy;
  if (wy > maxY) maxY = wy;

  // Corner 3: (hw, hh)
  x = hw;
  y = hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  wx = cx + (x * cos - y * sin);
  wy = cy + (x * sin + y * cos);
  if (wx < minX) minX = wx;
  if (wx > maxX) maxX = wx;
  if (wy < minY) minY = wy;
  if (wy > maxY) maxY = wy;

  // Corner 4: (-hw, hh)
  x = -hw;
  y = hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  wx = cx + (x * cos - y * sin);
  wy = cy + (x * sin + y * cos);
  if (wx < minX) minX = wx;
  if (wx > maxX) maxX = wx;
  if (wy < minY) minY = wy;
  if (wy > maxY) maxY = wy;

  const target = out ?? { x: 0, y: 0, width: 0, height: 0 };
  target.x = minX;
  target.y = minY;
  target.width = maxX - minX;
  target.height = maxY - minY;
  return target;
}

function getAabbForObject(rect: Rect, transform: DrawingTransform | undefined, out?: Rect): Rect {
  if (!hasNonIdentityTransform(transform)) return rect;
  return rectToAabb(rect, transform, out);
}

function pushClipRect(ctx: CanvasRenderingContext2D, clipRect: Rect): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
  ctx.clip();
}

function pushObjectTransform(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  transform: DrawingTransform,
  out: Rect,
): void {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;

  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(degToRad(transform.rotationDeg));
  ctx.scale(transform.flipH ? -1 : 1, transform.flipV ? -1 : 1);

  out.x = -rect.width / 2;
  out.y = -rect.height / 2;
  out.width = rect.width;
  out.height = rect.height;
}

function drawTransformedRect(ctx: CanvasRenderingContext2D, rect: Rect, transform: DrawingTransform): void {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;

  const trig = getTransformTrig(transform);
  const cos = trig.cos;
  const sin = trig.sin;

  ctx.beginPath();
  let x = -hw;
  let y = -hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  ctx.moveTo(cx + (x * cos - y * sin), cy + (x * sin + y * cos));

  x = hw;
  y = -hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  ctx.lineTo(cx + (x * cos - y * sin), cy + (x * sin + y * cos));

  x = hw;
  y = hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  ctx.lineTo(cx + (x * cos - y * sin), cy + (x * sin + y * cos));

  x = -hw;
  y = hh;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  ctx.lineTo(cx + (x * cos - y * sin), cy + (x * sin + y * cos));
  ctx.closePath();
  ctx.stroke();
}

function drawSelection(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  colors: OverlayColorTokens,
  transform: DrawingTransform | undefined,
  drawRotationHandle: boolean,
  resizeHandlesScratch: ResizeHandleCenter[],
  rotationHandleScratch: RotationHandleCenter,
): void {
  ctx.save();
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 2;
  ctx.setLineDash(LINE_DASH_NONE);
  if (hasNonIdentityTransform(transform)) {
    drawTransformedRect(ctx, rect, transform);
  } else {
    ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
  }

  const handle = RESIZE_HANDLE_SIZE_PX;
  const half = handle / 2;
  const points = getResizeHandleCentersInto(rect, transform, resizeHandlesScratch);

  ctx.fillStyle = colors.selectionHandleFill;
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 1;
  for (let i = 0; i < points.length; i += 1) {
    const p = points[i]!;
    ctx.beginPath();
    ctx.rect(p.x - half, p.y - half, handle, handle);
    ctx.fill();
    ctx.stroke();
  }

  // Optional Excel-style rotation handle.
  if (drawRotationHandle) {
    const rotHandle = ROTATION_HANDLE_SIZE_PX;
    const rotHalf = rotHandle / 2;
    const rot = getRotationHandleCenterInto(rect, transform, rotationHandleScratch);
    ctx.beginPath();
    ctx.rect(rot.x - rotHalf, rot.y - rotHalf, rotHandle, rotHandle);
    ctx.fill();
    ctx.stroke();
  }
  ctx.restore();
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  spec: ShapeRenderSpec,
  colors: OverlayColorTokens,
  cssVarStyle: CssVarStyle | null,
  zoom: number,
  dashScratch: number[],
  labelOverride?: string | null,
): void {
  // Clip to the anchored bounds; this matches the chart rendering behaviour and
  // avoids accidental overdraw if we misinterpret a shape transform.
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();

  const scale = zoom;
  const strokeWidthPx = spec.stroke ? emuToPx(spec.stroke.widthEmu, zoom) : 0;
  const inset = strokeWidthPx > 0 ? strokeWidthPx / 2 : 0;
  const x = rect.x + inset;
  const y = rect.y + inset;
  const w = Math.max(0, rect.width - inset * 2);
  const h = Math.max(0, rect.height - inset * 2);

  ctx.beginPath();
  switch (spec.geometry.type) {
    case "rect":
      ctx.rect(x, y, w, h);
      break;
    case "roundRect": {
      // Excel encodes round-rect corner rounding as an adjustment value in the
      // 0-100000 range (with 50000 roughly representing max rounding).
      const adj = spec.geometry.adj;
      const ratio =
        typeof adj === "number" && Number.isFinite(adj) ? Math.max(0, Math.min(50_000, adj)) / 100_000 : 0.2;
      const radius = Math.min(w, h) * ratio;
      roundRectPath(ctx, x, y, w, h, radius);
      break;
    }
    case "ellipse":
      ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
      break;
    case "line":
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y + h);
      break;
  }

  if (spec.geometry.type !== "line" && spec.fill.type === "solid") {
    ctx.fillStyle = resolveCanvasColor(cssVarStyle, spec.fill.color, colors.placeholderLabel);
    ctx.fill();
  }

  if (spec.stroke && strokeWidthPx > 0) {
    ctx.strokeStyle = resolveCanvasColor(cssVarStyle, spec.stroke.color, colors.placeholderLabel);
    ctx.lineWidth = strokeWidthPx;
    ctx.setLineDash(dashPatternForPreset(spec.stroke.dashPreset, strokeWidthPx, dashScratch));
    ctx.stroke();
  }

  const label = labelOverride !== undefined ? labelOverride : spec.label;
  if (label) {
    ctx.fillStyle = spec.labelColor ?? colors.placeholderLabel;
    ctx.globalAlpha = 0.8;
    const size =
      typeof spec.labelFontSizePx === "number" && Number.isFinite(spec.labelFontSizePx) ? spec.labelFontSizePx * scale : 12 * scale;
    const family = spec.labelFontFamily?.trim() ? spec.labelFontFamily : "sans-serif";
    const weight = spec.labelBold ? "bold " : "";
    ctx.font = `${weight}${size}px ${family}`;
    const align = spec.labelAlign ?? "left";
    const vAlign = spec.labelVAlign ?? "top";
    ctx.textAlign = align;
    ctx.textBaseline = vAlign === "middle" ? "middle" : vAlign;
    const padding = 4 * scale;
    const xText =
      align === "center" ? rect.x + rect.width / 2 : align === "right" ? rect.x + rect.width - padding : rect.x + padding;
    const yText =
      vAlign === "middle" ? rect.y + rect.height / 2 : vAlign === "bottom" ? rect.y + rect.height - padding : rect.y + padding;
    ctx.fillText(label, xText, yText);
  }
}

function roundRectPath(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
): void {
  const r = Math.max(0, Math.min(radius, width / 2, height / 2));
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + width, y, x + width, y + height, r);
  ctx.arcTo(x + width, y + height, x, y + height, r);
  ctx.arcTo(x, y + height, x, y, r);
  ctx.arcTo(x, y, x + width, y, r);
  ctx.closePath();
}

function dashPatternForPreset(preset: string | undefined, strokeWidthPx: number, out: number[]): number[] {
  out.length = 0;
  if (!preset || preset === "solid") return out;
  const unit = Math.max(1, strokeWidthPx);
  switch (preset) {
    case "dash":
    case "sysDash": {
      out[0] = 4 * unit;
      out[1] = 2 * unit;
      out.length = 2;
      return out;
    }
    case "dot":
    case "sysDot": {
      out[0] = unit;
      out[1] = 2 * unit;
      out.length = 2;
      return out;
    }
    case "dashDot":
    case "sysDashDot": {
      out[0] = 4 * unit;
      out[1] = 2 * unit;
      out[2] = unit;
      out[3] = 2 * unit;
      out.length = 4;
      return out;
    }
    case "lgDash": {
      out[0] = 8 * unit;
      out[1] = 3 * unit;
      out.length = 2;
      return out;
    }
    case "lgDashDot": {
      out[0] = 8 * unit;
      out[1] = 3 * unit;
      out[2] = unit;
      out[3] = 3 * unit;
      out.length = 4;
      return out;
    }
    case "lgDashDotDot":
    case "sysDashDotDot": {
      out[0] = 8 * unit;
      out[1] = 3 * unit;
      out[2] = unit;
      out[3] = 3 * unit;
      out[4] = unit;
      out[5] = 3 * unit;
      out.length = 6;
      return out;
    }
    default:
      return out;
  }
}

type PaneQuadrant = "topLeft" | "topRight" | "bottomLeft" | "bottomRight";

type PaneLayout = {
  frozenRows: number;
  frozenCols: number;
  headerOffsetX: number;
  headerOffsetY: number;
  quadrants: Record<PaneQuadrant, Rect>;
};

const PANE_CELL_SCRATCH = { row: 0, col: 0 };

function clampNumber(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function resolvePaneLayout(
  viewport: Viewport,
  geom: GridGeometry,
  out: PaneLayout,
): PaneLayout {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  const cellAreaWidth = Math.max(0, viewport.width - headerOffsetX);
  const cellAreaHeight = Math.max(0, viewport.height - headerOffsetY);

  // `frozenWidthPx/HeightPx` are specified in viewport coordinates (they represent
  // the frozen boundary position). When omitted, derive them from the grid geometry
  // (sheet-space frozen extents) plus any header offset.
  let derivedFrozenContentWidth = 0;
  if (frozenCols > 0) {
    try {
      PANE_CELL_SCRATCH.row = 0;
      PANE_CELL_SCRATCH.col = frozenCols;
      derivedFrozenContentWidth = geom.cellOriginPx(PANE_CELL_SCRATCH).x;
    } catch {
      derivedFrozenContentWidth = 0;
    }
  }
  let derivedFrozenContentHeight = 0;
  if (frozenRows > 0) {
    try {
      PANE_CELL_SCRATCH.row = frozenRows;
      PANE_CELL_SCRATCH.col = 0;
      derivedFrozenContentHeight = geom.cellOriginPx(PANE_CELL_SCRATCH).y;
    } catch {
      derivedFrozenContentHeight = 0;
    }
  }

  // Only treat `frozenWidthPx/HeightPx` as meaningful when the corresponding frozen row/col
  // count is non-zero. This mirrors hit testing semantics and guards against stale pixel
  // extents if a caller updates frozen row/col counts but forgets to reset the boundary.
  const frozenBoundaryX =
    frozenCols > 0
      ? clampNumber(
          Number.isFinite(viewport.frozenWidthPx) ? viewport.frozenWidthPx! : headerOffsetX + derivedFrozenContentWidth,
          headerOffsetX,
          viewport.width,
        )
      : headerOffsetX;
  const frozenBoundaryY =
    frozenRows > 0
      ? clampNumber(
          Number.isFinite(viewport.frozenHeightPx) ? viewport.frozenHeightPx! : headerOffsetY + derivedFrozenContentHeight,
          headerOffsetY,
          viewport.height,
        )
      : headerOffsetY;

  const frozenContentWidth = clampNumber(frozenBoundaryX - headerOffsetX, 0, cellAreaWidth);
  const frozenContentHeight = clampNumber(frozenBoundaryY - headerOffsetY, 0, cellAreaHeight);
  const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
  const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

  const x0 = headerOffsetX;
  const y0 = headerOffsetY;
  const x1 = headerOffsetX + frozenContentWidth;
  const y1 = headerOffsetY + frozenContentHeight;

  out.frozenRows = frozenRows;
  out.frozenCols = frozenCols;
  out.headerOffsetX = headerOffsetX;
  out.headerOffsetY = headerOffsetY;
  const quads = out.quadrants;
  quads.topLeft.x = x0;
  quads.topLeft.y = y0;
  quads.topLeft.width = frozenContentWidth;
  quads.topLeft.height = frozenContentHeight;
  quads.topRight.x = x1;
  quads.topRight.y = y0;
  quads.topRight.width = scrollableWidth;
  quads.topRight.height = frozenContentHeight;
  quads.bottomLeft.x = x0;
  quads.bottomLeft.y = y1;
  quads.bottomLeft.width = frozenContentWidth;
  quads.bottomLeft.height = scrollableHeight;
  quads.bottomRight.x = x1;
  quads.bottomRight.y = y1;
  quads.bottomRight.width = scrollableWidth;
  quads.bottomRight.height = scrollableHeight;
  return out;
}

const DEFAULT_SHAPE_FONT_FAMILY = "Calibri, sans-serif";
const DEFAULT_SHAPE_FONT_SIZE_PT = 11;
const PX_PER_PT = PX_PER_INCH / 72;

type RenderedSegment = {
  text: string;
  run: ShapeTextRun;
  font: string;
  fontSizePx: number;
  color: string;
  underline: boolean;
  width: number;
};

type RenderedLine = {
  segments: RenderedSegment[];
  width: number;
  height: number;
};

function shapeRunFont(run: ShapeTextRun, zoom: number): { font: string; fontSizePx: number } {
  const scale = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const fontSizePt = run.fontSizePt ?? DEFAULT_SHAPE_FONT_SIZE_PT;
  const fontSizePx = Math.max(1, fontSizePt * PX_PER_PT * scale);
  const family = run.fontFamily?.trim() || DEFAULT_SHAPE_FONT_FAMILY;
  const parts: string[] = [];
  if (run.italic) parts.push("italic");
  if (run.bold) parts.push("bold");
  parts.push(`${fontSizePx}px`);
  parts.push(family);
  return { font: parts.join(" "), fontSizePx };
}

function runKey(run: ShapeTextRun): string {
  return [
    run.bold ? "b" : "",
    run.italic ? "i" : "",
    run.underline ? "u" : "",
    run.fontSizePt ?? "",
    run.fontFamily ?? "",
    run.color ?? "",
  ].join("|");
}

function layoutShapeTextLines(
  ctx: CanvasRenderingContext2D,
  runs: ShapeTextRun[],
  maxWidth: number,
  opts: { wrap: boolean; defaultColor: string; zoom: number },
): RenderedLine[] {
  const lines: RenderedLine[] = [];
  let segments: RenderedSegment[] = [];
  let width = 0;
  let maxFontSizePx = 0;
  const scale = Number.isFinite(opts.zoom) && opts.zoom > 0 ? opts.zoom : 1;
  const normalizeText = (text: string): string => (text.includes("\t") ? text.replaceAll("\t", "    ") : text);

  const flush = () => {
    if (segments.length === 0) {
      lines.push({ segments: [], width: 0, height: DEFAULT_SHAPE_FONT_SIZE_PT * PX_PER_PT * scale * 1.2 });
    } else {
      lines.push({
        segments,
        width,
        height: Math.max(1, maxFontSizePx) * 1.2,
      });
    }
    segments = [];
    width = 0;
    maxFontSizePx = 0;
  };

  const appendText = (text: string, run: ShapeTextRun, alreadyNormalized = false) => {
    const normalized = alreadyNormalized ? text : normalizeText(text);
    if (normalized === "") return;
    const { font, fontSizePx } = shapeRunFont(run, scale);
    ctx.font = font;
    const measured = ctx.measureText(normalized).width;
    const color = run.color ?? opts.defaultColor;
    const underline = Boolean(run.underline);

    maxFontSizePx = Math.max(maxFontSizePx, fontSizePx);

    const key = runKey(run);
    const prev = segments[segments.length - 1];
    if (prev && runKey(prev.run) === key && prev.color === color && prev.font === font) {
      prev.text += normalized;
      prev.width += measured;
    } else {
      segments.push({ text: normalized, run, font, fontSizePx, color, underline, width: measured });
    }
    width += measured;
  };

  const wrapTextChunk = (chunk: string, run: ShapeTextRun) => {
    if (!opts.wrap || maxWidth <= 0) {
      appendText(chunk, run);
      return;
    }

    const tokens = chunk.split(/(\s+)/);
    for (const token of tokens) {
      if (token === "") continue;

      const { font } = shapeRunFont(run, scale);
      ctx.font = font;
      const normalizedToken = normalizeText(token);
      const tokenWidth = ctx.measureText(normalizedToken).width;
      if (segments.length > 0 && width + tokenWidth > maxWidth) {
        flush();
        if (token.trim() === "") continue;
      }

      appendText(normalizedToken, run, true);
    }
  };

  for (const run of runs) {
    const pieces = String(run.text ?? "").split("\n");
    for (let i = 0; i < pieces.length; i += 1) {
      const chunk = pieces[i] ?? "";
      wrapTextChunk(chunk, run);
      if (i !== pieces.length - 1) {
        flush();
      }
    }
  }

  if (segments.length > 0 || lines.length === 0) {
    flush();
  }

  return lines;
}

function renderShapeText(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  layout: ShapeTextLayout,
  opts: { defaultColor: string; zoom: number },
  cache?: ShapeTextCacheEntry,
): void {
  const scale = Number.isFinite(opts.zoom) && opts.zoom > 0 ? opts.zoom : 1;
  const defaultInsetPx = 4 * scale;
  const insetLeftPx =
    typeof layout.insetLeftEmu === "number" && Number.isFinite(layout.insetLeftEmu)
      ? emuToPx(layout.insetLeftEmu, scale)
      : defaultInsetPx;
  const insetTopPx =
    typeof layout.insetTopEmu === "number" && Number.isFinite(layout.insetTopEmu) ? emuToPx(layout.insetTopEmu, scale) : defaultInsetPx;
  const insetRightPx =
    typeof layout.insetRightEmu === "number" && Number.isFinite(layout.insetRightEmu)
      ? emuToPx(layout.insetRightEmu, scale)
      : defaultInsetPx;
  const insetBottomPx =
    typeof layout.insetBottomEmu === "number" && Number.isFinite(layout.insetBottomEmu)
      ? emuToPx(layout.insetBottomEmu, scale)
      : defaultInsetPx;
  const inner = SHAPE_TEXT_INNER_RECT_SCRATCH;
  inner.x = bounds.x + insetLeftPx;
  inner.y = bounds.y + insetTopPx;
  inner.width = Math.max(0, bounds.width - insetLeftPx - insetRightPx);
  inner.height = Math.max(0, bounds.height - insetTopPx - insetBottomPx);
  if (inner.width <= 0 || inner.height <= 0) return;
  if (layout.textRuns.length === 0) return;

  const wrap = layout.wrap ?? true;
  let lines = cache?.lines;
  const cacheValid =
    cache != null &&
    lines != null &&
    cache.linesMaxWidth === inner.width &&
    cache.linesWrap === wrap &&
    cache.linesZoom === scale &&
    cache.linesDefaultColor === opts.defaultColor;
  if (!cacheValid) {
    const layoutOpts = SHAPE_TEXT_LAYOUT_OPTS_SCRATCH;
    layoutOpts.wrap = wrap;
    layoutOpts.defaultColor = opts.defaultColor;
    layoutOpts.zoom = scale;
    lines = layoutShapeTextLines(ctx, layout.textRuns, inner.width, layoutOpts);
    if (cache) {
      cache.lines = lines;
      cache.linesMaxWidth = inner.width;
      cache.linesWrap = wrap;
      cache.linesZoom = scale;
      cache.linesDefaultColor = opts.defaultColor;
    }
  }
  if (lines.length === 0) return;

  let totalHeight = 0;
  for (let i = 0; i < lines.length; i += 1) {
    totalHeight += lines[i]!.height;
  }
  const vertical = layout.vertical ?? "top";
  let y = inner.y;
  if (vertical === "middle") {
    y = inner.y + (inner.height - totalHeight) / 2;
  } else if (vertical === "bottom") {
    y = inner.y + (inner.height - totalHeight);
  }

  ctx.textBaseline = "top";
  ctx.textAlign = "left";
  ctx.globalAlpha = 1;
  // Shapes may have dashed outlines; ensure underline strokes do not inherit the
  // dash pattern from the shape geometry rendering.
  ctx.setLineDash(LINE_DASH_NONE);

  const alignment = layout.alignment ?? "left";
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i]!;
    let x = inner.x;
    if (alignment === "center") {
      x = inner.x + (inner.width - line.width) / 2;
    } else if (alignment === "right") {
      x = inner.x + (inner.width - line.width);
    }

    for (let j = 0; j < line.segments.length; j += 1) {
      const seg = line.segments[j]!;
      if (!seg.text) continue;
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      ctx.fillText(seg.text, x, y);
      if (seg.underline) {
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = Math.max(1, scale);
        const underlineY = y + seg.fontSizePx;
        ctx.beginPath();
        ctx.moveTo(x, underlineY);
        ctx.lineTo(x + seg.width, underlineY);
        ctx.stroke();
      }
      x += seg.width;
      if (x > inner.x + inner.width) break;
    }

    y += line.height;
    if (y > inner.y + inner.height) break;
  }
}
