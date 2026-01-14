import type { Anchor, DrawingObject, DrawingTransform, ImageEntry, ImageStore, Rect } from "./types";
import { ImageBitmapCache } from "./imageBitmapCache";
import { graphicFramePlaceholderLabel, isGraphicFrame, parseShapeRenderSpec, type ShapeRenderSpec } from "./shapeRenderer";
import { parseDrawingMLShapeText, type ShapeTextLayout, type ShapeTextRun } from "./drawingml/shapeText";
import { getResizeHandleCenters, getRotationHandleCenter, RESIZE_HANDLE_SIZE_PX, ROTATION_HANDLE_SIZE_PX } from "./selectionHandles";
import { applyTransformVector, degToRad } from "./transform";
import { DrawingSpatialIndex } from "./spatialIndex";

import { EMU_PER_INCH, PX_PER_INCH, emuToPx, pxToEmu } from "../shared/emu.js";

export { EMU_PER_INCH, PX_PER_INCH, emuToPx, pxToEmu };
export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH;

type CssVarStyle = Pick<CSSStyleDeclaration, "getPropertyValue">;

type OverlayColorTokens = {
  placeholderChartStroke: string;
  placeholderOtherStroke: string;
  placeholderGraphicFrameStroke: string;
  placeholderLabel: string;
  selectionStroke: string;
  selectionHandleFill: string;
};

const DEFAULT_OVERLAY_COLOR_TOKENS: OverlayColorTokens = {
  placeholderChartStroke: "blue",
  placeholderOtherStroke: "cyan",
  placeholderGraphicFrameStroke: "magenta",
  placeholderLabel: "black",
  selectionStroke: "blue",
  selectionHandleFill: "white",
};

function getRootCssStyle(): CssVarStyle | null {
  if (typeof document === "undefined" || typeof getComputedStyle !== "function") return null;
  try {
    return getComputedStyle(document.documentElement);
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
  opts?: { maxDepth?: number; seen?: Set<string> },
): string {
  const maxDepth = opts?.maxDepth ?? 8;
  const read = (name: string) => {
    if (!style) return "";
    const raw = style.getPropertyValue(name);
    return typeof raw === "string" ? raw.trim() : "";
  };

  let current = String(value ?? "").trim();
  let lastFallback: string | null = null;
  const seen = opts?.seen ?? new Set<string>();

  for (let depth = 0; depth < maxDepth; depth += 1) {
    const parsed = parseCssVarFunction(current);
    if (!parsed) return current || fallback;

    const nextName = parsed.name;
    if (parsed.fallback != null) lastFallback = parsed.fallback;

    // Handle cycles and enforce a max indirection depth.
    if (seen.has(nextName)) {
      current = parsed.fallback ?? lastFallback ?? "";
      continue;
    }
    seen.add(nextName);

    const nextValue = read(nextName);
    if (nextValue) {
      current = nextValue;
      continue;
    }

    const fb = parsed.fallback ?? lastFallback;
    if (fb != null) return resolveCssValue(style, fb, fallback, { maxDepth, seen });
    return fallback;
  }

  if (lastFallback != null) return resolveCssValue(style, lastFallback, fallback, { maxDepth, seen });
  return fallback;
}

function resolveCssVarFromStyle(style: CssVarStyle | null, varName: string, fallback: string): string {
  if (!style) return fallback;
  const start = style.getPropertyValue(varName);
  const trimmed = typeof start === "string" ? start.trim() : "";
  if (!trimmed) return fallback;
  return resolveCssValue(style, trimmed, fallback, { seen: new Set([varName]) });
}

function resolveOverlayColorTokens(style: CssVarStyle | null): OverlayColorTokens {
  return {
    placeholderChartStroke: resolveCssVarFromStyle(style, "--chart-series-1", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderChartStroke),
    placeholderOtherStroke: resolveCssVarFromStyle(style, "--chart-series-2", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderOtherStroke),
    placeholderGraphicFrameStroke: resolveCssVarFromStyle(
      style,
      "--chart-series-3",
      DEFAULT_OVERLAY_COLOR_TOKENS.placeholderGraphicFrameStroke,
    ),
    placeholderLabel: resolveCssVarFromStyle(style, "--text-primary", DEFAULT_OVERLAY_COLOR_TOKENS.placeholderLabel),
    selectionStroke: resolveCssVarFromStyle(style, "--selection-border", DEFAULT_OVERLAY_COLOR_TOKENS.selectionStroke),
    selectionHandleFill: resolveCssVarFromStyle(style, "--bg-primary", DEFAULT_OVERLAY_COLOR_TOKENS.selectionHandleFill),
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
  destroy?(): void;
}

export function anchorToRectPx(anchor: Anchor, geom: GridGeometry, zoom: number = 1): Rect {
  const scale = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  switch (anchor.type) {
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        x: origin.x + emuToPx(anchor.from.offset.xEmu) * scale,
        y: origin.y + emuToPx(anchor.from.offset.yEmu) * scale,
        width: emuToPx(anchor.size.cx) * scale,
        height: emuToPx(anchor.size.cy) * scale,
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);

      const x1 = fromOrigin.x + emuToPx(anchor.from.offset.xEmu) * scale;
      const y1 = fromOrigin.y + emuToPx(anchor.from.offset.yEmu) * scale;

      // In DrawingML, `to` specifies the cell *containing* the bottom-right
      // corner (i.e. the first cell strictly outside the shape when the corner
      // lies on a grid boundary). The absolute end point is therefore the
      // origin of the `to` cell plus the offsets.
      const x2 = toOrigin.x + emuToPx(anchor.to.offset.xEmu) * scale;
      const y2 = toOrigin.y + emuToPx(anchor.to.offset.yEmu) * scale;

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
      const origin = geom.cellOriginPx({ row: 0, col: 0 });
      return {
        x: origin.x + emuToPx(anchor.pos.xEmu) * scale,
        y: origin.y + emuToPx(anchor.pos.yEmu) * scale,
        width: emuToPx(anchor.size.cx) * scale,
        height: emuToPx(anchor.size.cy) * scale,
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
  private readonly shapeTextCache = new Map<number, { rawXml: string; parsed: ShapeTextLayout | null }>();
  private readonly spatialIndex = new DrawingSpatialIndex();
  private selectedId: number | null = null;
  private renderSeq = 0;
  private renderAbort: AbortController | null = null;
  private preloadAbort: AbortController | null = typeof AbortController !== "undefined" ? new AbortController() : null;
  private preloadCount = 0;
  private destroyed = false;
  private cssVarStyle: CssVarStyle | null | undefined = undefined;
  private colorTokens: OverlayColorTokens | null = null;
  private orderedObjects: DrawingObject[] = [];
  private orderedObjectsSource: DrawingObject[] | null = null;
  private themeObserver: MutationObserver | null = null;
  private lastRenderArgs: { objects: DrawingObject[]; viewport: Viewport; options?: { drawObjects?: boolean } } | null = null;
  private readonly pendingImageHydrations = new Map<string, Promise<ImageEntry | undefined>>();
  private hydrationRerenderScheduled = false;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly images: ImageStore,
    private readonly geom: GridGeometry,
    private readonly chartRenderer?: ChartRenderer,
  ) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("overlay canvas 2d context not available");
    this.ctx = ctx;
    this.installThemeObserver();
  }

  private installThemeObserver(): void {
    if (typeof document === "undefined") return;
    const root = document.documentElement;
    if (!root) return;
    if (typeof MutationObserver !== "function") return;

    try {
      this.themeObserver?.disconnect();
      this.themeObserver = new MutationObserver(() => {
        // Defer recomputing CSS vars until the next render; we only need to
        // invalidate cached tokens here.
        this.refreshThemeTokens();
      });
      this.themeObserver.observe(root, { attributes: true, attributeFilter: ["data-theme"] });
    } catch {
      this.themeObserver = null;
    }
  }

  resize(viewport: Viewport): void {
    this.canvas.width = Math.floor(viewport.width * viewport.dpr);
    this.canvas.height = Math.floor(viewport.height * viewport.dpr);
    this.canvas.style.width = `${viewport.width}px`;
    this.canvas.style.height = `${viewport.height}px`;
    this.ctx.setTransform(viewport.dpr, 0, 0, viewport.dpr, 0, 0);
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
    this.cssVarStyle = getRootCssStyle();
    return this.cssVarStyle;
  }

  private getOrderedObjects(objects: DrawingObject[]): DrawingObject[] {
    if (this.orderedObjectsSource === objects) return this.orderedObjects;
    this.orderedObjectsSource = objects;

    let sorted = true;
    for (let i = 1; i < objects.length; i += 1) {
      if (objects[i - 1]!.zOrder > objects[i]!.zOrder) {
        sorted = false;
        break;
      }
    }

    this.orderedObjects = sorted ? objects : [...objects].sort((a, b) => a.zOrder - b.zOrder);
    return this.orderedObjects;
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
      const last = this.lastRenderArgs;
      if (!last) return;
      void this.render(last.objects, last.viewport, last.options).catch(() => {});
    });
  }

  private hydrateImage(imageId: string): void {
    const id = String(imageId ?? "");
    if (!id) return;
    if (typeof this.images.getAsync !== "function") return;
    if (this.pendingImageHydrations.has(id)) return;

    const promise = Promise.resolve()
      .then(() => this.images.getAsync!(id))
      .catch(() => undefined);

    this.pendingImageHydrations.set(id, promise);

    void promise.then((entry) => {
      this.pendingImageHydrations.delete(id);
      if (!entry) return;

      // Ensure subsequent sync `get()` calls can resolve without awaiting `getAsync`.
      try {
        if (!this.images.get(id)) {
          this.images.set(entry);
        }
      } catch {
        // Best-effort: ignore caching failures.
      }

      this.scheduleHydrationRerender();
    });
  }

  async render(objects: DrawingObject[], viewport: Viewport, options?: { drawObjects?: boolean }): Promise<void> {
    this.renderSeq += 1;
    const seq = this.renderSeq;
    let completed = false;
    let shapeCount = 0;
    const drawObjects = options?.drawObjects !== false;
    // Keep the latest render args around so async image hydration can trigger a follow-up render
    // once bytes are available (without relying on callers to poll/refresh).
    this.lastRenderArgs = { objects, viewport: { ...viewport }, options };

    // Cancel any prior render pass so we don't draw stale content after a newer
    // render begins (e.g. rapid scroll/zoom updates). This also lets callers
    // abort in-flight image decode awaits for offscreen images.
    this.renderAbort?.abort();
    const abort = typeof AbortController !== "undefined" ? new AbortController() : null;
    this.renderAbort = abort;
    const signal = abort?.signal;

    try {
    const ctx = this.ctx;
    ctx.clearRect(0, 0, viewport.width, viewport.height);

    const cssVarStyle = this.getCssVarStyle();
    const colors = this.colorTokens ?? (this.colorTokens = resolveOverlayColorTokens(cssVarStyle));
    const zoom = viewport.zoom ?? 1;

    const paneLayout = resolvePaneLayout(viewport, this.geom);
    const viewportRect = { x: 0, y: 0, width: viewport.width, height: viewport.height };
    const prefetchedImageBitmaps = new Map<string, Promise<ImageBitmap>>();

    // Spatial index: compute a small candidate list for the current viewport rather
    // than scanning every drawing on each render.
    this.spatialIndex.rebuild(objects, this.geom, zoom);
    const ordered: DrawingObject[] = [];
    const frozenContentWidth = paneLayout.quadrants.topLeft.width;
    const frozenContentHeight = paneLayout.quadrants.topLeft.height;
    const addCandidates = (quadrant: PaneQuadrant, rect: { x: number; y: number; width: number; height: number }) => {
      if (!(rect.width > 0 && rect.height > 0)) return;
      const candidates = this.spatialIndex.query(rect);
      if (candidates.length === 0) return;
      for (const obj of candidates) {
        const pane = resolveAnchorPane(obj.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
        if (pane.quadrant !== quadrant) continue;
        ordered.push(obj);
      }
    };

    addCandidates("topLeft", {
      x: 0,
      y: 0,
      width: paneLayout.quadrants.topLeft.width,
      height: paneLayout.quadrants.topLeft.height,
    });
    addCandidates("topRight", {
      x: viewport.scrollX + frozenContentWidth,
      y: 0,
      width: paneLayout.quadrants.topRight.width,
      height: paneLayout.quadrants.topRight.height,
    });
    addCandidates("bottomLeft", {
      x: 0,
      y: viewport.scrollY + frozenContentHeight,
      width: paneLayout.quadrants.bottomLeft.width,
      height: paneLayout.quadrants.bottomLeft.height,
    });
    addCandidates("bottomRight", {
      x: viewport.scrollX + frozenContentWidth,
      y: viewport.scrollY + frozenContentHeight,
      width: paneLayout.quadrants.bottomRight.width,
      height: paneLayout.quadrants.bottomRight.height,
    });

    const selectedId = this.selectedId;
    let selectedScreenRect: Rect | null = null;
    let selectedClipRect: Rect | null = null;
    let selectedAabb: Rect | null = null;
    let selectedTransform: DrawingTransform | undefined = undefined;

    if (drawObjects) {
      // First pass: kick off image decodes for all visible images without awaiting so
      // multiple images can decode concurrently on cold render.
      for (const obj of ordered) {
        if (obj.kind.type !== "image") continue;
        if (seq !== this.renderSeq || signal?.aborted) return;

        const rect = this.spatialIndex.getRect(obj.id) ?? anchorToRectPx(obj.anchor, this.geom, zoom);
        const pane = resolveAnchorPane(obj.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
        const scrollX = pane.inFrozenCols ? 0 : viewport.scrollX;
        const scrollY = pane.inFrozenRows ? 0 : viewport.scrollY;
        const screenRect = {
          x: rect.x - scrollX + paneLayout.headerOffsetX,
          y: rect.y - scrollY + paneLayout.headerOffsetY,
          width: rect.width,
          height: rect.height,
        };
        const clipRect = paneLayout.quadrants[pane.quadrant];
        const aabb = getAabbForObject(screenRect, obj.transform);

        if (clipRect.width <= 0 || clipRect.height <= 0) continue;
        if (!intersects(aabb, clipRect)) continue;
        if (!intersects(clipRect, viewportRect)) continue;
        if (!intersects(aabb, viewportRect)) continue;

        const imageId = obj.kind.imageId;
        if (prefetchedImageBitmaps.has(imageId)) continue;
        const entry = this.images.get(imageId);
        if (!entry) continue;
        const bitmapPromise = this.bitmapCache.get(entry, signal ? { signal } : undefined);
        // Attach a no-op rejection handler immediately so failures for images later in the
        // z-order (or in cancelled render passes) don't surface as unhandled promise
        // rejections before we reach their draw pass.
        void bitmapPromise.catch(() => {});
        prefetchedImageBitmaps.set(imageId, bitmapPromise);
      }

      for (const obj of ordered) {
        if (seq !== this.renderSeq || signal?.aborted) return;
        if (obj.kind.type === "shape") shapeCount += 1;
        const rect = this.spatialIndex.getRect(obj.id) ?? anchorToRectPx(obj.anchor, this.geom, zoom);
        const pane = resolveAnchorPane(obj.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
        const scrollX = pane.inFrozenCols ? 0 : viewport.scrollX;
        const scrollY = pane.inFrozenRows ? 0 : viewport.scrollY;
        const screenRect = {
          x: rect.x - scrollX + paneLayout.headerOffsetX,
          y: rect.y - scrollY + paneLayout.headerOffsetY,
          width: rect.width,
          height: rect.height,
        };
        const clipRect = paneLayout.quadrants[pane.quadrant];
        const aabb = getAabbForObject(screenRect, obj.transform);

        if (selectedId != null && obj.id === selectedId) {
          selectedScreenRect = screenRect;
          selectedClipRect = clipRect;
          selectedAabb = aabb;
          selectedTransform = obj.transform;
        }

        if (clipRect.width <= 0 || clipRect.height <= 0) continue;
        // Skip objects that are fully outside of their pane quadrant.
        if (!intersects(aabb, clipRect)) continue;
        // Paranoia: clip rects are expected to be within the viewport, but keep the
        // early-out for callers providing custom layouts.
        if (!intersects(clipRect, viewportRect)) continue;

        const withClip = (fn: () => void) => {
          ctx.save();
          ctx.beginPath();
          ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
          ctx.clip();
          try {
            fn();
          } finally {
            ctx.restore();
          }
        };

        if (!intersects(aabb, viewportRect)) {
          continue;
        }

        if (obj.kind.type === "image") {
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
            withClip(() => {
              ctx.save();
              ctx.strokeStyle = colors.placeholderOtherStroke;
              ctx.lineWidth = 1;
              ctx.setLineDash([4, 2]);
              if (hasNonIdentityTransform(obj.transform)) {
                drawTransformedRect(ctx, screenRect, obj.transform!);
              } else {
                ctx.strokeRect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
              }
              ctx.setLineDash([]);
              ctx.fillStyle = colors.placeholderLabel;
              ctx.globalAlpha = 0.6;
              ctx.font = "12px sans-serif";
              ctx.fillText("missing image", screenRect.x + 4, screenRect.y + 14);
              ctx.restore();
            });
            continue;
          }

          try {
            const bitmapPromise =
              prefetchedImageBitmaps.get(imageId) ??
              this.bitmapCache.get(entry, signal ? { signal } : undefined);
            const bitmap = await bitmapPromise;
            if (signal?.aborted) return;
            if (seq !== this.renderSeq) return;
            withClip(() => {
              if (hasNonIdentityTransform(obj.transform)) {
                withObjectTransform(ctx, screenRect, obj.transform, (localRect) => {
                  ctx.save();
                  try {
                    // Clip to the (possibly rotated/flipped) image bounds so we don't
                    // overdraw neighboring cells when transforms extend beyond the anchor.
                    ctx.beginPath();
                    ctx.rect(localRect.x, localRect.y, localRect.width, localRect.height);
                    ctx.clip();
                    ctx.drawImage(bitmap, localRect.x, localRect.y, localRect.width, localRect.height);
                  } finally {
                    ctx.restore();
                  }
                });
              } else {
                ctx.drawImage(bitmap, screenRect.x, screenRect.y, screenRect.width, screenRect.height);
              }
            });
            continue;
          } catch (err) {
            if (signal?.aborted || isAbortError(err)) return;
            if (seq !== this.renderSeq) return;
            // Fall through to placeholder rendering.
          }
        }

        if (obj.kind.type === "chart") {
          const chartId = obj.kind.chartId;
          if (this.chartRenderer && typeof chartId === "string" && chartId.length > 0) {
            let rendered = false;
            withClip(() => {
              ctx.save();
              try {
                if (hasNonIdentityTransform(obj.transform)) {
                  withObjectTransform(ctx, screenRect, obj.transform, (localRect) => {
                    ctx.save();
                    try {
                      ctx.beginPath();
                      ctx.rect(localRect.x, localRect.y, localRect.width, localRect.height);
                      ctx.clip();
                      this.chartRenderer!.renderToCanvas(ctx, chartId, localRect);
                      rendered = true;
                    } finally {
                      ctx.restore();
                    }
                  });
                } else {
                  ctx.beginPath();
                  ctx.rect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
                  ctx.clip();
                  this.chartRenderer!.renderToCanvas(ctx, chartId, screenRect);
                  rendered = true;
                }
              } catch {
                rendered = false;
              } finally {
                ctx.restore();
              }
            });

            if (rendered) continue;
          }
        }

        if (obj.kind.type === "shape") {
          const rawXml = (obj.kind as any).rawXml ?? (obj.kind as any).raw_xml;
          const rawXmlText = typeof rawXml === "string" ? rawXml : "";

          // Parse `<xdr:txBody>` once and cache; avoid reparsing XML on every frame.
          let cachedText = this.shapeTextCache.get(obj.id);
          if (!cachedText || cachedText.rawXml !== rawXmlText) {
            cachedText = { rawXml: rawXmlText, parsed: parseDrawingMLShapeText(rawXmlText) };
            this.shapeTextCache.set(obj.id, cachedText);
          }
          const textLayout = cachedText.parsed;
          const textParsed = textLayout !== null;
          const hasText = textLayout ? textLayout.textRuns.map((r) => r.text).join("").trim().length > 0 : false;
          const canRenderText = hasText && typeof (ctx as any).measureText === "function";

          let rendered = false;
          let spec: ShapeRenderSpec | null = null;
          try {
            spec = rawXmlText ? parseShapeRenderSpec(rawXmlText) : null;
          } catch {
            spec = null;
          }

          if (spec) {
            const specToDraw = canRenderText ? { ...spec, label: undefined } : spec;
            withClip(() => {
              try {
                withObjectTransform(ctx, screenRect, obj.transform, (localRect) => {
                  drawShape(ctx, localRect, specToDraw, colors, cssVarStyle);
                  if (canRenderText) {
                    renderShapeText(ctx, localRect, textLayout!, { defaultColor: colors.placeholderLabel });
                  }
                });
                rendered = true;
              } catch {
                rendered = false;
              }
            });
            if (rendered) continue;
          }

          // If we couldn't render the shape geometry but we did successfully parse text,
          // still render the text within the anchored bounds (and skip placeholders).
          if (canRenderText) {
            withClip(() => {
              withObjectTransform(ctx, screenRect, obj.transform, (localRect) => {
                ctx.save();
                try {
                  ctx.beginPath();
                  ctx.rect(localRect.x, localRect.y, localRect.width, localRect.height);
                  ctx.clip();
                  renderShapeText(ctx, localRect, textLayout!, { defaultColor: colors.placeholderLabel });
                } finally {
                  ctx.restore();
                }
              });
            });
            continue;
          }

          // Shape parsed but has no text: keep an empty bounds placeholder (no label).
          if (textParsed) {
            withClip(() => {
              ctx.save();
              ctx.strokeStyle = colors.placeholderOtherStroke;
              ctx.lineWidth = 1;
              ctx.setLineDash([4, 2]);
              if (hasNonIdentityTransform(obj.transform)) {
                drawTransformedRect(ctx, screenRect, obj.transform!);
              } else {
                ctx.strokeRect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
              }
              ctx.restore();
            });
            continue;
          }
        }

        // Placeholder rendering for shapes/charts/unknown.
        withClip(() => {
          ctx.save();
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
          ctx.setLineDash([4, 2]);
          if (hasNonIdentityTransform(obj.transform)) {
            drawTransformedRect(ctx, screenRect, obj.transform!);
          } else {
            ctx.strokeRect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
          }
          ctx.setLineDash([]);
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
          ctx.fillText(placeholderLabel, screenRect.x + 4, screenRect.y + 14);
          ctx.restore();
        });
      }
    }

    // Selection overlay.
    if (seq !== this.renderSeq) return;
    if (selectedId != null) {
      if (selectedScreenRect && selectedClipRect && selectedAabb) {
        if (selectedClipRect.width > 0 && selectedClipRect.height > 0 && intersects(selectedAabb, selectedClipRect)) {
          ctx.save();
          ctx.beginPath();
          ctx.rect(selectedClipRect.x, selectedClipRect.y, selectedClipRect.width, selectedClipRect.height);
          ctx.clip();
          drawSelection(ctx, selectedScreenRect, colors, selectedTransform);
          ctx.restore();
        }
      } else {
        // Fallback: selection can still be rendered when `drawObjects` is disabled.
        const selected = ordered.find((o) => o.id === selectedId);
        if (selected) {
          const rect = this.spatialIndex.getRect(selected.id) ?? anchorToRectPx(selected.anchor, this.geom, zoom);
          const pane = resolveAnchorPane(selected.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
          const scrollX = pane.inFrozenCols ? 0 : viewport.scrollX;
          const scrollY = pane.inFrozenRows ? 0 : viewport.scrollY;
          const screen = {
            x: rect.x - scrollX + paneLayout.headerOffsetX,
            y: rect.y - scrollY + paneLayout.headerOffsetY,
            width: rect.width,
            height: rect.height,
          };
          const clipRect = paneLayout.quadrants[pane.quadrant];
          const selectionAabb = getAabbForObject(screen, selected.transform);
          if (clipRect.width > 0 && clipRect.height > 0 && intersects(selectionAabb, clipRect)) {
            ctx.save();
            ctx.beginPath();
            ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
            ctx.clip();
            drawSelection(ctx, screen, colors, selected.transform);
            ctx.restore();
          }
        }
      }
    }
    completed = true;
    } finally {
      // Prune cached shape text layouts for shapes that no longer exist.
      //
      // Only the latest render pass should mutate shared caches; older async renders
      // can finish out-of-order and must not evict newer cache entries.
      if (completed && seq === this.renderSeq && this.shapeTextCache.size > 0) {
        // `shapeTextCache` is keyed by drawing id and can otherwise grow unbounded across
        // delete/undo/redo sessions. We only do the (allocating) prune when it's likely stale.
        const liveShapeCount = drawObjects
          ? shapeCount
          : (() => {
              let count = 0;
              for (const obj of objects) {
                if (obj.kind.type === "shape") count += 1;
              }
              return count;
            })();

        if (this.shapeTextCache.size > liveShapeCount) {
          const liveShapeIds = new Set<number>();
          for (const obj of objects) {
            if (obj.kind.type === "shape") liveShapeIds.add(obj.id);
          }
          for (const id of this.shapeTextCache.keys()) {
            if (!liveShapeIds.has(id)) this.shapeTextCache.delete(id);
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
    // If the bitmap bytes change while we are mid-render or mid-preload, the old decode result can
    // arrive after the cache entry has been invalidated. Abort any in-flight consumers first so the
    // stale ImageBitmap is closed when the decode eventually resolves.
    this.renderAbort?.abort();
    this.renderAbort = null;
    if (this.preloadCount > 0) {
      this.preloadAbort?.abort();
      this.preloadAbort = typeof AbortController !== "undefined" ? new AbortController() : null;
    }
    this.bitmapCache.invalidate(String(imageId ?? ""));
  }

  /**
   * Clear all cached decoded bitmaps.
   *
   * Useful after loading a new workbook snapshot where all image ids/bytes may change.
   */
  clearImageCache(): void {
    // When callers clear the cache (e.g. applying a new document snapshot), ensure any in-flight
    // decodes from older renders/preloads don't leak their ImageBitmaps after the cache entry is
    // dropped.
    this.renderAbort?.abort();
    this.renderAbort = null;
    this.preloadAbort?.abort();
    this.preloadAbort = typeof AbortController !== "undefined" ? new AbortController() : null;
    this.bitmapCache.clear();
  }

  destroy(): void {
    // Cancel any in-flight render and release cached bitmap resources.
    this.destroyed = true;
    this.renderAbort?.abort();
    this.renderAbort = null;
    this.preloadAbort?.abort();
    this.preloadAbort = null;
    this.renderSeq += 1;
    this.chartRenderer?.destroy?.();
    this.themeObserver?.disconnect();
    this.themeObserver = null;
    this.bitmapCache.clear();
    this.shapeTextCache.clear();
    this.cssVarStyle = undefined;
    this.colorTokens = null;
    this.orderedObjects = [];
    this.orderedObjectsSource = null;
    this.selectedId = null;
    this.lastRenderArgs = null;
    this.pendingImageHydrations.clear();
    this.hydrationRerenderScheduled = false;
  }

  /**
   * Alias for `destroy()` (matches other UI controller teardown semantics).
   */
  dispose(): void {
    this.destroy();
  }
}

function isAbortError(err: unknown): boolean {
  return typeof (err as { name?: unknown } | null)?.name === "string" && (err as any).name === "AbortError";
}

function intersects(a: Rect, b: Rect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): boolean {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

type CornerPoint = { x: number; y: number };

function getTransformedCorners(rect: Rect, transform: DrawingTransform): CornerPoint[] {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;
  const local = [
    applyTransformVector(-hw, -hh, transform),
    applyTransformVector(hw, -hh, transform),
    applyTransformVector(hw, hh, transform),
    applyTransformVector(-hw, hh, transform),
  ];
  return local.map((p) => ({ x: cx + p.x, y: cy + p.y }));
}

function getAabbForObject(rect: Rect, transform: DrawingTransform | undefined): Rect {
  if (!hasNonIdentityTransform(transform)) return rect;
  const corners = getTransformedCorners(rect, transform!);
  let minX = corners[0]!.x;
  let maxX = corners[0]!.x;
  let minY = corners[0]!.y;
  let maxY = corners[0]!.y;
  for (let i = 1; i < corners.length; i += 1) {
    const p = corners[i]!;
    if (p.x < minX) minX = p.x;
    if (p.x > maxX) maxX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.y > maxY) maxY = p.y;
  }
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

function withObjectTransform(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  transform: DrawingTransform | undefined,
  fn: (localRect: Rect) => void,
): void {
  if (!hasNonIdentityTransform(transform)) {
    fn(rect);
    return;
  }

  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;

  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(degToRad(transform!.rotationDeg));
  ctx.scale(transform!.flipH ? -1 : 1, transform!.flipV ? -1 : 1);
  try {
    fn({ x: -rect.width / 2, y: -rect.height / 2, width: rect.width, height: rect.height });
  } finally {
    ctx.restore();
  }
}

function drawTransformedRect(ctx: CanvasRenderingContext2D, rect: Rect, transform: DrawingTransform): void {
  const corners = getTransformedCorners(rect, transform);
  ctx.beginPath();
  ctx.moveTo(corners[0]!.x, corners[0]!.y);
  for (let i = 1; i < corners.length; i += 1) {
    ctx.lineTo(corners[i]!.x, corners[i]!.y);
  }
  ctx.closePath();
  ctx.stroke();
}

function drawSelection(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  colors: OverlayColorTokens,
  transform?: DrawingTransform,
): void {
  ctx.save();
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  if (hasNonIdentityTransform(transform)) {
    drawTransformedRect(ctx, rect, transform!);
  } else {
    ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
  }

  const handle = RESIZE_HANDLE_SIZE_PX;
  const half = handle / 2;
  const points = getResizeHandleCenters(rect, transform);

  ctx.fillStyle = colors.selectionHandleFill;
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 1;
  for (const p of points) {
    ctx.beginPath();
    ctx.rect(p.x - half, p.y - half, handle, handle);
    ctx.fill();
    ctx.stroke();
  }

  // Optional Excel-style rotation handle.
  const rotHandle = ROTATION_HANDLE_SIZE_PX;
  const rotHalf = rotHandle / 2;
  const rot = getRotationHandleCenter(rect, transform);
  ctx.beginPath();
  ctx.rect(rot.x - rotHalf, rot.y - rotHalf, rotHandle, rotHandle);
  ctx.fill();
  ctx.stroke();
  ctx.restore();
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  spec: ShapeRenderSpec,
  colors: OverlayColorTokens,
  cssVarStyle: CssVarStyle | null,
): void {
  // Clip to the anchored bounds; this matches the chart rendering behaviour and
  // avoids accidental overdraw if we misinterpret a shape transform.
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();

  const strokeWidthPx = spec.stroke ? emuToPx(spec.stroke.widthEmu) : 0;
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
    ctx.setLineDash(dashPatternForPreset(spec.stroke.dashPreset, strokeWidthPx));
    ctx.stroke();
  }

  if (spec.label) {
    ctx.fillStyle = spec.labelColor ?? colors.placeholderLabel;
    ctx.globalAlpha = 0.8;
    const size = typeof spec.labelFontSizePx === "number" && Number.isFinite(spec.labelFontSizePx) ? spec.labelFontSizePx : 12;
    const family = spec.labelFontFamily?.trim() ? spec.labelFontFamily : "sans-serif";
    const weight = spec.labelBold ? "bold " : "";
    ctx.font = `${weight}${size}px ${family}`;
    const align = spec.labelAlign ?? "left";
    const vAlign = spec.labelVAlign ?? "top";
    ctx.textAlign = align;
    ctx.textBaseline = vAlign === "middle" ? "middle" : vAlign;
    const padding = 4;
    const xText =
      align === "center" ? rect.x + rect.width / 2 : align === "right" ? rect.x + rect.width - padding : rect.x + padding;
    const yText =
      vAlign === "middle" ? rect.y + rect.height / 2 : vAlign === "bottom" ? rect.y + rect.height - padding : rect.y + padding;
    ctx.fillText(spec.label, xText, yText);
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

function dashPatternForPreset(preset: string | undefined, strokeWidthPx: number): number[] {
  if (!preset) return [];
  const unit = Math.max(1, strokeWidthPx);
  switch (preset) {
    case "dash":
    case "sysDash":
      return [4 * unit, 2 * unit];
    case "dot":
    case "sysDot":
      return [unit, 2 * unit];
    case "dashDot":
    case "sysDashDot":
      return [4 * unit, 2 * unit, unit, 2 * unit];
    case "lgDash":
      return [8 * unit, 3 * unit];
    case "lgDashDot":
      return [8 * unit, 3 * unit, unit, 3 * unit];
    case "lgDashDotDot":
    case "sysDashDotDot":
      return [8 * unit, 3 * unit, unit, 3 * unit, unit, 3 * unit];
    case "solid":
    default:
      return [];
  }
}

type PaneQuadrant = "topLeft" | "topRight" | "bottomLeft" | "bottomRight";

function resolveAnchorPane(
  anchor: Anchor,
  frozenRows: number,
  frozenCols: number,
): { quadrant: PaneQuadrant; inFrozenRows: boolean; inFrozenCols: boolean } {
  if (anchor.type === "absolute") {
    return { quadrant: "bottomRight", inFrozenRows: false, inFrozenCols: false };
  }
  const fromRow = anchor.from.cell.row;
  const fromCol = anchor.from.cell.col;
  const inFrozenRows = fromRow < frozenRows;
  const inFrozenCols = fromCol < frozenCols;

  if (inFrozenRows && inFrozenCols) return { quadrant: "topLeft", inFrozenRows, inFrozenCols };
  if (inFrozenRows && !inFrozenCols) return { quadrant: "topRight", inFrozenRows, inFrozenCols };
  if (!inFrozenRows && inFrozenCols) return { quadrant: "bottomLeft", inFrozenRows, inFrozenCols };
  return { quadrant: "bottomRight", inFrozenRows, inFrozenCols };
}

function resolvePaneLayout(
  viewport: Viewport,
  geom: GridGeometry,
): {
  frozenRows: number;
  frozenCols: number;
  headerOffsetX: number;
  headerOffsetY: number;
  quadrants: Record<PaneQuadrant, Rect>;
} {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));
  const cellAreaWidth = Math.max(0, viewport.width - headerOffsetX);
  const cellAreaHeight = Math.max(0, viewport.height - headerOffsetY);

  // `frozenWidthPx/HeightPx` are specified in viewport coordinates (they represent
  // the frozen boundary position). When omitted, derive them from the grid geometry
  // (sheet-space frozen extents) plus any header offset.
  const derivedFrozenContentWidth = (() => {
    if (frozenCols <= 0) return 0;
    try {
      return geom.cellOriginPx({ row: 0, col: frozenCols }).x;
    } catch {
      return 0;
    }
  })();
  const derivedFrozenContentHeight = (() => {
    if (frozenRows <= 0) return 0;
    try {
      return geom.cellOriginPx({ row: frozenRows, col: 0 }).y;
    } catch {
      return 0;
    }
  })();

  // Only treat `frozenWidthPx/HeightPx` as meaningful when the corresponding frozen row/col
  // count is non-zero. This mirrors hit testing semantics and guards against stale pixel
  // extents if a caller updates frozen row/col counts but forgets to reset the boundary.
  const frozenBoundaryX = clamp(
    frozenCols > 0
      ? Number.isFinite(viewport.frozenWidthPx)
        ? viewport.frozenWidthPx!
        : headerOffsetX + derivedFrozenContentWidth
      : headerOffsetX,
    headerOffsetX,
    viewport.width,
  );
  const frozenBoundaryY = clamp(
    frozenRows > 0
      ? Number.isFinite(viewport.frozenHeightPx)
        ? viewport.frozenHeightPx!
        : headerOffsetY + derivedFrozenContentHeight
      : headerOffsetY,
    headerOffsetY,
    viewport.height,
  );

  const frozenContentWidth = clamp(frozenBoundaryX - headerOffsetX, 0, cellAreaWidth);
  const frozenContentHeight = clamp(frozenBoundaryY - headerOffsetY, 0, cellAreaHeight);
  const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
  const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

  const x0 = headerOffsetX;
  const y0 = headerOffsetY;
  const x1 = headerOffsetX + frozenContentWidth;
  const y1 = headerOffsetY + frozenContentHeight;

  return {
    frozenRows,
    frozenCols,
    headerOffsetX,
    headerOffsetY,
    quadrants: {
      topLeft: { x: x0, y: y0, width: frozenContentWidth, height: frozenContentHeight },
      topRight: { x: x1, y: y0, width: scrollableWidth, height: frozenContentHeight },
      bottomLeft: { x: x0, y: y1, width: frozenContentWidth, height: scrollableHeight },
      bottomRight: { x: x1, y: y1, width: scrollableWidth, height: scrollableHeight },
    },
  };
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

function shapeRunFont(run: ShapeTextRun): { font: string; fontSizePx: number } {
  const fontSizePt = run.fontSizePt ?? DEFAULT_SHAPE_FONT_SIZE_PT;
  const fontSizePx = Math.max(1, fontSizePt * PX_PER_PT);
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
  opts: { wrap: boolean; defaultColor: string },
): RenderedLine[] {
  const lines: RenderedLine[] = [];
  let segments: RenderedSegment[] = [];
  let width = 0;
  let maxFontSizePx = 0;

  const flush = () => {
    if (segments.length === 0) {
      lines.push({ segments: [], width: 0, height: DEFAULT_SHAPE_FONT_SIZE_PT * PX_PER_PT * 1.2 });
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

  const appendText = (text: string, run: ShapeTextRun) => {
    if (text === "") return;
    const { font, fontSizePx } = shapeRunFont(run);
    ctx.font = font;
    const measured = ctx.measureText(text).width;
    const color = run.color ?? opts.defaultColor;
    const underline = Boolean(run.underline);

    maxFontSizePx = Math.max(maxFontSizePx, fontSizePx);

    const key = runKey(run);
    const prev = segments[segments.length - 1];
    if (prev && runKey(prev.run) === key && prev.color === color && prev.font === font) {
      prev.text += text;
      prev.width += measured;
    } else {
      segments.push({ text, run, font, fontSizePx, color, underline, width: measured });
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
      if (segments.length === 0 && token.trim() === "") continue;

      const { font } = shapeRunFont(run);
      ctx.font = font;
      const tokenWidth = ctx.measureText(token).width;
      if (segments.length > 0 && width + tokenWidth > maxWidth) {
        flush();
        if (token.trim() === "") continue;
      }

      appendText(token, run);
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
  opts: { defaultColor: string },
): void {
  const padding = 4;
  const inner = {
    x: bounds.x + padding,
    y: bounds.y + padding,
    width: Math.max(0, bounds.width - padding * 2),
    height: Math.max(0, bounds.height - padding * 2),
  };
  if (inner.width <= 0 || inner.height <= 0) return;
  if (layout.textRuns.length === 0) return;

  const wrap = layout.wrap ?? true;
  const lines = layoutShapeTextLines(ctx, layout.textRuns, inner.width, { wrap, defaultColor: opts.defaultColor });
  if (lines.length === 0) return;

  const totalHeight = lines.reduce((sum, line) => sum + line.height, 0);
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
  ctx.setLineDash([]);

  const alignment = layout.alignment ?? "left";
  for (const line of lines) {
    let x = inner.x;
    if (alignment === "center") {
      x = inner.x + (inner.width - line.width) / 2;
    } else if (alignment === "right") {
      x = inner.x + (inner.width - line.width);
    }

    for (const seg of line.segments) {
      if (!seg.text) continue;
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      ctx.fillText(seg.text, x, y);
      if (seg.underline) {
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = 1;
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
