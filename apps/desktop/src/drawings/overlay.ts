import type { Anchor, DrawingObject, ImageStore, Rect } from "./types";
import { ImageBitmapCache } from "./imageBitmapCache";
import { resolveCssVar } from "../theme/cssVars.js";
import { graphicFramePlaceholderLabel, isGraphicFrame } from "./shapeRenderer";
import { getResizeHandleCenters, RESIZE_HANDLE_SIZE_PX } from "./selectionHandles";

import { EMU_PER_INCH, PX_PER_INCH, emuToPx, pxToEmu } from "../shared/emu.js";

export { EMU_PER_INCH, PX_PER_INCH, emuToPx, pxToEmu };
export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH;

function resolveOverlayColorTokens(): {
  placeholderChartStroke: string;
  placeholderOtherStroke: string;
  placeholderGraphicFrameStroke: string;
  placeholderLabel: string;
  selectionStroke: string;
  selectionHandleFill: string;
} {
  return {
    placeholderChartStroke: resolveCssVar("--chart-series-1", { fallback: "blue" }),
    placeholderOtherStroke: resolveCssVar("--chart-series-2", { fallback: "cyan" }),
    placeholderGraphicFrameStroke: resolveCssVar("--chart-series-3", { fallback: "magenta" }),
    placeholderLabel: resolveCssVar("--text-primary", { fallback: "black" }),
    selectionStroke: resolveCssVar("--selection-border", { fallback: "blue" }),
    selectionHandleFill: resolveCssVar("--bg-primary", { fallback: "white" })
  };
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
}

export function anchorToRectPx(anchor: Anchor, geom: GridGeometry): Rect {
  switch (anchor.type) {
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        x: origin.x + emuToPx(anchor.from.offset.xEmu),
        y: origin.y + emuToPx(anchor.from.offset.yEmu),
        width: emuToPx(anchor.size.cx),
        height: emuToPx(anchor.size.cy),
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);

      const x1 = fromOrigin.x + emuToPx(anchor.from.offset.xEmu);
      const y1 = fromOrigin.y + emuToPx(anchor.from.offset.yEmu);

      // In DrawingML, `to` specifies the cell *containing* the bottom-right
      // corner (i.e. the first cell strictly outside the shape when the corner
      // lies on a grid boundary). The absolute end point is therefore the
      // origin of the `to` cell plus the offsets.
      const x2 = toOrigin.x + emuToPx(anchor.to.offset.xEmu);
      const y2 = toOrigin.y + emuToPx(anchor.to.offset.yEmu);

      return {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        width: Math.abs(x2 - x1),
        height: Math.abs(y2 - y1),
      };
    }
    case "absolute":
      return {
        x: emuToPx(anchor.pos.xEmu),
        y: emuToPx(anchor.pos.yEmu),
        width: emuToPx(anchor.size.cx),
        height: emuToPx(anchor.size.cy),
      };
  }
}

export class DrawingOverlay {
  private readonly ctx: CanvasRenderingContext2D;
  private readonly bitmapCache = new ImageBitmapCache();
  private selectedId: number | null = null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly images: ImageStore,
    private readonly geom: GridGeometry,
    private readonly chartRenderer?: ChartRenderer,
  ) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("overlay canvas 2d context not available");
    this.ctx = ctx;
  }

  resize(viewport: Viewport): void {
    this.canvas.width = Math.floor(viewport.width * viewport.dpr);
    this.canvas.height = Math.floor(viewport.height * viewport.dpr);
    this.canvas.style.width = `${viewport.width}px`;
    this.canvas.style.height = `${viewport.height}px`;
    this.ctx.setTransform(viewport.dpr, 0, 0, viewport.dpr, 0, 0);
  }

  async render(objects: DrawingObject[], viewport: Viewport): Promise<void> {
    const ctx = this.ctx;
    ctx.clearRect(0, 0, viewport.width, viewport.height);

    const colors = resolveOverlayColorTokens();
    const ordered = [...objects].sort((a, b) => a.zOrder - b.zOrder);

    const paneLayout = resolvePaneLayout(viewport, this.geom);
    const viewportRect = { x: 0, y: 0, width: viewport.width, height: viewport.height };

    for (const obj of ordered) {
      const rect = anchorToRectPx(obj.anchor, this.geom);
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

      if (clipRect.width <= 0 || clipRect.height <= 0) continue;
      // Skip objects that are fully outside of their pane quadrant.
      if (!intersects(screenRect, clipRect)) continue;
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

      if (!intersects(screenRect, viewportRect)) {
        continue;
      }

      if (obj.kind.type === "image") {
        const entry = this.images.get(obj.kind.imageId);
        if (!entry) continue;
        const bitmap = await this.bitmapCache.get(entry);
        withClip(() => {
          ctx.drawImage(bitmap, screenRect.x, screenRect.y, screenRect.width, screenRect.height);
        });
        continue;
      }

      if (obj.kind.type === "chart") {
        const chartId = obj.kind.chartId;
        if (this.chartRenderer && typeof chartId === "string" && chartId.length > 0) {
          let rendered = false;
          withClip(() => {
            ctx.save();
            try {
              ctx.beginPath();
              ctx.rect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
              ctx.clip();
              this.chartRenderer!.renderToCanvas(ctx, chartId, screenRect);
              rendered = true;
            } catch {
              rendered = false;
            } finally {
              ctx.restore();
            }
          });

          if (rendered) continue;
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
        ctx.strokeRect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
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

    // Selection overlay.
    if (this.selectedId != null) {
      const selected = objects.find((o) => o.id === this.selectedId);
      if (selected) {
        const rect = anchorToRectPx(selected.anchor, this.geom);
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
        if (clipRect.width > 0 && clipRect.height > 0 && intersects(screen, clipRect)) {
          ctx.save();
          ctx.beginPath();
          ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
          ctx.clip();
          drawSelection(ctx, screen, colors);
          ctx.restore();
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
}

function intersects(a: Rect, b: Rect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function drawSelection(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  colors: ReturnType<typeof resolveOverlayColorTokens>,
): void {
  ctx.save();
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);

  const handle = RESIZE_HANDLE_SIZE_PX;
  const half = handle / 2;
  const points = getResizeHandleCenters(rect);

  ctx.fillStyle = colors.selectionHandleFill;
  ctx.strokeStyle = colors.selectionStroke;
  ctx.lineWidth = 1;
  for (const p of points) {
    ctx.beginPath();
    ctx.rect(p.x - half, p.y - half, handle, handle);
    ctx.fill();
    ctx.stroke();
  }
  ctx.restore();
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

  const frozenBoundaryX = clamp(
    Number.isFinite(viewport.frozenWidthPx) ? viewport.frozenWidthPx! : headerOffsetX + derivedFrozenContentWidth,
    headerOffsetX,
    viewport.width,
  );
  const frozenBoundaryY = clamp(
    Number.isFinite(viewport.frozenHeightPx) ? viewport.frozenHeightPx! : headerOffsetY + derivedFrozenContentHeight,
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
